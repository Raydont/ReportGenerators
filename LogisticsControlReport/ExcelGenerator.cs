using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace LogisticsControlReport
{
    public class ExcelGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xls";
            }
        }

        /// <summary>
        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            try
            {
                using (new WaitCursorHelper(false))
                {
                    selectionForm.Init(context);

                    if (selectionForm.ShowDialog() == DialogResult.OK)
                    {
                        var reportClass = new ControlExecuteReport();
                        reportClass.MakeReport(selectionForm.reportParameters);
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
            }
        }
    }

    public delegate void LogDelegate(string line);
    public class ControlExecuteReport
    {
        public IndicationForm m_form = new IndicationForm();
        public void Dispose()
        {
        }

        public void MakeReport(ReportParameters reportParameters)
        {
            ControlExecuteReport report = new ControlExecuteReport();
            LogDelegate m_writeToLog;

            report.Make(reportParameters);
        }

        public void Make(ReportParameters reportParameters)
        {
            var m_writeToLog = new LogDelegate(m_form.writeToLog);
            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            Xls xls = new Xls();
            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            m_writeToLog("=== Обработка данных ===");

            //справочник ОМТС
            var referenceInfo = ReferenceCatalog.FindReference(Guids.References.OMTS);
            var reference = referenceInfo.CreateReference();

            //Формирование фильтра
            var filter = new Filter(referenceInfo);

            //Если выбран исполнитель
            if (reportParameters.SelectedExecuter)
            {
                filter.Terms.AddTerm("[" + Guids.OMTSParameters.Executer + "]", ComparisonOperator.Equal, reportParameters.Executer[Guids.UsersParameters.FullName].Value.ToString());
            }

            List<Stage> stages = new List<Stage>();

            //Добавляем к фильтру условия по дате исполнения
            switch (reportParameters.DateExecute)
            {
                case "только просроченные" :       
                    stages.Add(Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.Open));
                    stages.Add(Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.InWork));
                    stages.Add(Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.ReWork));

                    filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Stage], ComparisonOperator.IsIncludedInGroup, stages);
                    break;
                case "только в работе/на доработке":
                    stages.Add(Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.InWork));
                    stages.Add(Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.ReWork));
                    filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Stage], ComparisonOperator.IsIncludedInGroup, stages);
                    break;
                case "только выполненные":
                    filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Stage], ComparisonOperator.Equal, Stage.Find(ServerGateway.Connection, Guids.StagesOMTS.Close));
                    break;
                default:
                    break;
            }

           // MessageBox.Show(filter.ToString());

            var findedObjects = reference.Find(filter).ToList();

            if (findedObjects.Count == 0)
            {
                MessageBox.Show("Количество объектов для отчета 0. Возможно некорректно сработал фильтр. Обратитесь в Отдел 911. ", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var progressStep = Convert.ToDouble(m_form.progressBar.Maximum) / (findedObjects.Count * 2);

            if (m_form.progressBar.Maximum < (findedObjects.Count * 2))
            {
                m_form.progressBar.Step = 1;
            }
            else
            {
                m_form.progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            if (findedObjects.Count == 0)
            {
                this.Dispose();
                m_form.Close();
            
            }
            var reportData = ProcessingData(reportParameters, findedObjects);

            try
            {
                xls.Application.DisplayAlerts = false;
                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportParameters, reportData, m_writeToLog, m_form.progressBar);
                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                xls.Application.DisplayAlerts = true;
                xls.Visible = true;
            }
            m_form.Dispose();
        }

        public DateTime GetLastDate(ReferenceObject refObject)
        {
            var executeDate = Convert.ToDateTime(refObject[Guids.OMTSParameters.PlanExecuteDate].Value).Date;
            var listChange = refObject.GetObjects(Guids.OMTSParameters.ChangeDateHistoryList); 
            

            if (listChange.Count > 0)
            {
                var lastChange = Convert.ToDateTime(listChange.OrderBy(t => Convert.ToDateTime(t[Guids.OMTSParameters.ChangeDate].Value)).Last()[Guids.OMTSParameters.ChangeDate].Value);
                if (lastChange < executeDate)
                {
                    return lastChange;
                }
            }

            return executeDate;
        }

        public bool FailsInPeriods (DateTime beginDate, DateTime endDate, ReferenceObject refObject)
        {
            var fails = refObject.GetObjects(Guids.OMTSParameters.FailsListObjects).Select(t => Convert.ToDateTime(t[Guids.OMTSParameters.Fail].Value).Date);

            foreach (var fail in fails)
                if (beginDate <= fail && endDate >= fail)            
                    return true;

            return false;
        }


        public List<ReferenceObject> ProcessingData(ReportParameters reportParameters, List<ReferenceObject> reportData)
        {
            List<ReferenceObject> resultData = reportData;

            if (reportParameters.SelectedPeriod)
            {
                resultData = reportData.Where(t => reportParameters.BeginPeriod <= GetLastDate(t) && reportParameters.EndPeriod >= GetLastDate(t)).ToList();
            }

            if (reportParameters.DateExecute == "только просроченные")
            {
                resultData = resultData.Where(t => reportParameters.BeginPeriod <= GetLastDate(t) && reportParameters.EndPeriod >= GetLastDate(t)).ToList();
            }

            if (reportParameters.SelectedPeriodFails)
            {
                resultData = resultData.Where(t => FailsInPeriods(reportParameters.BeginPeriodFails, reportParameters.EndPeriodFails, t)).ToList();
            }

            return resultData;
        }


        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].WrapText = true;

            col++;

            return col;
        }

        public int PasteHeader(Xls xls, int col, int row, ReportParameters reportParameters)
        {
            col = InsertHeader("Запись", col, row, xls);
            col = InsertHeader("№ вх.", col, row, xls);
            col = InsertHeader("Дата вх.", col, row, xls);
            col = InsertHeader("Краткое содержание", col, row, xls);
            col = InsertHeader("Контрагент", col, row, xls);
            col = InsertHeader("Отвественный", col, row, xls);
            col = InsertHeader("Дата передачи в работу", col, row, xls);
            col = InsertHeader("Срок исполнения", col, row, xls);
            col = InsertHeader("Дата исполнения", col, row, xls);
            col = InsertHeader("Просрочено дней", col, row, xls);
            col = InsertHeader("Количество незачетов", col, row, xls);

            if (reportParameters.SelectedPeriodFails)
                col = InsertHeader("Незачеты", col, row, xls);


            xls[1, row, col-1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            return row;
        }



        private void MakeSelectionByTypeReport(Xls xls, ReportParameters reportParameters, List<ReferenceObject> reportData, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (reportData.Count*2);

            if (progressBar.Maximum < (reportData.Count*2))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            double step = 0;

            int errorId = 0;
            int col = 1;
            int row = 1;
            string prewMaker = string.Empty;

            var textHead = "Контроль обработки и исполнения документов ОМТС";

            if (reportParameters.AllTime)
                textHead += " за весь период";
            else
            {
                if (reportParameters.SelectedPeriod)
                    textHead += " за период по срокам исполнения c " + reportParameters.BeginPeriod.ToShortDateString() + " по " + reportParameters.EndPeriod.ToShortDateString();
                else
                {
                    if (reportParameters.SelectedPeriodFails)
                        textHead += " за период по незачетам c " + reportParameters.BeginPeriodFails.ToShortDateString() + " по " + reportParameters.EndPeriodFails.ToShortDateString();
                }
            }        

            if (reportParameters.AllExecuter)
                textHead += " по всем исполнителям";

            if (reportParameters.SelectedExecuter)
                textHead += " по исполнителю " + reportParameters.Executer[Guids.UsersParameters.ShortName].Value.ToString();

            textHead += " по дате исполнения " + reportParameters.DateExecute;

            PasteHeadName(xls, col, row, 12, 11, textHead);
            row++;

            col = 1;
            row = PasteHeader(xls, col, row, reportParameters);
            var beginRow = row;
            foreach (var record in reportData.OrderBy(t => Convert.ToDateTime(t[Guids.OMTSParameters.RegData].Value)))
            {
                logDelegate(String.Format("Запись объекта: " + record[Guids.OMTSParameters.RegNumber].Value + " " + Convert.ToDateTime(record[Guids.OMTSParameters.RegData].Value).ToShortDateString()));

                xls[1, row].NumberFormat = "@";
                xls[1, row].SetValue(record[Guids.OMTSParameters.RegNumber].Value.ToString());

                if (record.GetObject(Guids.OMTSParameters.InDocLink) != null)
                {
                    xls[2, row].SetValue(record.GetObject(Guids.OMTSParameters.InDocLink)[Guids.RKKParametetrs.RegNumber].Value.ToString());
                    xls[3, row].SetValue(Convert.ToDateTime(record.GetObject(Guids.OMTSParameters.InDocLink)[Guids.RKKParametetrs.RegData].Value).ToShortDateString());
                }
                xls[4, row].SetValue(record[Guids.OMTSParameters.Content].Value.ToString());
                xls[4, row].WrapText = true;
                xls[5, row].SetValue(record[Guids.OMTSParameters.Contractor].Value.ToString());
                xls[6, row].SetValue(record[Guids.OMTSParameters.Executer].Value.ToString());
                xls[7, row].SetValue(Convert.ToDateTime(record[Guids.OMTSParameters.RegData].Value).ToShortDateString());
                xls[8, row].SetValue(GetLastDate(record).ToShortDateString());

                if (Convert.ToDateTime(record[Guids.OMTSParameters.ExecuteData].Value) != DateTime.MinValue)
                {
                    xls[9, row].SetValue(Convert.ToDateTime(record[Guids.OMTSParameters.ExecuteData].Value).ToShortDateString());
                    var timeSpan = Convert.ToDateTime(record[Guids.OMTSParameters.ExecuteData].Value).Date - GetLastDate(record).Date;
                    xls[10, row].SetValue(timeSpan.Days.ToString());
                }


                if (reportParameters.SelectedPeriodFails)
                {
                    col = 12;
                    var failsInPeriods = record.GetObjects(Guids.OMTSParameters.FailsListObjects).Where(t=>Convert.ToDateTime(t.ToString()) >= reportParameters.BeginPeriodFails 
                    && Convert.ToDateTime(t.ToString()) <= reportParameters.EndPeriodFails);

                    xls[11, row].SetValue(failsInPeriods.Count().ToString());

                    xls[12, row].SetValue(string.Join("\r\n", record.GetObjects(Guids.OMTSParameters.FailsListObjects).Select(t => Convert.ToDateTime(t.ToString()).ToShortDateString())));
                }
                else
                {
                    xls[11, row].SetValue(record.GetObjects(Guids.OMTSParameters.FailsListObjects).Count.ToString());
                    col = 11;
                }

                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                   XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                row++;
            }
          
            PrintSetting(xls);

            xls[1, 3, col, row].WrapText = true;
            xls.AutoHeight();

        }


        public int PasteHeadName(Xls xls, int col, int row, int sizeFont, int countMerge, string head)
        {
            xls[1, row].SetValue(head);
            xls[1, row].Font.Bold = true;
            xls[col, row].Font.Size = sizeFont;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, countMerge, 1].MergeCells = true;

            row++;

            return row;
        }

        private static void PrintSetting(Xls xls)
        {
            xls.SetColumnWidth(1, 8);
            xls.SetColumnWidth(2, 12);
            xls.SetColumnWidth(3, 12);
            xls.SetColumnWidth(4, 30);
            xls.SetColumnWidth(5, 20);
            xls.SetColumnWidth(6, 20);
            xls.SetColumnWidth(7, 12);
            xls.SetColumnWidth(8, 12);
            xls.SetColumnWidth(9, 12);
            xls.SetColumnWidth(10, 12);
            xls.SetColumnWidth(11, 12);
            xls.SetColumnWidth(12, 12);
        }

   

    }
}
