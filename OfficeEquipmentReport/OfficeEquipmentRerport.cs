using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.Reporting;

namespace OfficeEquipmentReport
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

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            try
            {
                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    OfficeEquipmentRerport.MakeReport(context, selectionForm.reportParameters);
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
    public class OfficeEquipmentRerport
    {

        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            OfficeEquipmentRerport report = new OfficeEquipmentRerport();
            LogDelegate m_writeToLog;

            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
        {
            var m_writeToLog = new LogDelegate(m_form.writeToLog);
            context.CopyTemplateFile();    // Создаем копию шаблона

            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            Xls xls = new Xls();
            m_form.setStage(IndicationForm.Stages.DataAcquisition);

            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);


            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            m_writeToLog("=== Сортировка ===");

            var sortReportData = SortPositionsAccounting(reportData.OrderBy(t=>t.Type).ToList());

            try
            {
                xls.Application.DisplayAlerts = false;

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeSelectionByTypeReport(xls, sortReportData, reportParameters, m_writeToLog, m_form.progressBar);

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


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№ п/п", col, row, xls);
            col = InsertHeader("ФИО", col, row, xls);          
            col = InsertHeader("Комната", col, row, xls);
            col = InsertHeader("Наименование оргтехники", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);

            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
            return col;
        }

        public Dictionary<string, List<PositionAccounting>> SortPositionsAccounting (List<PositionAccounting> data)
        {
            
            Dictionary<string, List<PositionAccounting>> sortingData = new Dictionary<string, List<PositionAccounting>>();
            string type = data[0].Type;

            List<PositionAccounting> listPosAccounting = new List<PositionAccounting>();
            listPosAccounting.Add(data[0]);

            for (int i = 1; i < data.Count; i++)
            {
                if (data[i].Type == type)
                {
                    listPosAccounting.Add(data[i]);
                }
                else
                {
                    sortingData.Add(type, listPosAccounting);
                    type = data[i].Type;
                    listPosAccounting = new List<PositionAccounting>();
                    listPosAccounting.Add(data[i]);
                }
            }

            sortingData.Add(type, listPosAccounting);

            return sortingData;
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

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            xls[1, row].SetValue("Оргтехника " + reportParameters.Department[Guids.DepartmentParameters.NameGuid].Value);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 5, 1].MergeCells = true;

            row = row + 1;

            return row;
        }


        private void MakeSelectionByTypeReport(Xls xls, Dictionary<string, List<PositionAccounting>> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Sum(t => t.Value.Count));

            if (progressBar.Maximum < reportData.Sum(t=>t.Value.Count))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            double step = 0;

            int row = 1;
            int col = 1;
            row = PasteHeadName(xls, reportParameters, col, row);
            xls[1, row + 1, 1, 6].Select();
            xls.Application.ActiveWindow.FreezePanes = true;

            HeaderTable(xls, reportParameters, col, row);
            int i = 1;
            int beginTypeRow = 0;
            row++;

            foreach (var type in reportData)
            {
                WriteType(xls, beginTypeRow, row, type.Key);
                row++;
                beginTypeRow = row;
                string prewName = string.Empty;
                var countMergeRow = 0;
                var beginMergeRow = 0;
                var endMergeRow = 0;
                foreach (var pos in type.Value.OrderBy(t=>t.User.Name).ThenBy(t => t.Name))
                {
                    step += progressStep;
                    if (step >= 1)
                    {
                        progressBar.PerformStep();
                        step = step - progressBar.Step;
                    }
                    if (prewName == pos.User.Name)
                    {
                        endMergeRow = row;
                        WriteRow(xls, logDelegate, row, i, pos, false, 0);               
                    }
                    else
                    {
                        countMergeRow = endMergeRow - beginMergeRow + 1;

                        if (countMergeRow > 1 && beginMergeRow != 0 && endMergeRow != 0)
                        {
                            xls[2, beginMergeRow, 1, countMergeRow].MergeCells = true;                        
                            endMergeRow = 0;
                        }
 
                        beginMergeRow = row;
                        WriteRow(xls, logDelegate, row, i, pos, true, type.Value.Where(t=>t.User.Name == pos.User.Name && t.Status == 0).ToList().Count);
                    }
                    i++;
                    row++;
                    prewName = pos.User.Name;
                }

                countMergeRow = endMergeRow - beginMergeRow + 1;

                if (countMergeRow > 1 && beginMergeRow != 0 && endMergeRow != 0)
                {
                    xls[2, beginMergeRow, 1, countMergeRow].MergeCells = true;
                }

                WriteSummary(xls, type.Value.Count(t => t.Status == 0), type.Value.Count, beginTypeRow, row, type.Key);
                row++;
            }
            xls.AutoWidth();
        }

        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, PositionAccounting data, bool writeFIO, int countPos)
        {
            logDelegate(String.Format("Запись объекта: " + data.Name + " " + data.User.Name));
            xls[1, row].SetValue(i.ToString());

            if (writeFIO)
            {
                if (data.User.Status == 1 || (data.User.User != null && data.User.User.Class.IsInherit(Guids.GroupAndUser.BlockingUserTypeGuid)))
                {
                    xls[2, row].Interior.Color = Color.Yellow;
                    xls[2, row].SetValue(data.User.Name + " не работает" + ": " + countPos + "шт.");
                    xls[2, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                }
                else
                {
                    xls[2, row].SetValue(data.User.Name + ": " + countPos + "шт");
                    xls[2, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                }
            }

            xls[3, row].SetValue(data.Place.Name + " " + data.Place.Building);

            if (data.Status == 1)
            {
                xls[4, row].Interior.Color = Color.Yellow;
                xls[4, row].SetValue(data.Name + " списано");
            }
            else
            {
                xls[4, row].SetValue(data.Name);
            }

            xls[5, row].SetValue(1);

            xls[1, row, 5, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteSummary(Xls xls, int countAll, int countReal, int beginTypeRow, int row,  string type)
        {
            xls[4, row].SetValue("Итого:");           
            xls[4, row].Font.Bold = true;
            xls[4, row].HorizontalAlignment = XlHAlign.xlHAlignRight;

            if (countAll == countReal)
            {
                xls[5, row].SetFormula("=СУММ(" + xls[5, beginTypeRow].Address + ":" + xls[5, row - 1].Address + ")");
            }
            else
            {
                xls[5, row].SetValue(countReal + "(" + countAll + ")");
            }
            xls[5, row].Font.Bold = true;

            xls[1, row, 5, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }


        private static void WriteType(Xls xls, int beginTypeRow, int row, string type)
        {
            xls[1, row].SetValue(type);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 5, 1].MergeCells = true;

            for (int i = 1; i < 5; i++)
            {
                xls[i, row].Interior.Color = Color.LightGray;
            }


            xls[1, row, 5, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private List<PositionAccounting> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");
            List<PosAccountingForReport> reportData = new List<PosAccountingForReport>();

            var departmentInfo = new Department(reportParameters.Department);

            List<PositionAccounting> listPositionsAccounting = new List<PositionAccounting>();

            foreach (var employee in departmentInfo.Employees)
            {             
                foreach (var pos in employee.PosAccounting)
                {
                    if (pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.ComputerTypeGuid) && reportParameters.Computer ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.MonitorTypeGuid) && reportParameters.Monitor ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.PrinterTypeGuid) && reportParameters.Printer ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.ScannerTypeGuid) && reportParameters.Scanner ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.StorageTypeGuid) && reportParameters.Storage ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.NotebookTypeGuid) && reportParameters.Notebook ||
                        pos.Name.ToLower().Trim().Contains("мфу") && reportParameters.MFU ||
                        pos.RefPosAc.Class.IsInherit(Guids.PositionAccountingTypes.OtherDevicesTypeGuid) && reportParameters.OtherDevices)
                    {
                        listPositionsAccounting.Add(new PositionAccounting(pos, employee));
                    }
                }
            }

            return listPositionsAccounting;
        }
    }
}
