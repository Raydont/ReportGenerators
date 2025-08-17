using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace RunningContractsReport
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
                using (new WaitCursorHelper(false))
                {
                    ReferenceInfo info = ReferenceCatalog.FindReference(ContractParameters.Users);
                    Reference reference = info.CreateReference();

                    if (!ClientView.Current.GetUser().Parents.Contains(reference.Find(ContractParameters.Department908)) &&
                        !ClientView.Current.GetUser().ToString().Contains("Гудков ") && !ClientView.Current.GetUser().ToString().Contains("Муравьева "))
                    {
                        MessageBox.Show("Сформировать отчет о действующих ДОД и ДЗУ могут только сотрудники отдела 908", "Недоступная операция", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    selectionForm.Init(context);
                    if (selectionForm.ShowDialog() == DialogResult.OK)
                    {
                        RunningContractsReport.MakeReport(context, selectionForm.reportParameters);
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

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);

    public class RunningContractsReport
    {
        public IndicationForm m_form = new IndicationForm();
        public  void ReadDataDOD(List<ReportItem> reportItems, ReferenceObject obj, bool actualContracts, DateTime With, DateTime On)
        {
            var item = new ReportItem();
            if (obj[ContractParameters.DODGuids.NumberOrderClose].Value.ToString().Trim() == string.Empty)
                item.ActualContract = true;
            else
                item.ActualContract = false;

            if (actualContracts && item.ActualContract)
            {
                item.NameContarctor = obj[ContractParameters.DODGuids.Contractor].Value.ToString().Trim();
                item.NumberOrder = obj[ContractParameters.DODGuids.NumberOrder].Value.ToString().Trim();
                item.Date = Convert.ToDateTime(obj[ContractParameters.DODGuids.From].Value).Date;
                item.Number = obj[ContractParameters.DODGuids.RegNumber].Value.ToString().Trim();
                item.Cost = Convert.ToDouble(obj[ContractParameters.DODGuids.CostWithNDS].Value);
                item.CostNDS = Convert.ToDouble(obj[ContractParameters.DODGuids.CostNDS].Value);
                item.Currency = obj[ContractParameters.DODGuids.Currency].Value.ToString().Trim();
                item.Content = obj[ContractParameters.DODGuids.Content].Value.ToString().Trim();
                item.RealNumber = obj[ContractParameters.DODGuids.RealNumber] == null ? string.Empty : obj[ContractParameters.DODGuids.RealNumber].Value.ToString().Trim();
                // item.Term = Convert.ToDateTime(obj[DODGuids.DateClose].Value).Date;

                //получение объектов списка или объектов по связи 1:N, N:N, на любой справочник
                List<ReferenceObject> LinkRkk = obj.GetObjects(ContractParameters.DocumentLinks);
                double advance = 0;


                foreach (var kz in LinkRkk)
                {
                    try
                    {

                        advance += Convert.ToDouble(kz[ContractParameters.DODGuids.AdvancePaymentPlan].Value);
                        item.ListNumberOrder.Add(kz[ContractParameters.DODGuids.NumberOrderKZ].Value.ToString().Trim());
                        item.ListTerm.Add(Convert.ToDateTime(kz[ContractParameters.DODGuids.DateEndKZ].Value).Date);
                    }
                    catch
                    {

                    }

                }

                if (item.Cost != 0)
                    item.AdvancePayment = Math.Round(advance / item.Cost, 2);
                else
                    item.AdvancePayment = 0;

                reportItems.Add(item);
            }
            else
            {
                if (With < Convert.ToDateTime(obj[ContractParameters.DODGuids.From].Value).Date && On > Convert.ToDateTime(obj[ContractParameters.DODGuids.From].Value).Date)
                {
                    item.NameContarctor = obj[ContractParameters.DODGuids.Contractor].Value.ToString().Trim();
                    item.NumberOrder = @obj[ContractParameters.DODGuids.NumberOrder].Value.ToString().Trim();
                    item.Date = Convert.ToDateTime(obj[ContractParameters.DODGuids.From].Value).Date;
                    item.Number = obj[ContractParameters.DODGuids.RegNumber].Value.ToString().Trim();
                    item.Cost = Convert.ToDouble(obj[ContractParameters.DODGuids.CostWithNDS].Value);
                    item.CostNDS = Convert.ToDouble(obj[ContractParameters.DODGuids.CostNDS].Value);
                    item.Currency = obj[ContractParameters.DODGuids.Currency].Value.ToString().Trim();
                    item.Content = obj[ContractParameters.DODGuids.Content].Value.ToString().Trim();
                    item.RealNumber = obj[ContractParameters.DODGuids.RealNumber] == null ? string.Empty : obj[ContractParameters.DODGuids.RealNumber].Value.ToString().Trim();

                    //получение объектов списка или объектов по связи 1:N, N:N, на любой справочник
                    List<ReferenceObject> LinkRkk = obj.GetObjects(ContractParameters.DocumentLinks);
                    double advance = 0;

                    foreach (var kz in LinkRkk)
                    {
                        try
                        {

                            advance += Convert.ToDouble(kz[ContractParameters.DODGuids.AdvancePaymentPlan].Value);
                            item.ListNumberOrder.Add(kz[ContractParameters.DODGuids.NumberOrderKZ].Value.ToString().Trim());
                            item.ListTerm.Add(Convert.ToDateTime(kz[ContractParameters.DODGuids.DateEndKZ].Value).Date);

                        }
                        catch
                        {

                        }
                    }
                    if (item.Cost != 0)
                        item.AdvancePayment = Math.Round(advance / item.Cost, 2);
                    else
                        item.AdvancePayment = 0;

                    reportItems.Add(item);
                }
            }
        }

        public void ReadDataDZU(List<ReportItem> reportItems, ReferenceObject obj, DateTime With, DateTime On)
        {
            var item = new ReportItem();

            if (With < Convert.ToDateTime(obj[ContractParameters.DZUGuids.From].Value).Date && On > Convert.ToDateTime(obj[ContractParameters.DODGuids.From].Value).Date)
            {
                item.NameContarctor = obj[ContractParameters.DZUGuids.Contractor].Value.ToString().Trim();
                item.Date = Convert.ToDateTime(obj[ContractParameters.DZUGuids.From].Value).Date;
                item.Number = obj[ContractParameters.DZUGuids.RegNumber].Value.ToString().Trim();
                item.Cost = Convert.ToDouble(obj[ContractParameters.DZUGuids.CostWithNDS].Value);
                item.CostNDS = Convert.ToDouble(obj[ContractParameters.DZUGuids.CostNDS].Value);
                item.Currency = obj[ContractParameters.DZUGuids.Currency].Value.ToString().Trim();
                item.Content = obj[ContractParameters.DZUGuids.Content].Value.ToString().Trim();
                item.RealNumber = obj[ContractParameters.DZUGuids.RealNumber] == null ? string.Empty : obj[ContractParameters.DZUGuids.RealNumber].Value.ToString();
                reportItems.Add(item);
            }
        }

        public void HeadTableDOD(Xls xls, int row, int col)
        {
            col = InsertHeader("Наименование\nконтрагента", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("Номер\nзаказа", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("№ и дата\nдоговора", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("Стоимость товаров (работ, услуг)", col, row, xls);
            xls[col - 1, row, 4, 1].MergeCells = true;
            col--;
            row++;
            col = InsertHeader("сумма", col, row, xls);
            col = InsertHeader("в т.ч.\nНДС", col, row, xls);
            col = InsertHeader("валюта\nдоговора", col, row, xls);
            col = InsertHeader("величина\nаванса, %", col, row, xls);
            row--;
            col = InsertHeader("Предмет\nдоговора", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("Номер\nзаказа/\nэтапа", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("Срок\nисполнения\nдоговора", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
            xls[1, row, col-1, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight,
                                                                                             XlBordersIndex.xlInsideVertical);
            row++;
        }
        public void HeadTableDZU(Xls xls, int row, int col)
        {
            col = InsertHeader("Наименование\nконтрагента", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("№ и дата\nдоговора", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;
            col = InsertHeader("Стоимость товаров\n(работ, услуг)", col, row, xls);
            xls[col - 1, row, 3, 1].MergeCells = true;
            col--;
            row++;
            col = InsertHeader("сумма", col, row, xls);
            col = InsertHeader("в т.ч.\nНДС", col, row, xls);
            col = InsertHeader("валюта\nдоговора", col, row, xls);
            row--;
            col = InsertHeader("Предмет\nдоговора", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
            xls[1, row, col - 1, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight,
                                                                                             XlBordersIndex.xlInsideVertical);
            row++;
        }

        public int InsertHeader(string header, int col, int row, Xls xls)
        {

            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].Interior.Color = Color.LightGray;
            col++;

            return col;
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
            bool isCatch = false;

            RunningContractsReport report = new RunningContractsReport();

            m_form.setStage(IndicationForm.Stages.DataAcquisition);

            List<ReportItem> reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                WriteReport(xls, reportData, reportParameters, m_writeToLog, m_form.progressBar);

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");

                xls.Application.DisplayAlerts = true;               
                xls.AutoWidth();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                xls.Visible = true;
            }

            m_form.Dispose();
        }

        private List<ReportItem> GetData(ReportParameters reportParameters, LogDelegate m_writeToLog, ProgressBar progressBar)
        {           
            // получение описания справочника          
            var referenceInfo = ReferenceCatalog.FindReference(ContractParameters.RKKReference);

            // создание объекта для работы с данными
            var reference = referenceInfo.CreateReference();
            // получение объектов справочника          

            var reportItems = new List<ReportItem>();

            ClassObject classObject;

            if (reportParameters.DOD)
            {
                classObject = referenceInfo.Classes.Find(ContractParameters.DODType);
            }
            else
            {
                classObject = referenceInfo.Classes.Find(ContractParameters.DZUType);
            }
            //Создаем фильтр
            var filter = new TFlex.DOCs.Model.Search.Filter(referenceInfo);
            //Добавляем условие поиска – «Тип = Длина»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
            ComparisonOperator.Equal, classObject);

            if (reportParameters.DOD && reportParameters.ActualContract)
            {
                filter.Terms.AddTerm("Номер приказа на закрытие", ComparisonOperator.Equal, string.Empty);
            }
            else
            {
                filter.Terms.AddTerm("От", ComparisonOperator.GreaterThan, reportParameters.BeginDate);
                filter.Terms.AddTerm("От", ComparisonOperator.LessThan, reportParameters.EndDate);
            }
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
            var listObj = reference.Find(filter);
            //Задание начальных условий для progressBar
            var progressStep = listObj.Count / 50;
            int step = 0;

            //Считывание данных        
            foreach (var obj in listObj)
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }

                if (reportParameters.DOD)
                {
                    ReadDataDOD(reportItems, obj, reportParameters.ActualContract, reportParameters.BeginDate, reportParameters.EndDate);
                    m_writeToLog(string.Format("Загрузка объекта: {0} {1}", obj[ContractParameters.DODGuids.RegNumber].Value.ToString(), obj[ContractParameters.DODGuids.Contractor].Value.ToString()));
                }
                else
                {
                    ReadDataDZU(reportItems, obj, reportParameters.BeginDate, reportParameters.EndDate);
                    m_writeToLog(string.Format("Загрузка объекта: {0} {1}", obj[ContractParameters.DZUGuids.RegNumber].Value.ToString(), obj[ContractParameters.DZUGuids.Contractor].Value.ToString()));
                }
            }

            reportItems = reportItems.OrderByDescending(t => t.Date).ToList();
            return reportItems;
        }

        public  void WriteReport(Xls xls, List<ReportItem> reportItems, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = reportItems.Count / 50;
            int step = 0;

            var col = 1;
            var row = 2;

            if (reportParameters.DOD)
            {
                HeadTableDOD(xls, row, col);
            }
            else
            {
                HeadTableDZU(xls, row, col);
            }

            string fileName = string.Empty;

            row = 4;

            // Запись в excel файл
            foreach (var item in reportItems)
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }

                col = 1;
                xls[col, row].SetValue(item.NameContarctor);
                col++;
                if (reportParameters.DOD)
                {
                    xls[col, row].SetValue(item.NumberOrder);
                    col++;
                }
                xls[col, row].SetValue(item.RealNumber + " (" + item.Number + " " + item.Date.ToShortDateString() + ")");
                col++;
                xls[col, row].SetValue(item.Cost);
                col++;
                xls[col, row].SetValue(item.CostNDS);
                col++;
                xls[col, row].SetValue(item.Currency);
                col++;
                if (reportParameters.DOD)
                {
                    xls[col, row].SetValue(item.AdvancePayment);
                    col++;
                }
                xls[col, row].SetValue(item.Content);

                if (reportParameters.DOD)
                {
                    col++;
                    var linesNumberOrder = string.Empty;
                    int i = 0;
                    foreach (var term in item.ListNumberOrder)
                    {
                        i++;
                        if (i != item.ListNumberOrder.Count)
                            linesNumberOrder += term + "\n";
                        else
                            linesNumberOrder += term;
                    }

                    xls[col, row].SetValue(linesNumberOrder);
                }


                if (reportParameters.DOD)
                {
                    col++;
                    var linesTerm = string.Empty;
                    int i = 0;
                    foreach (var term in item.ListTerm)
                    {
                        i++;
                        if (i != item.ListTerm.Count)
                            linesTerm += term.ToShortDateString() + "\n";
                        else
                            linesTerm += term.ToShortDateString();
                    }
                    xls[col, row].SetValue(linesTerm);
                }
                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);

                logDelegate(string.Format("Запись объекта: {0} {1}", item.Number.ToString(), item.NameContarctor));

                row++;
            }

            if (reportParameters.DOD)
            {
                xls.SetColumnWidth(1,30);
                xls.SetColumnWidth(2, 15);
                xls.SetColumnWidth(3, 30);
                xls.SetColumnWidth(8, 50);

                xls[1, 1, 1, row].WrapText = true;
                xls[2, 1, 1, row].WrapText = true;
                xls[3, 1, 1, row].WrapText = true;
                xls[8, 1, 1, row].WrapText = true;
            }
            else
            {
                xls.SetColumnWidth(1, 30);
                xls.SetColumnWidth(2, 30);
                xls.SetColumnWidth(6, 50);
               
                xls[1, 1, 1, row].WrapText = true;
                xls[2, 1, 1, row].WrapText = true;
                xls[6, 1, 1, row].WrapText = true;
            }
            xls[1, 4, col, row].VerticalAlignment = XlVAlign.xlVAlignTop;
        }

        internal static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            RunningContractsReport report = new RunningContractsReport();
            report.Make(context, reportParameters);
        }
    }
}
