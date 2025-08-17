using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using ReportHelpers;
using TFlex.DOCs.Model.References;
using System.Text.RegularExpressions;

namespace CooperationReport
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
                    CooperationReport.MakeReport(context, selectionForm.reportParameters);
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

    public class CooperationReport
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            CooperationReport report = new CooperationReport();
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

            List<TableData> reportData = new List<TableData>();


            Xls xls = new Xls();
            bool isCatch = false;

            CooperationReport report = new CooperationReport();

            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            m_writeToLog("=== Получение данных ===");

            reportData.AddRange(GetData(reportParameters, m_writeToLog, m_form.progressBar));



            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);


                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportData, reportParameters, m_writeToLog, m_form.progressBar);

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");

                xls.AutoWidth();
                xls.Save();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    Process.Start(context.ReportFilePath);
                }
            }

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("Заказ", col, row, xls);
            col = InsertHeader("Устройство", col, row, xls);
            col = InsertHeader("Входящее", col, row, xls);
            col = InsertHeader("Децимальный №", col, row, xls);
            col = InsertHeader("Тип входящего", col, row, xls);
            col = InsertHeader("Цена за единицу\nбез НДС, руб.", col, row, xls);
            col = InsertHeader("Контрагент", col, row, xls);
            col = InsertHeader("Договор", col, row, xls);
            col = InsertHeader("№ Документа", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("800.", col, row, xls);
            col = InsertHeader("850", col, row, xls);
            col = InsertHeader("750.", col, row, xls);
            col = InsertHeader("Итого", col, row, xls);
            if (reportParameters.ExpenseData)
            {
                col = InsertHeader("Сумма без НДС, руб.", col, row, xls);
            }
            else
            {
                col = InsertHeader("Сумма с НДС, руб.", col, row, xls);
            }
            col = InsertHeader("Дата", col, row, xls);

            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);
            return col;
        }

        public static int InsertHeader(string header, int col, int row, Xls xls)
        {

            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            col++;

            return col;
        }


        private void MakeSelectionByTypeReport(Xls xls, List<TableData> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            if (reportParameters.PeriodReport)
            {
                xls[1, 1].SetValue("ОТЧЕТ БК 936 за период " + reportParameters.BeginDate.ToShortDateString() + " - " + reportParameters.EndDate.ToShortDateString());
            }

            if (reportParameters.MonthReport)
            {
                xls[1, 1].SetValue("ОТЧЕТ БК 936 за " + reportParameters.MonthString + " " + reportParameters.Year);
            }

            if (reportParameters.DateReport)
            {
                xls[1, 1].SetValue("ОТЧЕТ БК 936 за " + reportParameters.MonthString + " " + reportParameters.Year + " до " + reportParameters.Date);
            }

            xls[1, 1].Font.Bold = true;
            xls[1, 1].Font.Size = 16;
            xls[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1, 16, 1].MergeCells = true;
            int row = 2;
            var colCount = HeaderTable(xls, reportParameters, 1, row);
            int headerRow = 0;
            bool insertSection = false;

            xls[1, row + 1, 1, 16].Select();
            xls.Application.ActiveWindow.FreezePanes = true;
            int rowMainTable = row;

            var reportText = new StringBuilder();
            var progressStep = reportData.Count / 50;
            int step = 0;
            string nameOrder = reportData[0].Order;
            string sumLabor750 = "=СУММ(";
            string sumLabor850 = "=СУММ(";
            string sumLabor800 = "=СУММ(";
            string sumSumLabor = "=СУММ(";
            string sumSum = "=СУММ(";
            int countSum = 0;

            string deviceSumLabor750 = "=СУММ(";
            string deviceSumLabor850 = "=СУММ(";
            string deviceSumLabor800 = "=СУММ(";
            string deviceSumSumLabor = "=СУММ(";
            int countDevice = 0;
            int countContracts = 0;

            string ordersSum750 = "=СУММ(";
            string ordersSum800 = "=СУММ(";
            string ordersSum850 = "=СУММ(";
            string ordersSumLabor = "=СУММ(";
            string ordersSum = "=СУММ(";
            string sumCostContract = "=СУММ(";
            string nameContract = string.Empty;
            string lastDevice = reportData[0].Device.Trim().ToLower();

            for (int i = 0; i < reportData.Count; i++)
            {
                row++;
                step++;
                if (step > progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }
                if (i != 0)
                {

                    if (reportParameters.SumOnDevices && (lastDevice.Trim().ToLower() != reportData[i].Device.Trim().ToLower() || (nameOrder != reportData[i].Order && i != 0)
                        || (reportData[i - 1].ArrivalExpenses.Count > 1 && i > 0 && reportData[i].Device.Trim().ToLower() != reportData[i - 1].Device.Trim().ToLower())))
                    {
                        //Сумма для устройства 
                        if (countDevice > 1 || reportData[i - 1].ArrivalExpenses.Count > 1)
                        {
                            xls[1, row].SetValue(nameOrder);
                            xls[2, row].SetValue(lastDevice + " Итого:");
                            xls[11, row].SetFormula(ModifyFormula(deviceSumLabor800.Remove(deviceSumLabor800.Length - 1, 1) + ")"));
                            xls[12, row].SetFormula(ModifyFormula(deviceSumLabor850.Remove(deviceSumLabor850.Length - 1, 1) + ")"));
                            xls[13, row].SetFormula(ModifyFormula(deviceSumLabor750.Remove(deviceSumLabor750.Length - 1, 1) + ")"));
                            xls[14, row].SetFormula(ModifyFormula(deviceSumSumLabor.Remove(deviceSumSumLabor.Length - 1, 1) + ")"));
                            xls[2, row, 8, 1].MergeCells = true;
                            xls[1, row, 16, 1].Interior.Color = Color.LightYellow;
                            row++;


                            xls[1, row, 16, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                            countSum = countSum + 1;
                        }

                        deviceSumLabor800 = "=СУММ(";
                        deviceSumLabor850 = "=СУММ(";
                        deviceSumLabor750 = "=СУММ(";
                        deviceSumSumLabor = "=СУММ(";


                        countDevice = 1;
                    }
                    else
                    {
                        countDevice++;
                    }
                }
                else
                {

                    if (reportData[i].Device.Trim().ToLower() == reportData[i+1].Device.Trim().ToLower())
                    {
                        countDevice++;
                    }
                }
                lastDevice = reportData[i].Device;

                if (i != 0)
                    nameContract = reportData[i - 1].Contract;

                // Суммирование по договорам
                if (!reportParameters.SumOnDevices && (nameContract != reportData[i].Contract || nameOrder != reportData[i].Order) && i != 0 && countContracts > 0)
                {
                    countSum++;

                    xls[1, row].SetValue("По договору № " + nameContract + " Итого:\t");
                    xls[15, row].SetFormula(ModifyFormula(sumCostContract.Remove(sumCostContract.Length - 1, 1) + ")"));
                    xls[2, row, 13, 1].MergeCells = true;
                    xls[1, row, 16, 1].Interior.Color = Color.LightYellow;
                    row++;

                    sumCostContract = "=СУММ(";

                    xls[1, row, 16, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

                }

                if (i != 0 && reportData[i - 1].Contract != reportData[i].Contract)
                {
                    sumCostContract = "=СУММ(";
                }

                if (nameOrder != reportData[i].Order && i != 0)
                {
                    countSum++;
                    xls[1, row].SetValue("Итого:");

                    xls[11, row].SetFormula(ModifyFormula(sumLabor800.Remove(sumLabor800.Length - 1, 1) + ")"));
                    xls[12, row].SetFormula(ModifyFormula(sumLabor850.Remove(sumLabor850.Length - 1, 1) + ")"));
                    xls[13, row].SetFormula(ModifyFormula(sumLabor750.Remove(sumLabor750.Length - 1, 1) + ")"));
                    xls[14, row].SetFormula(ModifyFormula(sumSumLabor.Remove(sumSumLabor.Length - 1, 1) + ")"));


                    xls[1, row, 16, 1].Font.Bold = true;
                    xls[1, row, 16, 1].Interior.Color = Color.LightGray;
                    xls[1, row, 9, 1].MergeCells = true;
                    xls[15, row].SetFormula(ModifyFormula(sumSum.Remove(sumSum.Length - 1, 1) + ")"));

                    ordersSum750 += xls[13, row].Address + ";";
                    ordersSum850 += xls[12, row].Address + ";";
                    ordersSum800 += xls[11, row].Address + ";";
                    ordersSumLabor += xls[14, row].Address + ";";
                    ordersSum += xls[15, row].Address + ";";
                    row++;

                    nameOrder = reportData[i].Order;

                    sumCostContract = "=СУММ(";

                    sumLabor800 = "=СУММ(";
                    sumLabor850 = "=СУММ(";
                    sumLabor750 = "=СУММ(";
                    sumSumLabor = "=СУММ(";
                    sumSum = "=СУММ(";

                }

                xls[1, row].SetValue(reportData[i].Order);
                xls[2, row].SetValue(reportData[i].Device);
                xls[3, row].SetValue(reportData[i].IncomeDevice);
                xls[4, row].SetValue(reportData[i].Denotation);
                xls[5, row].SetValue(reportData[i].IncomeType);
                xls[6, row].SetValue(reportData[i].Cost);
                xls[7, row].SetValue(reportData[i].Contractor);
                xls[8, row].SetValue(reportData[i].Contract);
                xls[9, row].SetValue(reportData[i].ArrivalExpenses[0].DocumentNumber);
                xls[10, row].SetValue(reportData[i].ArrivalExpenses[0].Count);
                xls[11, row].SetValue(reportData[i].ArrivalExpenses[0].Labor800);
                deviceSumLabor800 += xls[11, row].Address + ";";
                sumLabor800 += xls[11, row].Address + ";";


                xls[12, row].SetValue(reportData[i].ArrivalExpenses[0].Labor850);
                deviceSumLabor850 += xls[12, row].Address + ";";
                sumLabor850 += xls[12, row].Address + ";";

                xls[13, row].SetValue(reportData[i].ArrivalExpenses[0].Labor750);
                deviceSumLabor750 += xls[13, row].Address + ";";
                sumLabor750 += xls[13, row].Address + ";";

                xls[14, row].SetValue(reportData[i].ArrivalExpenses[0].SumLabor);
                sumSumLabor += xls[14, row].Address + ";";
                deviceSumSumLabor += xls[14, row].Address + ";";

                xls[15, row].SetValue(reportData[i].ArrivalExpenses[0].Sum);

                sumCostContract += xls[15, row].Address + ";";

                sumSum += xls[15, row].Address + ";";

                xls[16, row].SetValue(reportData[i].ArrivalExpenses[0].Date.ToShortDateString());

                if (reportData[i].ArrivalExpenses.Count > 1)
                {

                    for (int j = 1; j < reportData[i].ArrivalExpenses.Count; j++)
                    {
                        row++;
                        xls[9, row].SetValue(reportData[i].ArrivalExpenses[j].DocumentNumber);
                        xls[10, row].SetValue(reportData[i].ArrivalExpenses[j].Count);
                        xls[11, row].SetValue(reportData[i].ArrivalExpenses[j].Labor800);
                        deviceSumLabor800 += xls[11, row].Address + ";";
                        sumLabor800 += xls[11, row].Address + ";";
                        xls[12, row].SetValue(reportData[i].ArrivalExpenses[j].Labor850);
                        deviceSumLabor850 += xls[12, row].Address + ";";
                        sumLabor850 += xls[11, row].Address + ";";
                        xls[13, row].SetValue(reportData[i].ArrivalExpenses[j].Labor750);
                        deviceSumLabor750 += xls[13, row].Address + ";";
                        sumLabor750 += xls[13, row].Address + ";";
                        xls[14, row].SetValue(reportData[i].ArrivalExpenses[j].SumLabor);
                        deviceSumSumLabor += xls[14, row].Address + ";";
                        sumSumLabor += xls[14, row].Address + ";";
                        xls[15, row].SetValue(reportData[i].ArrivalExpenses[j].Sum);
                        sumCostContract += xls[15, row].Address + ";";
                        sumSum += xls[15, row].Address + ";";
                        xls[16, row].SetValue(reportData[i].ArrivalExpenses[j].Date.ToShortDateString());

                        xls[1, row, 16, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

                    }
                }

                logDelegate(String.Format("Запись объекта: {0} {1} {2}", i.ToString(), reportData[i].Device, reportData[i].Denotation));

                if (i != 0 && reportData[i - 1].Contract == reportData[i].Contract)
                    countContracts++;
                else
                    countContracts = 0;

                xls[1, row, 16, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

            }


            if (reportParameters.SumOnDevices && (countDevice > 1 || reportData[reportData.Count - 1].ArrivalExpenses.Count > 1))
            {
                xls[1, row].SetValue(nameOrder);
                xls[2, row].SetValue(lastDevice + " Итого:");

                xls[11, row].SetFormula(ModifyFormula(deviceSumLabor800.Remove(deviceSumLabor800.Length - 1, 1) + ")"));
                xls[12, row].SetFormula(ModifyFormula(deviceSumLabor850.Remove(deviceSumLabor850.Length - 1, 1) + ")"));
                xls[13, row].SetFormula(ModifyFormula(deviceSumLabor750.Remove(deviceSumLabor750.Length - 1, 1) + ")"));
                xls[14, row].SetFormula(ModifyFormula(deviceSumSumLabor.Remove(deviceSumSumLabor.Length - 1, 1) + ")"));

            }

            row++;
            xls[1, row].SetValue("Итого:");

            xls[11, row].SetFormula(ModifyFormula(sumLabor800.Remove(sumLabor800.Length - 1, 1) + ")"));
            xls[12, row].SetFormula(ModifyFormula(sumLabor850.Remove(sumLabor850.Length - 1, 1) + ")"));
            xls[13, row].SetFormula(ModifyFormula(sumLabor750.Remove(sumLabor750.Length - 1, 1) + ")"));
            xls[14, row].SetFormula(ModifyFormula(sumSumLabor.Remove(sumSumLabor.Length - 1, 1) + ")"));
            xls[15, row].SetFormula(ModifyFormula(sumSum.Remove(sumSum.Length - 1, 1) + ")"));
            xls[1, row, 16, 1].Font.Bold = true;
            xls[1, row, 16, 1].Interior.Color = Color.LightGray;
            xls[1, row, 9, 1].MergeCells = true;

            ordersSum750 += xls[13, row].Address + ";";
            ordersSum850 += xls[12, row].Address + ";";
            ordersSum800 += xls[11, row].Address + ";";
            ordersSumLabor += xls[14, row].Address + ";";
            ordersSum += xls[15, row].Address + ";";

            xls[1, 1, 16, row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

            row++;

            xls[1, row].SetValue("Итого по всем заказам:");

            xls[11, row].SetFormula(ModifyFormula(ordersSum800.Remove(ordersSum800.Length - 1, 1) + ")"));
            xls[12, row].SetFormula(ModifyFormula(ordersSum850.Remove(ordersSum850.Length - 1, 1) + ")"));
            xls[13, row].SetFormula(ModifyFormula(ordersSum750.Remove(ordersSum750.Length - 1, 1) + ")"));
            xls[14, row].SetFormula(ModifyFormula(ordersSumLabor.Remove(ordersSumLabor.Length - 1, 1) + ")"));
            xls[15, row].SetFormula(ModifyFormula(ordersSum.Remove(ordersSum.Length - 1, 1) + ")"));

            xls[1, row, 16, 1].Font.Bold = true;
            xls[1, row, 16, 1].Interior.Color = Color.LightGray;
            xls[1, row, 9, 1].MergeCells = true;

            xls[1, 1, 16, row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }
        public string ModifyFormula(string formula)
        {
            Regex regexDigitals = new Regex(@"(?<=\$)\d{1,}");
            Regex regexLetter = new Regex(@"(?<=\$)[A-Z]");
            formula = formula.Replace("=СУММ(", "").Replace(")", "").Trim();

            var cells = formula.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<int> digitals = new List<int>();
            string letter = regexLetter.Match(cells[0]).Value.ToString();

            foreach (var cell in cells)
            {
                var matchDigit = regexDigitals.Match(cell);
                if (matchDigit.Success)
                {
                    digitals.Add(Convert.ToInt32(matchDigit.Value));
                }
            }


            string resultString = string.Empty;
            int order = 0;
            for (int i = 0; i < digitals.Count; i++)
            {

                if ((i + 1) < digitals.Count && (digitals[i] + 1) == digitals[i + 1])
                {
                    if (order == 0)
                    {
                        if (i == 0)
                        {
                            resultString = letter + digitals[i].ToString();
                        }
                        else
                        {
                            resultString += ";" + letter + digitals[i];
                        }
                    }

                    order++;
                }
                else
                {
                    if (order == 0)
                    {
                        if (i == 0)
                            resultString = letter + digitals[i].ToString();
                        else
                        {
                            resultString += ";" + letter + digitals[i];
                        }
                    }
                    else
                    {
                        if (order == 1)
                        {
                            resultString += ";" + letter + digitals[i];
                        }
                        else
                        {
                            if (order >= 2)
                            {
                                resultString += ":" + letter + digitals[i];
                            }
                        }

                    }

                    order = 0;
                }
                if (i == digitals.Count - 1 && !resultString.Contains(letter + digitals.Last().ToString()))
                {
                    resultString += ":" + letter + digitals[i];
                }
            }
            var countSeparate = 0;
            //    resultString = "(" + resultString;
            var allSeparate = 0;
        foreach(var ch in resultString)
            {
                if (ch == ';')
                {
                    allSeparate++;
                }
            }
            if (allSeparate > 29)
            {
                for (int i = 0; i < resultString.Length; i++)
                {

                    if (resultString[i] == ';')
                    {
                        if (countSeparate == 0)
                        {
                            resultString = resultString.Insert(i + 1, "(");
                        }
                        countSeparate++;

                    }

                    if (countSeparate == 29)
                    {
                        resultString = resultString.Insert(i, ")");
                        countSeparate = 0;
                    }


                }
                resultString = resultString +")";
                // MessageBox.Show(resultString);
            }


            return "=СУММ(" +resultString + ")";
        }

        private int WriteFormula(Xls xls, int col, int row, string formula)
        {
            xls[col, row].SetFormula(formula);
            col++;
            return row;
        }

        private int WriteParameter(Xls xls, int col, int row, string parameter)
        {
            xls[col, row].SetValue(parameter);
            col++;
            return col;
        }

        private int WriteParameter(Xls xls, int col,  int row, int parameter)
        {
            xls[col, row].SetValue(parameter);
            col++;
            return col;
        }

        private static readonly Guid CooperationReferenceGuid = new Guid("38cdec3c-d13e-462e-8be1-d191b83c7150"); //Справочник Кооперация
        private static readonly Guid OrderGuid = new Guid("7c5fab79-967f-45ff-8841-c50f5c44a40a");                //Заказ
        private static readonly Guid DeviceGuid = new Guid("bf4ea4f5-2686-411e-a285-9f9ed13ebbd7");                  //устройство
        private static readonly Guid IncomeDeviceGuid = new Guid("00f07ef9-fe47-4199-9fbe-2de000bd37e7");             //входящее
        private static readonly Guid DenotationGuid = new Guid("dc573b74-1da6-409d-bbe1-edbdd2b9e568");               //обозначение
        private static readonly Guid IncomeTypeGuid = new Guid("fa531f98-47c5-4d00-8ab3-7353930c6f86");              //тип входящего
        private static readonly Guid CostGuid = new Guid("263e7dea-cd0f-4da2-bd47-e42482b1907f");                      //цена
        private static readonly Guid ContractorGuid = new Guid("a61be2ee-9698-4afc-8fe5-88c96fc7edb3");               //контрагент
        private static readonly Guid ContractGuid = new Guid("cfa44127-c027-4032-aadf-cd721a567160");                 //договор
        private static readonly Guid DocumentNumberGuid = new Guid("a5253c96-7c46-4841-9acf-5a27c8e85d53");           //Номер Документа   

        private static readonly Guid LinkRKKGuid = new Guid("87a0b7da-7239-4a3c-8b04-82baaa0e5028");                 //Связь список с одним Кооперация с ДЗУ    
        private static readonly Guid ContractNumberGuid = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4");           //Номер договора
        private static readonly Guid ContractDateGuid = new Guid("89a50a7f-2db1-48d7-9b84-e24afa2a5da7");             //Дата договора  
        private static readonly Guid ContractorRKKGuid = new Guid("a2d8a1f1-0e4f-41f0-ad7d-512fe751d157");              //контрагент

        private static readonly Guid ListParamsIncomeExpendGuid = new Guid("b17b3f24-42d2-4a3f-87ad-1501a51ce187");
        private static readonly Guid CountGuid = new Guid("cf8647c6-4523-4ff4-8250-ad8e6215d5ee");                        
        private static readonly Guid Labor800Guid = new Guid("9554d60a-6851-4d1f-988e-203c0b0e69d1");                      
        private static readonly Guid Labor850Guid = new Guid("7ebcd5d9-57ee-4300-a640-2bf0e73e74e7");
        private static readonly Guid Labor750Guid = new Guid("214d1212-7679-43c1-9f16-7153a74cd2dd");
        private static readonly Guid SumLaborGuid = new Guid("451df4c3-db66-448e-b9ff-f0c0b2aaad97");
        private static readonly Guid SumGuid = new Guid("116d6117-dad1-41a0-b329-d6348d3d622c");
        private static readonly Guid DateGuid = new Guid("eb7b8338-48e2-4b30-b096-7158149ee61e");
        private static readonly Guid MotionTypeGuid = new Guid("e6d2e17b-20f9-4399-987b-4e1d633e6272");



        private List<TableData> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");
            ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(CooperationReferenceGuid);
            Reference reference = referenceInfo.CreateReference();
            reference.LoadSettings.AddAllParameters();
            var relation = reference.LoadSettings.AddRelation(ListParamsIncomeExpendGuid);
            relation.AddAllParameters();
            relation = reference.LoadSettings.AddRelation(LinkRKKGuid);
            relation.AddParameters(ContractNumberGuid, ContractDateGuid, ContractorRKKGuid);
            //reference.LoadSettings.AddParameters(OrderGuid, DeviceGuid, IncomeDeviceGuid, DenotationGuid, IncomeTypeGuid, CostGuid, ContractorGuid);
            reference.Objects.Load();

            var objects = reference.Objects.ToList();
            int offset = 0;
            while (offset < objects.Count)
            {
                var objsToLoad = objects.Skip(offset).Take(100).ToList();
                if (objsToLoad.Count > 0) reference.LoadLinks(objsToLoad, reference.LoadSettings);
                offset += 100;
            }
            var reportData = new List<TableData>();

            int i = 0;
            var progressStep = objects.Count / 50;
            int step = 0;
            foreach (var obj in objects)
            {
                var cooperData = new TableData();
                int iderror = 0;
                try
                {
                    i++;
                    step++;
                    if (step > progressStep)
                    {
                        progressBar.PerformStep();
                        step = 0;
                    }
                   
                    iderror = 1;
                    cooperData.Order = obj[OrderGuid].Value.ToString();
                    iderror = 2;
                    cooperData.Device = obj[DeviceGuid].Value.ToString();
                    iderror = 3;
                    cooperData.IncomeDevice = obj[IncomeDeviceGuid].Value.ToString();
                    iderror = 4;
                    cooperData.Denotation = obj[DenotationGuid].Value.ToString();
                    iderror = 5;
                    cooperData.IncomeType = obj[IncomeTypeGuid].Value.ToString();
                    iderror = 6;
                    cooperData.Cost = Math.Round(Convert.ToDouble(obj[CostGuid].Value), 2);
                    iderror = 7;
                    cooperData.Contractor = obj[ContractorGuid].Value.ToString();
                    iderror = 8;

                    var docName = obj[ContractGuid].GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(docName))
                        iderror = 800;

                    if (docName.ToUpper().Contains("ДЗУ"))
                    {
                        if (obj.GetObject(LinkRKKGuid) == null)
                            iderror = 80;
                        if (obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractNumberGuid].Value == null)
                            iderror = 81;
                        if (obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractDateGuid].Value == null)
                            iderror = 82;
                        if (obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractorRKKGuid].Value == null)
                            iderror = 83;

                        cooperData.Contract = obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractNumberGuid].Value.ToString() + " от "
                            + Convert.ToDateTime(obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractDateGuid].Value).ToShortDateString() + "г.";
                        cooperData.Contractor = obj.Links.ToOne[LinkRKKGuid].LinkedObject[ContractorRKKGuid].Value.ToString();
                      
                    }
                    else
                    {
                        cooperData.Contract = obj[ContractGuid].Value.ToString().Replace("№", "").Trim();
                    }
                    iderror = 9;

                    var arrivalExpenses = new List<IncomeExpend>();

                    foreach (var incomeExpend in obj.GetObjects(ListParamsIncomeExpendGuid))
                    {
                        if (incomeExpend[MotionTypeGuid].Value.ToString() == "Расход" && reportParameters.ExpenseData)
                        {
                            arrivalExpenses.Add(new IncomeExpend(incomeExpend[CountGuid].GetInt32(), incomeExpend[Labor800Guid].GetDouble(),
                                incomeExpend[Labor850Guid].GetDouble(), incomeExpend[Labor750Guid].GetDouble(), incomeExpend[SumLaborGuid].GetDouble(),
                                incomeExpend[SumGuid].GetDouble(), Convert.ToDateTime(incomeExpend[DateGuid].Value), incomeExpend[DocumentNumberGuid].Value.ToString()));
                        }
                        if (incomeExpend[MotionTypeGuid].Value.ToString() == "Приход" && !reportParameters.ExpenseData)
                        {
                            arrivalExpenses.Add(new IncomeExpend(incomeExpend[CountGuid].GetInt32(), incomeExpend[Labor800Guid].GetDouble(),
                                incomeExpend[Labor850Guid].GetDouble(), incomeExpend[Labor750Guid].GetDouble(), incomeExpend[SumLaborGuid].GetDouble(),
                                Math.Round(incomeExpend[SumGuid].GetDouble() * 1.18, 2), Convert.ToDateTime(incomeExpend[DateGuid].Value), incomeExpend[DocumentNumberGuid].Value.ToString()));
                        }
                    }
                    iderror = 10;
                    if (reportParameters.MonthReport)
                    {
                        var incomeSelectMonth = arrivalExpenses.Where(t => t.Date.Month == reportParameters.Month && t.Date.Year == reportParameters.Year).OrderBy(t => t.Date).ToList();
                        if (incomeSelectMonth.Count > 0)
                        {
                            cooperData.ArrivalExpenses.AddRange(incomeSelectMonth);
                            reportData.Add(cooperData);
                            logDelegate(String.Format("Добавлен объект: {0} {1} {2}", i.ToString(), cooperData.Device, cooperData.Denotation));
                        }
                    }
                    else
                    {
                        if (reportParameters.PeriodReport)
                        {
                            var incomeSelectDate = arrivalExpenses.Where(t => t.Date < reportParameters.EndDate.AddDays(1) && t.Date > reportParameters.BeginDate.AddDays(-1)).OrderBy(t => t.Date).ToList();
                            if (incomeSelectDate.Count > 0)
                            {
                                cooperData.ArrivalExpenses.AddRange(incomeSelectDate);
                                reportData.Add(cooperData);
                                logDelegate(String.Format("Добавлен объект: {0} {1} {2}", i.ToString(), cooperData.Device, cooperData.Denotation));
                            }
                        }
                        else
                        {
                            var incomeSelectDate = arrivalExpenses.Where(t => t.Date < reportParameters.Date && t.Date.Month == reportParameters.Date.Month &&
                            t.Date.Year == reportParameters.Date.Year).OrderBy(t => t.Date).ToList();
                            if (incomeSelectDate.Count > 0)
                            {
                                cooperData.ArrivalExpenses.AddRange(incomeSelectDate);
                                reportData.Add(cooperData);
                                logDelegate(String.Format("Добавлен объект: {0} {1} {2}", i.ToString(), cooperData.Device, cooperData.Denotation));
                            }
                        }
                    }
                    iderror = 11;
                }
                catch
                {
                    MessageBox.Show("Внимание! Ошибка обратитесь в отдел 911! Некорректный объект " +  cooperData.Device + " " + cooperData.Denotation + ". Код ошибки " + iderror, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            //reportData = reportData.OrderBy(t => t.Device).ToList();
            //return reportData.OrderBy(t => t.Order).ToList();
            if (reportParameters.SumOnDevices)
            {
                reportData = reportData.OrderBy(t => t.Order).ThenBy(t => t.Device).ToList();
            }
            else
            {
                reportData = reportData.OrderBy(t => t.Order).ThenBy(t => t.Contract).ToList();
            }
            return reportData;
        }

    }
}
