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
using TFlex.DOCs.Model.Classes;

namespace ChargesReport
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
                ReferenceInfo info = ReferenceCatalog.FindReference(ChargesParameters.ReferenceGuid.Users);
                Reference reference = info.CreateReference();

                if (!ClientView.Current.GetUser().Parents.Contains(reference.Find(ChargesParameters.WorkToCharges)) &&
                    !ClientView.Current.GetUser().ToString().Contains("Гудков ") &&
                    !ClientView.Current.GetUser().ToString().Contains("Александров "))
                {
                    MessageBox.Show("Формировать Ведомость начисления премии рабочим могут только сотрудники Бюро Нормирования", "Недоступная операция", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    ChargesReport.MakeReport(context, selectionForm.reportParameters);
                }

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e+"", "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {

            }
        }

    }

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);

    public class ChargesReport
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            ChargesReport report = new ChargesReport();
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
            bool isCatch = false;

            ChargesReport report = new ChargesReport();

            m_form.setStage(IndicationForm.Stages.DataAcquisition);

            List<TableData> reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;
              //  xls.Open(context.ReportFilePath);

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportData, reportParameters, m_writeToLog, m_form.progressBar);

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");

              //  xls.Save();
                xls.Application.DisplayAlerts = true;

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e+"", "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                xls.Visible = true;
                //if (isCatch)
                //{
                //   // xls.Quit(false);
                //}
                //else
                //{

                //    xls.Quit(true);
                //    // Открытие файла
                //    Process.Start(context.ReportFilePath);
                //}
            }

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row, bool additionalColumn)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№ п/п", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("Фамилия И.О.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("таб.№", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("проф.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("разряд", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("тариф. ставка (оклад), руб.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("отработ. часов в месяц", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("сумма зараб. платы по тарифу, руб.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            col = InsertHeader("доплаты и надбавки, руб.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            if (reportParameters.MethodPayment == "сдельщикам")
            {
                col = InsertHeader("отраб. по д/листу, час.", col, row, xls);
                xls[col - 1, row, 1, 2].MergeCells = true;

                col = InsertHeader("сумма по д/листу, руб.", col, row, xls);
                xls[col - 1, row, 1, 2].MergeCells = true;
            }

            col = InsertHeader("процент премии к тарифу", col, row, xls);
            xls[col - 1, row, 3, 1].MergeCells = true;

            col = InsertHeader("основные показат., %", col-1, row+1, xls);
            col = InsertHeader("дополнит. показат., %", col, row+1, xls);
            col = InsertHeader("итого, % ", col, row+1, xls);

            col = InsertHeader("сумма премии, руб.", col, row, xls);
            xls[col - 1, row, 1, 2].MergeCells = true;

            if (reportParameters.MethodPayment == "сдельщикам")
            {
                col = InsertHeader("сумма зар. платы за перевып. плана, руб.", col, row, xls);
                xls[col - 1, row, 1, 2].MergeCells = true;
                col = InsertHeader("сверхнормы, часы", col, row, xls);
                if (additionalColumn)
                {
                    xls[col - 1, row, 1, 2].MergeCells = true;
                    col = InsertHeader("% Доплат и надбавок", col, row, xls);
                }
            }
            else
            {
                if (reportParameters.TypeReport == "Доплаты")
                {
                    col = InsertHeader("% занятости", col, row, xls);
                    xls[col - 1, row, 1, 2].MergeCells = true;
                }
                //if (reportParameters.TypeReport == "Сверхурочные")
                //{
                //    col = InsertHeader("ИТОГО", col, row, xls);
                //}
                col = InsertHeader("примечание", col, row, xls);
            }

            xls[col - 1, row, 1, 2].MergeCells = true;

            xls.SetRowHeight(row + 1, 43.5);

            xls[1, row, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                              XlBordersIndex.xlEdgeRight,
                                                                                              XlBordersIndex.xlInsideVertical);

            xls[1, row + 1, col - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
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
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].WrapText = true; 

            col++;

            return col;
        }

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;

            xls[col, row].SetValue("К выплате утверждаю");
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[col, row, 2, 1].MergeCells = true;

            xls[5, row].SetValue("ВЕДОМОСТЬ");
            xls[5, row].Font.Bold = true;
            xls[5, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[5, row, 6, 1].MergeCells = true;

            row = row + 2;

            xls[col, row].SetValue("Генеральный директор");
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[col, row, 2, 1].MergeCells = true;

            if (reportParameters.MethodPayment == "сдельщикам")
            {
                xls[4, row].SetValue("начисления заработной платы рабочим " + reportParameters.MethodPayment);
            }
            else
            {
                switch (reportParameters.TypeReport)
                {
                    case "По табельному времени":
                        {
                            xls[4, row].SetValue("начисления заработной платы рабочим " + reportParameters.MethodPayment);
                            break;
                        }
                    case "Доплаты":
                        {
                            xls[4, row].SetValue("начисления доплат рабочим " + reportParameters.MethodPayment);
                            break;
                        }
                    case "Сверхурочные":
                        {
                            xls[4, row].SetValue("начисления заработной платы рабочим " + reportParameters.MethodPayment + " за сверхурочную работу");
                            break;
                        }
                    case "Выходные - РВ1":
                        {
                            xls[4, row].SetValue("начисления заработной платы рабочим " + reportParameters.MethodPayment + " за работу в выходные дни с оплатой в одинарном размере");
                            break;
                        }
                    case "Выходные - РВ2":
                        {
                            xls[4, row].SetValue("начисления заработной платы рабочим " + reportParameters.MethodPayment + " за работу в выходные дни с оплатой в двойном размере");
                            break;
                        }
                    default:
                        break;
                }
            }
            xls[4, row].Font.Bold = true;
            xls[4, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[4, row, 10, 1].MergeCells = true;

            row++;

            xls[5, row].SetValue(reportParameters.Department + " за " + reportParameters.Month.ToLower() + " " + reportParameters.Year + " года");
            xls[5, row].Font.Bold = true;
            xls[5, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[5, row, 6, 1].MergeCells = true;

            row = row + 2;

            return row;
        }

        public void PasteWorkHours(Xls xls, ReportParameters reportParameters)
        {
            if (reportParameters.TypeReport == "Сверхурочные")
            {
                xls[16, 5].Font.Name = "Calibri";
                xls[16, 5].Font.Size = 11;
                xls[16, 5].SetValue(reportParameters.CountWorkHours);
                xls[16, 5].Font.Bold = true;
                xls[16, 5].Interior.Color = Color.LightPink;
                xls[16, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                xls[16, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[16, 5].WrapText = true;

                xls[16, 5, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                XlBordersIndex.xlEdgeRight,
                                                                                                XlBordersIndex.xlInsideVertical);
            }
            else
            {
                xls[15, 5].Font.Name = "Calibri";
                xls[15, 5].Font.Size = 11;
                xls[15, 5].SetValue(reportParameters.CountWorkHours);
                xls[15, 5].Font.Bold = true;
                xls[15, 5].Interior.Color = Color.LightPink;
                xls[15, 5].VerticalAlignment = XlVAlign.xlVAlignCenter;
                xls[15, 5].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[15, 5].WrapText = true;

                xls[15, 5, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                XlBordersIndex.xlEdgeRight,
                                                                                                XlBordersIndex.xlInsideVertical);
            }
        }

        public void PasteTypeReport(Xls xls, ReportParameters reportParameters)
        {
            xls[13, 2].Font.Name = "Calibri";
            xls[13, 2].Font.Size = 11;
            xls[13, 2].SetValue(reportParameters.TypeReport);
            xls[13, 2].Font.Bold = true;
            xls[13, 2].Interior.Color = Color.LightYellow;
            xls[13, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[13, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[13, 2].WrapText = true;

            xls[13, 2, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                            XlBordersIndex.xlEdgeRight,
                                                                                            XlBordersIndex.xlInsideVertical);
        }


        private void MakeSelectionByTypeReport(Xls xls, List<TableData> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            var wageSheet = 0;
            if (reportParameters.MethodPayment == "сдельщикам")
            {
                wageSheet = 2;
            }
            //Задание начальных условий для progressBar
            var progressStep = reportData.Count / 50;
            int step = 0;

            int row = 1;
            int col = 1;
            string department = string.Empty;

            row = PasteHeadName(xls, reportParameters, col, row);

            bool additionalColumn = false;
            if (reportData.Where(t=>t.PersentForExtracharges > 0).ToList().Count()>0)
            {
                additionalColumn = true;
            }

            HeaderTable(xls, reportParameters, col, row, additionalColumn);
            row = row + 2;
            int i = 1;

            PasteWorkHours(xls, reportParameters);

            if (reportParameters.TypeReport != "По табельному времени")
                PasteTypeReport(xls, reportParameters);

            foreach (var data in reportData.OrderBy(t=>t.FIO))
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }
                if (data.FulfilledHours == 0)
                    continue;

                logDelegate(String.Format("Запись объекта: {0} {1} {2}", data.FIO, data.Profession, data.Category));

                if (reportParameters.TypeReport == "Сверхурочные")
                {
                    xls[15, row].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                             XlBordersIndex.xlEdgeBottom,
                                             XlBordersIndex.xlEdgeLeft,
                                             XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

                    if (data.OvertimeNextHours != 0)
                        xls[15, row + 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                    XlBordersIndex.xlEdgeBottom,
                                                    XlBordersIndex.xlEdgeLeft,
                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

                    if (data.SverkhurochnyeGOZ != 0)
                        xls[15, row + 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                 XlBordersIndex.xlEdgeBottom,
                                                 XlBordersIndex.xlEdgeLeft,
                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                }    

                xls[1, row].SetValue(i.ToString());
                xls[2, row].SetValue(data.FIO);
                xls[3, row].NumberFormat = "@";
                xls[3, row].SetValue(data.TimeBoardNumber);
                xls[4, row].SetValue(data.Profession);
                xls[5, row].SetValue(data.Category);
                xls[6, row].SetValue(data.WageRate);
                if (reportParameters.TypeReport != "Сверхурочные")
                {
                    xls[7, row].SetValue(data.FulfilledHours);
                }
                else
                {
                    xls[7, row].SetValue(data.OvertimeFirst2Hours);
                    if (data.OvertimeNextHours != 0)
                    {
                        xls[7, row + 1].SetValue(data.OvertimeNextHours);
                        xls[2, row + 1].SetValue(data.FIO);
                        xls[3, row + 1].NumberFormat = "@";
                        xls[3, row + 1].SetValue(data.TimeBoardNumber);
                        xls[4, row + 1].SetValue(data.Profession);
                        xls[5, row + 1].SetValue(data.Category);
                        xls[6, row + 1].SetValue(data.WageRate);
                    }
                    if (data.SverkhurochnyeGOZ != 0)
                    {
                        xls[7, row + 2].SetValue(data.SverkhurochnyeGOZ);
                        xls[2, row + 2].SetValue(data.FIO);
                        xls[3, row + 2].NumberFormat = "@";
                        xls[3, row + 2].SetValue(data.TimeBoardNumber);
                        xls[4, row + 2].SetValue(data.Profession);
                        xls[5, row + 2].SetValue(data.Category);
                        xls[6, row + 2].SetValue(data.WageRate);
                    }
                }

                if (reportParameters.MethodPayment == "сдельщикам")
                {
                    if (data.InterdigitDifferenceSum == 0)
                    {
                        if ((data.NormoHours - data.FulfilledHours <= 1 && data.NormoHours - data.FulfilledHours > 0) || data.PartTimeWorker)
                        {
                            xls[8, row].SetFormula((@"=ЕСЛИ(" + data.NormoHours + @"<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + ";2));)"));
                        }
                        else
                        {
                            xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[7, row].Address + @"<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[7, row].Address + ";2));)"));
                        }
                    }
                    else
                        xls[8, row].SetValue(data.InterdigitDifferenceSum);
                }
                else
                {
                    if (reportParameters.TypeReport == "Доплаты")
                        xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + xls[7, row].Address + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[7, row].Address + "*" + xls[14, row].Address + "/100" + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + "*" + xls[14, row].Address + "/100" + ";2)))"));
                    else
                    {
                        if (reportParameters.TypeReport == "Выходные - РВ2")
                            xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + xls[7, row].Address + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[7, row].Address + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + ";2)))*2"));
                        else
                        {
                            if (reportParameters.TypeReport == "Сверхурочные")
                            {
                                xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + xls[7, row].Address + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[7, row].Address + "*1,5;2)););(ОКРУГЛ(" + xls[6, row].Address + "/$P$5*" + xls[7, row].Address + "*1,5;2)))"));
                                if (data.OvertimeNextHours != 0)
                                    xls[8, row + 1].SetFormula((@"=ЕСЛИ(" + xls[6, row + 1].Address + "<4000;ЕСЛИ(" + xls[7, row + 1].Address + "<>0;(ОКРУГЛ(" + xls[6, row + 1].Address + "*" + xls[7, row + 1].Address + "*2;2)););(ОКРУГЛ(" + xls[6, row + 1].Address + "/$P$5*" + xls[7, row + 1].Address + "*2;2)))"));
                                if (data.SverkhurochnyeGOZ != 0)
                                    xls[8, row + 2].SetFormula((@"=ЕСЛИ(" + xls[6, row + 2].Address + "<4000;ЕСЛИ(" + xls[7, row + 2].Address + "<>0;(ОКРУГЛ(" + xls[6, row + 2].Address + "*" + xls[7, row + 2].Address + "*2;2)););(ОКРУГЛ(" + xls[6, row + 2].Address + "/$P$5*" + xls[7, row + 2].Address + "*2;2)))"));
                            }
                            else
                            {
                                xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + xls[7, row].Address + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[7, row].Address + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + ";2)))"));
                            }
                        }
                    }
                }

                //Если выработанные нормо-часы больше или равны отр времени по табелю, основная премия (40%) записывается, в противном случае - нет
                //И расчет идет не по фактиче отр времеени а по нормо-часам 
                if (data.NormoHours >= data.FulfilledHours || data.Pupil)
                {
                    //Если параметр "Уволен" false премия 40% не платится
                    if (data.Dismissed)
                    {
                        xls[10 + wageSheet, row].SetValue(0);
                    }
                    else
                    {
                        if (reportParameters.TypeReport == "Сверхурочные" && data.OvertimeNextHours != 0)
                        {
                            xls[10 + wageSheet, row + 1].SetValue(data.PercentBasicAwards);
                        }
                        if (reportParameters.TypeReport == "Сверхурочные" && data.SverkhurochnyeGOZ != 0)
                        {
                            xls[10 + wageSheet, row + 2].SetValue(data.PercentBasicAwards);
                        }

                        xls[10 + wageSheet, row].SetValue(data.PercentBasicAwards);
                    }
                }
                else
                {
                    if (reportParameters.MethodPayment == "сдельщикам")
                    {
                        if (data.InterdigitDifferenceSum == 0)
                            xls[8, row].SetFormula((@"=ЕСЛИ(" + data.NormoHours + @"<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + ";2));)"));

                    }
                    else
                    {
                        if (reportParameters.TypeReport == "Доплаты")
                            xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + "*" + xls[14, row].Address + "/100" + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + "*" + xls[14, row].Address + "/100" + ";2)))"));
                        else
                        {
                            if (reportParameters.TypeReport == "Выходные - РВ2")
                                xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + ";2)))*2"));
                            else
                            {
                                if (reportParameters.TypeReport == "Сверхурочные")
                                {
                                    xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + "*1,5;2)););(ОКРУГЛ(" + xls[6, row].Address + "/$P$5*" + xls[7, row].Address + "*1,5;2)))"));
                                    if (data.OvertimeNextHours != 0)
                                        xls[8, row + 1].SetFormula((@"=ЕСЛИ(" + xls[6, row + 1].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row + 1].Address + "*" + data.NormoHours + "*2;2)););(ОКРУГЛ(" + xls[6, row + 1].Address + "/$P$5*" + xls[7, row + 1].Address + "*2;2)))"));
                                    if (data.SverkhurochnyeGOZ != 0)
                                        xls[8, row + 2].SetFormula((@"=ЕСЛИ(" + xls[6, row + 2].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row + 2].Address + "*" + data.NormoHours + "*2;2)););(ОКРУГЛ(" + xls[6, row + 2].Address + "/$P$5*" + xls[7, row + 2].Address + "*2;2)))"));
                                }
                                else
                                {
                                    xls[8, row].SetFormula((@"=ЕСЛИ(" + xls[6, row].Address + "<4000;ЕСЛИ(" + data.NormoHours + "<>0;(ОКРУГЛ(" + xls[6, row].Address + "*" + data.NormoHours + ";2)););(ОКРУГЛ(" + xls[6, row].Address + "/$O$5*" + xls[7, row].Address + ";2)))"));
                                }
                            }
                        }
                    }
                }

                if (reportParameters.TypeReport == "Сверхурочные" && data.OvertimeNextHours != 0)
                {
                    xls[11 + wageSheet, row + 1].SetValue(data.PercentAdditionalAwards);
                }
                if (reportParameters.TypeReport == "Сверхурочные" && data.SverkhurochnyeGOZ != 0)
                {
                    xls[11 + wageSheet, row + 2].SetValue(data.PercentAdditionalAwards);
                }
                xls[11 + wageSheet, row].SetValue(data.PercentAdditionalAwards);

                if (data.NormoHours >= data.FulfilledHours && !data.Dismissed)
                {
                    if (reportParameters.TypeReport == "Сверхурочные" && data.OvertimeNextHours != 0)
                    {
                        xls[12 + wageSheet, row + 1].SetFormula(@"=ЕСЛИ(" + xls[10 + wageSheet, row + 1].Address + "<>0;(" + xls[10 + wageSheet, row + 1].Address + "+" + xls[11 + wageSheet, row + 1].Address + ");)");
                    }
                    if (reportParameters.TypeReport == "Сверхурочные" && data.SverkhurochnyeGOZ != 0)
                    {
                        xls[12 + wageSheet, row + 2].SetFormula(@"=ЕСЛИ(" + xls[10 + wageSheet, row + 2].Address + "<>0;(" + xls[10 + wageSheet, row + 2].Address + "+" + xls[11 + wageSheet, row + 2].Address + ");)");
                    }

                    xls[12 + wageSheet, row].SetFormula(@"=ЕСЛИ(" + xls[10 + wageSheet, row].Address + "<>0;(" + xls[10 + wageSheet, row].Address + "+" + xls[11 + wageSheet, row].Address + ");)");
                }
                else
                {
                    if (reportParameters.TypeReport == "Сверхурочные" && data.OvertimeNextHours != 0)
                    {
                        xls[12 + wageSheet, row + 1].SetFormula(@"=(" + xls[10 + wageSheet, row + 1].Address + "+" + xls[11 + wageSheet, row + 1].Address + ")");
                    }
                    if (reportParameters.TypeReport == "Сверхурочные" && data.SverkhurochnyeGOZ != 0)
                    {
                        xls[12 + wageSheet, row + 2].SetFormula(@"=(" + xls[10 + wageSheet, row + 2].Address + "+" + xls[11 + wageSheet, row + 2].Address + ")");
                    }

                    xls[12 + wageSheet, row].SetFormula(@"=(" + xls[10 + wageSheet, row].Address + "+" + xls[11 + wageSheet, row].Address + ")");
                }

                if (reportParameters.MethodPayment == "сдельщикам")
                {
                    var суммаПроизведений = string.Join(" + ", data.ВыполняемыеРаботы.Select(t => t.СуммаЗаработнойПлаты.ToString() + " * " +
                    (t.ОсновнаяПрофессия ? "(" + t.ПроцентОсновнойПремии + " + " + xls[13, row].Address + ")" :
                    t.ПрофессияПоСовместительству ? "(" + t.ПроцентОсновнойПремии + " + " + data.PercentAdditionalAwards2 + ")" : t.ПроцентОсновнойПремии.ToString())));
                    var суммаПремии = data.InterdigitDifferenceSum == 0 ?
                        @"=ЕСЛИ(" + xls[8, row].Address + "<>0;ОКРУГЛ((" + xls[8, row].Address + "*" + xls[14, row].Address + "+(" + xls[9, row].Address + "+" + xls[11, row].Address + ")*" + xls[12, row].Address + ")/100;2);)" :
                        @"=ЕСЛИ(" + xls[8, row].Address + "<>0;ОКРУГЛ((" + суммаПроизведений + "+(" + xls[9, row].Address + "+" + xls[11, row].Address + ")*" + xls[12, row].Address + ")/100;2);)";
                    xls[15, row].SetFormula(суммаПремии);
                    xls[10, row].SetValue(data.Surcharges);
                    xls[11, row].SetFormula((@"=(ОКРУГЛ(" + xls[6, row].Address + "*" + xls[10, row].Address + ";2))"));
                }
                else
                {
                    if (reportParameters.TypeReport == "Выходные - РВ2")
                        xls[13, row].SetFormula(@"=ЕСЛИ(" + xls[8, row].Address + "<>0;ОКРУГЛ(((" + xls[8, row].Address + "+" + xls[9, row].Address + ")*" + xls[10, row].Address + " + (" + xls[8, row].Address + "+" + xls[9, row].Address + ")*" + xls[11, row].Address + "/2)/100;2);)");
                    else
                    {
                        if (reportParameters.TypeReport == "Сверхурочные")
                        {
                            //    ЕСЛИ($F$8 < 4000; ЕСЛИ($G$8 <> 0; (ОКРУГЛ($F$8 *$G$8 * 1, 5; 2));); (ОКРУГЛ($F$8 /$O$5 *$G$8 * 1, 5; 2)))
                            xls[13, row].SetFormula(@"=ЕСЛИ(" + xls[8, row].Address + "<>0;ОКРУГЛ((" + xls[8, row].Address + "*" + xls[12, row].Address + "+" + xls[9, row].Address + "*" + xls[10, row].Address + ")/100;2);)");
                            if (data.OvertimeNextHours != 0)
                                xls[13, row + 1].SetFormula(@"=ЕСЛИ(" + xls[8, row + 1].Address + "<>0;ОКРУГЛ((" + xls[8, row + 1].Address + "*" + xls[12, row + 1].Address + "+" + xls[9, row + 1].Address + "*" + xls[10, row + 1].Address + ")/100;2);)");
                            if (data.SverkhurochnyeGOZ != 0)
                                xls[13, row + 2].SetFormula(@"=ЕСЛИ(" + xls[8, row + 2].Address + "<>0;ОКРУГЛ((" + xls[8, row + 2].Address + "*" + xls[12, row + 2].Address + "+" + xls[9, row + 2].Address + "*" + xls[10, row + 2].Address + ")/100;2);)");
                        }
                         else
                         {
                            xls[13, row].SetFormula(@"=ЕСЛИ(" + xls[8, row].Address + "<>0;ОКРУГЛ((" + xls[8, row].Address + "*" + xls[12, row].Address + "+" + xls[9, row].Address + "*" + xls[10, row].Address + ")/100;2);)");
                        }
                    }
                }

                if (reportParameters.MethodPayment == "сдельщикам")
                {
                    double costHours = GetCostHours(data);
                    if (costHours > 0)
                    {
                        if (!data.Pupil && !data.PartTimeWorker)
                        {
                            //Если выработанные нормо-часы больше или равны  фактически отр времени, основная премия (40%) записывается, в противном случае - нет
                            //И расчет идет не по отр времеени а по нормо-часам
                            if (data.NormoHours >= data.FulfilledHours)
                            {
                                if (data.NormoHours - data.FulfilledHours >= 1.0167)
                                {

                                    if (data.ВыполняемыеРаботы.Count > 0)
                                    {
                                        var formula = string.Empty;
                                        formula = "= (1 -" + xls[7, row].Address + "/" + data.NormoHours + ")*(";
                                        foreach (var выполняемаяРабота in data.ВыполняемыеРаботы)
                                        {
                                            formula = formula + выполняемаяРабота.ОтработанныеЧасы.ToString() + "*" + GetCostHoursByCategory(data, выполняемаяРабота.СтоимостьЧасовСверхНорм).ToString() + "+";
                                        }
                                        //Удаление из формулы последний "+" и добавление скобки
                                        formula = formula.Remove(formula.Length - 1, 1) + ")";
                                        try
                                        {
                                            xls[16, row].SetFormula(formula);
                                            xls[16, row].NumberFormat = "#,##0.00";
                                        }
                                        catch
                                        {
                                            Clipboard.SetText(formula);
                                            MessageBox.Show("2", formula);
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            xls[16, row].SetFormula("= (" + data.NormoHours + " - " + xls[7, row].Address + ")*" + costHours);
                                        }
                                        catch
                                        {
                                            Clipboard.SetText("= (" + data.NormoHours + " - " + xls[7, row].Address + ")*" + costHours);
                                            MessageBox.Show("1", "= (" + data.NormoHours + " - " + xls[7, row].Address + ")*" + costHours);
                                        }
                                    }
                                }
                                xls[17, row].SetFormula("= (" + data.NormoHours + " - " + xls[7, row].Address + ")");
                            }
                            else
                            {
                            }

                            //Если процент выработки больше 140% меняем цвет строки
                            if(data.ActualFulfilledHours > 0 && data.NormoHours/data.ActualFulfilledHours > 1.4)
                            {
                                xls[1, row, 17, 1].Interior.Color = Color.LightBlue;
                            }
                        }
                    }
                       
                }

                if (data.PercentForSurcharges != 0)
                {
                    if (reportParameters.TypeReport == "Доплаты")
                    {
                        xls[15, row].SetValue(data.Comment);
                       // xls[9, row].SetFormula("=ОКРУГЛ(" + xls[6, row].Address + "*" + xls[14, row].Address + "/100;2)");
                    }
                }

                if (data.PersentForExtracharges != 0)
                {
                    if (reportParameters.TypeReport != "Доплаты")
                    {
                        // xls[14, row].SetValue(data.PersentForExtracharges);
                        xls[17 + wageSheet - 1, row].SetValue(data.PersentForExtracharges);
                        xls[9, row].SetFormula("=ОКРУГЛ(" + xls[8, row].Address + "*" + xls[17 + wageSheet - 1, row].Address + "/100;2)");
                    }
                    //MessageBox.Show("2 - " + xls[16, row].Value);
                }

                if (reportParameters.TypeReport == "Доплаты")
                {
                    xls[14, row].SetValue(data.PercentForSurcharges);
                    xls[1, row, 15, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                              XlBordersIndex.xlEdgeBottom,
                                                                                                              XlBordersIndex.xlEdgeLeft,
                                                                                                              XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                }
                else
                {
                    if (additionalColumn)
                    {
                        if (reportParameters.TypeReport == "Сверхурочные")
                        {
                           // xls[14, row].SetFormula(@"= " + xls[8, row].Address + " + " + xls[13, row].Address);
                            xls[1, row, 15, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                            if (data.OvertimeNextHours != 0)
                            {
                                //xls[14, row + 1].SetFormula(@"= " + xls[8, row + 1].Address + " + " + xls[13, row + 1].Address);
                                xls[1, row + 1, 15, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                            }
                            if (data.SverkhurochnyeGOZ != 0)
                            {
                               // xls[14, row + 2].SetFormula(@"= " + xls[8, row + 2].Address + " + " + xls[13, row + 2].Address);
                                xls[1, row + 2, 15, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                            }
                        }
                        else
                            xls[1, row, 18, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                    }
                    else
                    {
                        if (reportParameters.MethodPayment == "сдельщикам")
                            xls[1, row, 17, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                        else
                        {
                            if (reportParameters.TypeReport == "Сверхурочные")
                            {
                               // xls[14, row].SetFormula(@"= " + xls[8, row].Address + " + " + xls[13, row].Address);
                                xls[1, row, 13, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                                XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                                if (data.OvertimeNextHours != 0)
                                {
                                  //  xls[14, row + 1].SetFormula(@"= " + xls[8, row + 1].Address + " + " + xls[13, row + 1].Address);
                                    xls[1, row + 1, 13, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                                }
                                if (data.SverkhurochnyeGOZ != 0)
                                {
                                //    xls[14, row + 2].SetFormula(@"= " + xls[8, row + 2].Address + " + " + xls[13, row + 2].Address);
                                    xls[1, row + 2, 13, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                                }
                            }
                            else
                            {
                                xls[1, row, 14, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                                XlBordersIndex.xlEdgeBottom,
                                                                                                                XlBordersIndex.xlEdgeLeft,
                                                                                                                XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                            }
                        }
                    }
                }
                if (reportParameters.TypeReport == "Сверхурочные" && data.SverkhurochnyeGOZ != 0)
                {
                    xls[15, row + 2].SetValue("ГОЗ");
                    //xls[15, row + 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
                    xls[15, row + 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    // xls[9, row].SetFormula("=ОКРУГЛ(" + xls[6, row].Address + "*" + xls[14, row].Address + "/100;2)");
                }

                if (reportParameters.TypeReport == "Сверхурочные")
                {
                    row++;
                    if (data.OvertimeNextHours != 0)
                        row++;
                    if (data.SverkhurochnyeGOZ != 0)
                        row++;
                }
                else
                {
                    row++;
                }

                i++;
            }

            row++;

            xls[12, row].SetValue("Проверено");
          //  xls[12, row].HorizontalAlignment = HorizontalAlignment.Left;
            xls[12, row, 3, 1].MergeCells = true;
            row++;

            xls[1, row].SetValue("Начальник подразделения _____________________");
           // xls[1, row].HorizontalAlignment = HorizontalAlignment.Left;
            xls[1, row, 4, 1].MergeCells = true;
            xls[12, row].SetValue("Начальник отд. 904  __________________");
           // xls[12, row].HorizontalAlignment = HorizontalAlignment.Left;
            xls[12, row, 3, 1].MergeCells = true;

            xls.SetColumnWidth(1, 7);
            xls.SetColumnWidth(2, 23.71);
            xls.SetColumnWidth(3, 8.43);
            xls.SetColumnWidth(4, 10.14);
            xls.SetColumnWidth(5, 8.43);
            xls.SetColumnWidth(6, 8.43);
            xls.SetColumnWidth(7, 8.43);
            xls.SetColumnWidth(8, 11.86);
            xls.SetColumnWidth(9, 11.86);
            xls.SetColumnWidth(10, 11.86);
            xls.SetColumnWidth(11, 11.86);
            xls.SetColumnWidth(12, 9.57);
            xls.SetColumnWidth(13, 13.57);
            xls.SetColumnWidth(14, 11.71);
            if (reportParameters.TypeReport != "Сверхурочные")
            {
                xls.SetColumnWidth(15, 11.86);
                xls.SetColumnWidth(16, 13.57);
                xls.SetColumnWidth(17, 11.86);
                xls.SetColumnWidth(18, 11.86);
            }
            else
            {
                xls.SetColumnWidth(15, 13.57);
                xls.SetColumnWidth(16, 11.86);
                xls.SetColumnWidth(17, 13.57);
                xls.SetColumnWidth(18, 11.86);
                xls.SetColumnWidth(19, 11.86);
            }
        }

        // Вычисление стоимости часа в соответствии с сеткой сверхнорм
        private static double GetCostHours(TableData data)
        {
            double countHoursOverFullfilledTime = data.NormoHours - data.FulfilledHours;
            double costHours = 0;
            if (countHoursOverFullfilledTime > 0)
            {
                ReferenceObject profession = data.Worker.GetObject(ChargesParameters.WorkerGuid.LinkToOverFullfilledTime); // Профессия (станочник или остальные проффессии в соответствии с связью)
                var tableOverFullfilledTime = profession.GetObjects(ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHoursListObjects);
                foreach (var countHours in tableOverFullfilledTime.OrderBy(t => Convert.ToInt32(t[ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHours].Value)).ToList())
                {
                    if (countHoursOverFullfilledTime < Convert.ToInt32(countHours[ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHours].Value))
                    {
                        costHours = Convert.ToDouble(countHours[ChargesParameters.CostHoursOverFullfilledTimeGuid.CostHours].Value);
                        break;
                    }
                }
            }
            return costHours;
        }

        // Вычисление стоимости часа в соответствии с сеткой сверхнорм для рабочих с несколькими профессиями по совместительству
        private static double GetCostHoursByCategory(TableData data, ReferenceObject стоимостьЧасовСверхНорм)
        {
            double countHoursOverFullfilledTime = data.NormoHours - data.FulfilledHours;
            double costHours = 0;


            if (countHoursOverFullfilledTime > 0)
            {
                var tableOverFullfilledTime = стоимостьЧасовСверхНорм.GetObjects(ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHoursListObjects);
                foreach (var countHours in tableOverFullfilledTime.OrderBy(t => Convert.ToInt32(t[ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHours].Value)).ToList())
                {
                    if (countHoursOverFullfilledTime < Convert.ToInt32(countHours[ChargesParameters.CostHoursOverFullfilledTimeGuid.CountHours].Value))
                    {
                        costHours = Convert.ToDouble(countHours[ChargesParameters.CostHoursOverFullfilledTimeGuid.CostHours].Value);
                        break;
                    }
                }
            }
            return costHours;
        }
        private List<TableData> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");
            List<TableData> reportData = new List<TableData>();

            string year = reportParameters.Year.ToString();
            string month = reportParameters.MonthInt + " - " + reportParameters.Month;
            string department = reportParameters.Department;

            //Создание папок в справочнике Начисления в соответствие с выбранным месяцем и годом
            //------------------------------------------------------------------------------------
            ReferenceInfo referenceChargesInfo = ReferenceCatalog.FindReference(ChargesParameters.ReferenceGuid.Charges);
            Reference referenceCharges = referenceChargesInfo.CreateReference();
            //Находим тип «Папка»
            ClassObject classFolder = referenceChargesInfo.Classes.Find(ChargesParameters.ReferenceTypesGuid.Folder);
            //Создаем фильтр
            TFlex.DOCs.Model.Search.Filter filterFolder = new TFlex.DOCs.Model.Search.Filter(referenceChargesInfo);
            //Добавлем условие поиска - «Тип = Папка»
            filterFolder.Terms.AddTerm(referenceCharges.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classFolder);
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр     
            List<ReferenceObject> folders = referenceCharges.Find(filterFolder);
            List<ReferenceObject> foldersYear = folders.Where(t => t.ToString() == year && t.Parent == null).ToList();


            if (foldersYear.Count > 0)
            {
                List<ReferenceObject> foldersMonth = folders.Where(t => t.ToString() == month && foldersYear.Contains(t.Parent)).ToList();

                if (foldersMonth.Count > 0)
                {
                    string methodPayment = string.Empty;

                    if (reportParameters.MethodPayment == "сдельщикам")
                    {
                        methodPayment = "сдельная";
                    }    
                    else
                    {
                        methodPayment = "повременная";
                    }
                    List<ReferenceObject> foldersDepartment = folders.Where(t => t.ToString() == department && foldersMonth.Contains(t.Parent)).ToList();
                   
                    if (foldersDepartment.Count > 0)
                    {
                        /*                        
                        На следующей строке отчет падает в следубщем случае:
                        Если создано начисление, но был удален рабочий (либо рабочий находится в другом подразделении, я так и не понял, возможно он был перемещен).
                        Так же если попытаться открыть запись с данным начислением, то макрос начала редактирования объекта падает и все заблокированные поля информации о данном рабочем остаются пустыми.
                        Пришлось удалить одну сбойную запись в Цех 75, помоему фамилия Порошкин (после сбоя открытия диалога редактирования начисления все поля стали пустыми и я точно не запомнил). 
                        После этого отчет заработал.

                        Андрей

                        */
                        List<ReferenceObject> chargesByMethodPayment = foldersDepartment[0].Children
                            .Where(t => t.GetObject(ChargesParameters.ChargesGuid.LinkToWorker)[ChargesParameters.WorkerGuid.Payment].GetString() == methodPayment).ToList();

                        if (chargesByMethodPayment.Count() > 0)
                        {
                            var progressStep = chargesByMethodPayment.Count / 50;
                            int step = 0;

                            foreach (ReferenceObject obj in chargesByMethodPayment.OrderBy(t=>t.ToString()))
                            {
                                step++;
                                if (step >= progressStep)
                                {
                                    progressBar.PerformStep();
                                    step = 0;
                                }
                                var charge = new TableData();
                                var worker = obj.Links.ToOne[ChargesParameters.ChargesGuid.LinkToWorker].LinkedObject;
                                charge.FIO =worker[ChargesParameters.WorkerGuid.Name].Value.ToString();

                                var number = Convert.ToDouble(worker[ChargesParameters.WorkerGuid.TimeBoardNumber].Value).ToString().Trim();
                                if (number.Length >= 4)
                                    charge.TimeBoardNumber = number;
                                else
                                {
                                    if (number.Length == 3)
                                        charge.TimeBoardNumber = "0" + number;
                                    else
                                    {
                                        if (number.Length == 2)
                                            charge.TimeBoardNumber = "00" + number;
                                        else
                                        {
                                            if (number.Length == 1)
                                                charge.TimeBoardNumber = "000" + number;
                                            else
                                                charge.TimeBoardNumber = "0000";
                                        }
                                    }
                                }
                               // charge.TimeBoardNumber = Convert.ToInt32(worker[ChargesParameters.WorkerGuid.TimeBoardNumber].Value);
                                charge.Profession = GetProfession(worker[ChargesParameters.WorkerGuid.Profession].Value.ToString());
                                charge.Category = worker[ChargesParameters.WorkerGuid.Category].Value.ToString();
                                charge.WageRate = Convert.ToDouble(worker[ChargesParameters.WorkerGuid.WageRate].Value.ToString());
                                charge.PersentForExtracharges = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.PercentForExtraCharges].Value);
                                charge.Comment = obj[ChargesParameters.ChargesGuid.Comment].Value.ToString();
                                charge.NormoHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.NormoHours].Value);
                                charge.ActualFulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.ActualFullFilledHours].Value);

                                if (obj.Links.ToMany[ChargesParameters.ChargesGuid.TariffCategoryLink].Objects.Count() > 0)
                                {
                                    charge.InterdigitDifferenceSum = Math.Round(obj.Links.ToMany[ChargesParameters.ChargesGuid.TariffCategoryLink].Objects.Sum(t => Convert.ToDouble(t[ChargesParameters.ChargesGuid.SumSalaryTariffCategoryLink].Value)),2);
                                    //Если у рабочего есть профессия по совместительству
                                    var межразряднаяРазница = obj.Links.ToMany[ChargesParameters.ChargesGuid.TariffCategoryLink].Objects;
                                    if (worker[ChargesParameters.WorkerGuid.ProfessionInCombination].Value.ToString().Length > 1 || межразряднаяРазница.Count() > 0)
                                    {
                                        foreach (var work in межразряднаяРазница)
                                        {
                                            // тарифная группа станочников за исключением граверов
                                            var стоимостьЧасовСверхНорм = work.Links.ToOne[ChargesParameters.ChargesGuid.LinkTariffGuid]
                                                .LinkedObject.Links.ToOne[ChargesParameters.ChargesGuid.LinkTariffGroupsGuid].LinkedObject
                                                .GetObject(ChargesParameters.ChargesGuid.СвязьТарифнойГруппыСоСтоимостьюЧасовСвышеНорм);

                                            var связьСПрфессией = work.GetObject(ChargesParameters.ChargesGuid.СвязьВыполняемойРаботыСПрофессией);
                                            var премия = связьСПрфессией != null ? Convert.ToDouble(связьСПрфессией[ChargesParameters.ChargesGuid.ПроцентПремииПрофессии].Value) : 40;

                                            var выполняемыеРаботы = new ВыполняемаяРабота(стоимостьЧасовСверхНорм,
                                                worker[ChargesParameters.WorkerGuid.Profession].Value.ToString() == work[ChargesParameters.ChargesGuid.ПрофессияВыполняемойРаботы].Value.ToString(),
                                                worker[ChargesParameters.WorkerGuid.ProfessionInCombination].Value.ToString() == work[ChargesParameters.ChargesGuid.ПрофессияВыполняемойРаботы].Value.ToString(),
                                                Convert.ToDouble(work[ChargesParameters.ChargesGuid.CountHoursWorkGuid].Value),
                                                премия,
                                                Convert.ToDouble(work[ChargesParameters.ChargesGuid.SumSalaryTariffCategoryLink].Value));

                                            charge.ВыполняемыеРаботы.Add(выполняемыеРаботы);
                                        }
                                    }
                                }
                                else
                                {
                                    charge.InterdigitDifferenceSum = 0;
                                }

                                switch (reportParameters.TypeReport)
                                {
                                    case "По табельному времени":
                                        {
                                            charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.FullFilledHours].Value);
                                            break;
                                        }
                                    case "Доплаты":
                                        {
                                            charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.SurchargeFullFilledHours].Value);
                                            break;
                                        }
                                    case "Сверхурочные":
                                        {
                                            charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.OverTimeFullFilledHours].Value);
                                            charge.OvertimeFirst2Hours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.OvertimeFirst2Hours].Value);
                                            charge.OvertimeNextHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.OvertimeNextHours].Value);
                                            charge.SverkhurochnyeGOZ = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.SverkhurochnyeGOZ].Value);
                                            break;
                                        }
                                    case "Выходные - РВ1":
                                        {
                                            charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.DaysOffFulFilledHours].Value);
                                            break;
                                        }
                                    case "Выходные - РВ2":
                                        {
                                            charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.DaysOffFulFilledHours2].Value);
                                            break;
                                        }
                                    default:
                                        break;
                                }
                                int errorId = 0;
                                try
                                {
                                    errorId = 1;
                                    if (reportParameters.TypeReport == "По табельному времени")
                                    charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.FullFilledHours].Value);
                                    errorId = 2;

                                    //   charge.ActualFulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.ActualFullFilledHours].Value);
                                    charge.Pupil = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.Pupil].Value);
                                    charge.PartTimeWorker = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.PartTimeWorker].Value);
                                    errorId = 3;
                                    charge.Dismissed = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.DismissedGuid].Value);
                                    errorId = 4;
                                    charge.CalcNormByOtherProf = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.CalcNormByOtherProf].Value);
                                    errorId = 5;
                                    charge.Surcharges = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.Surcharges].Value);
                                    errorId = 6;
                                    charge.PercentBasicAwards = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.BasicAward].Value);
                                    errorId = 7;
                                    charge.PercentAdditionalAwards = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.AdditionalAward].Value);
                                    charge.PercentAdditionalAwards2 = Convert.ToDouble(obj.GetObject(ChargesParameters.ChargesGuid.LinkToWorker)[ChargesParameters.WorkerGuid.PersentAdditionalAwards2].Value);

                                    errorId = 8;
                                    if (reportParameters.TypeReport == "Доплаты")
                                        charge.PercentForSurcharges = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.EmploymentPercent].Value);
                                    errorId = 9;
                                    charge.Worker = worker;
                                    errorId = 10;
                                    reportData.Add(charge);
                                    logDelegate(String.Format("Загрузка объекта: {0} {1} {2}", charge.FIO, charge.Profession, charge.Category));
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show("Ошибка. Обратитесь в отдел 911. Код ошибки " + errorId, "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Данных для указанного подразделения " + department + " не найдено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Данных для указанного подразделения " + department + " не найдено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }
                }
                else
                {
                    MessageBox.Show("Данных за указанный месяц " + month + " не найдено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
            }
            else
            {
                MessageBox.Show("Данных за указанный год " + year + " не найдено", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            //--------------------------------------------------------------------------------------------------------------------------------

            reportData = reportData.OrderBy(t => t.FIO).ToList();
            return reportData;
        }

        private string GetProfession(string profession)
        {
            string resultProfession = profession;

            switch (profession)
            {
                case "аппаратчик":
                    {
                        resultProfession = "аппарат.";
                        break;
                    }
                case "гальваник":
                    {
                        resultProfession = "гальв.";
                        break;
                    }
                case "гравер":
                    {
                        resultProfession = "грав.";
                        break;
                    }
                case "заточник":
                    {
                        resultProfession = "заточ.";
                        break;
                    }
                case "измеритель электрофизических параметров изделий":
                    {
                        resultProfession = "изм.э/ф п.";
                        break;
                    }
                case "каменщик":
                    {
                        resultProfession = "кам.";
                        break;
                    }
                case "кладовщик":
                    {
                        resultProfession = "клад.";
                        break;
                    }
                case "комлектовщик":
                    {
                        resultProfession = "компл.";
                        break;
                    }
                case "литейщик":
                    {
                        resultProfession = "лит.";
                        break;
                    }
                case "маркировщик":
                    {
                        resultProfession = "марк.";
                        break;
                    }
                case "монтажник":
                    {
                        resultProfession = "монт.";
                        break;
                    }
                case "монтажник (гибридные микросборки)":
                    {
                        resultProfession = "монт.";
                        break;
                    }
                case "наладчик":
                    {
                        resultProfession = "налад.";
                        break;
                    }
                case "намотчик":
                    {
                        resultProfession = "намот.";
                        break;
                    }
                case "окрасчик":
                    {
                        resultProfession = "окрас.";
                        break;
                    }
                case "оператор":
                    {
                        resultProfession = "опер.";
                        break;
                    }
                case "оператор очистного оборудования":
                    {
                        resultProfession = "опер. о.о.";
                        break;
                    }
                case "оператор лазерных установок":
                    {
                        resultProfession = "оператор ЛУ";
                        break;
                    }
                case "прессовщик":
                    {
                        resultProfession = "пресс.";
                        break;
                    }
                case "промывщик плат":
                    {
                        resultProfession = "промыв.";
                        break;
                    }
                case "пропитчик":
                    {
                        resultProfession = "пропит.";
                        break;
                    }
                case "распределитель работ":
                    {
                        resultProfession = "распр.";
                        break;
                    }
                case "регулировщик":
                    {
                        resultProfession = "регул.";
                        break;
                    }
                case "сварщик":
                    {
                        resultProfession = "сварщ.";
                        break;
                    }
                case "слесарь":
                    {
                        resultProfession = "слес.";
                        break;
                    }
                case "слесарь-инструментальщик":
                    {
                        resultProfession = "слес.-ин.";
                        break;
                    }
                case "столяр":
                    {
                        resultProfession = "стол.";
                        break;
                    }
                case "ученик":
                    {
                        resultProfession = "уч.";
                        break;
                    }
                case "оператор оборудования цифровой печати":
                    {
                        resultProfession = "опер.о.ц.п.";
                        break;
                    }
                case "фрезеровщик":
                    {
                        resultProfession = "фрез.";
                        break;
                    }
                case "шлифовщик":
                    {
                        resultProfession = "шлиф.";
                        break;
                    }
                case "штамповщик":
                    {
                        resultProfession = "штамп.";
                        break;
                    }
                case "штукатур":
                    {
                        resultProfession = "штук.";
                        break;
                    }
            }

            return resultProfession;
        }
    }
}
