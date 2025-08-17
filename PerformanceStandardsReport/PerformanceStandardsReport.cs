using System;
using System.Collections.Generic;
using System.Linq;
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

namespace PerformanceStandardsReport
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
                    !ClientView.Current.GetUser().ToString().Contains("Гудков "))
                {
                    MessageBox.Show("Формировать Ведомость начисления премии рабочим могут только сотрудники Бюро Нормирования", "Недоступная операция", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    PerformanceStandardsReport.MakeReport(context, selectionForm.reportParameters);
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

    public class PerformanceStandardsReport
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            PerformanceStandardsReport report = new PerformanceStandardsReport();
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

            PerformanceStandardsReport report = new PerformanceStandardsReport();

            m_form.setStage(IndicationForm.Stages.DataAcquisition);

            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

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
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                //if (isCatch)
                //{
                //    xls.Application.DisplayAlerts = true;
                //    xls.Quit(false);
                //}
                //else
                //{
                //    xls.Application.DisplayAlerts = true;
                //    xls.Quit(true);
                //    // Открытие файла
                //    Process.Start(context.ReportFilePath);
                //}
                xls.Application.DisplayAlerts = true;
                xls.Visible = true;
            }

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№ п/п", col, row, xls);
            col = InsertHeader("Фамилия И.О.", col, row, xls);
            col = InsertHeader("таб.№", col, row, xls);
            col = InsertHeader("Профессия", col, row, xls);
            col = InsertHeader("Разряд", col, row, xls);
            col = InsertHeader("Отработано по табелю", col, row, xls);
            col = InsertHeader("Сверхурочные", col, row, xls);
            col = InsertHeader("Выходные", col, row, xls);
            col = InsertHeader("Фактическое отр. вр.", col, row, xls);
            col = InsertHeader("Выработ. н. часов", col, row, xls);
            col = InsertHeader("% выполнения", col, row, xls);

            xls.SetRowHeight(row, 45);

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

            xls[2, row].SetValue("СВЕДЕНИЯ");
            xls[2, row].Font.Bold = true;
            xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[2, row, 7, 1].MergeCells = true;

            row++;

            xls[2, row].SetValue("о выполнении норм выработки рабочими " + reportParameters.Department + " за " + reportParameters.Month.ToLower() + " " + reportParameters.Year + " года");
            xls[2, row].Font.Bold = true;
            xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[2, row, 7, 1].MergeCells = true;

            row = row + 2;

            return row;
        }

        public void PasteWorkHours(Xls xls, ReportParameters reportParameters)
        {
            xls[10, 1].Font.Name = "Calibri";
            xls[10, 1].Font.Size = 11;
            xls[10, 1].SetValue(reportParameters.CountWorkHours);
            xls[10, 1].Font.Bold = true;
            xls[10, 1].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[10, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[10, 1].WrapText = true;

            xls[10, 1, 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                            XlBordersIndex.xlEdgeRight,
                                                                                            XlBordersIndex.xlInsideVertical);
        }


        private void MakeSelectionByTypeReport(Xls xls, Dictionary<string, List<TableData>> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            var priceWorkData = new Dictionary<string, List<TableData>>();
            var timeWorkData = new Dictionary<string, List<TableData>>();
            var pupilData = new Dictionary<string, List<TableData>>();
            var partTimeWorkerData = new Dictionary<string, List<TableData>>();
            var totalRows = new List<int>();
            foreach (var data in reportData)
            {
                var priceWorkCharges = data.Value.Where(t => t.Payment.Trim().ToLower() == "сдельная" && !t.Pupil && !t.PartTimeWorker).ToList();
                var timeWorkCharges = data.Value.Where(t => t.Payment.Trim().ToLower() == "повременная" && !t.Pupil && !t.PartTimeWorker).ToList();
                var pupilCharges = data.Value.Where(t =>  t.Pupil).ToList();
                var partTimeWorkerCharges = data.Value.Where(t => t.PartTimeWorker).ToList();

                if (priceWorkCharges.Count > 0)
                {
                    priceWorkData.Add(data.Key, priceWorkCharges);
                }

                if (timeWorkCharges.Count > 0)
                {
                    timeWorkData.Add(data.Key, timeWorkCharges);
                }

                if (pupilCharges.Count > 0)
                {
                    pupilData.Add(data.Key, pupilCharges);
                }

                if (partTimeWorkerCharges.Count > 0)
                {
                    partTimeWorkerData.Add(data.Key, partTimeWorkerCharges);
                }
            }
            //Задание начальных условий для progressBar
            var progressStep = reportData.Sum(t => t.Value.Count);
            int step = 0;

            int row = 1;
            int col = 1;
            string department = string.Empty;
            row = PasteHeadName(xls, reportParameters, col, row);

            HeaderTable(xls, reportParameters, col, row);

            xls[1, row+1, 1, 11].Select();
            xls.Application.ActiveWindow.FreezePanes = true;
            row++;

            PasteWorkHours(xls, reportParameters);

            var priceWorkRows = WriteTypePayment(xls, priceWorkData, logDelegate, progressBar, progressStep, ref step, ref row);
            var timeWorkRows = WriteTypePayment(xls, timeWorkData, logDelegate, progressBar, progressStep, ref step, ref row);
            var pupilRows = WriteTypePayment(xls, pupilData, logDelegate, progressBar, progressStep, ref step, ref row);
            var partTimeWorkerRows = WriteTypePayment(xls, partTimeWorkerData, logDelegate, progressBar, progressStep, ref step, ref row);

            if (priceWorkRows.Count > 0)
            {
                row = TotalResult(xls, row, priceWorkRows, "сдельщики");
                totalRows.Add(row);
                row = row + 2;
            }
            if (timeWorkRows.Count > 0)
            {
                row = TotalResult(xls, row, timeWorkRows, "повременщики");
                totalRows.Add(row);
                row = row + 2;
            }

            if (pupilRows.Count > 0)
            {
                row = TotalResult(xls, row, pupilRows, "уч.");
                totalRows.Add(row);
                row = row + 2;
            }

            if (partTimeWorkerRows.Count > 0)
            {
                row = TotalResult(xls, row, partTimeWorkerRows, "совместители");
                totalRows.Add(row);
                row = row + 2;
            }

            if (totalRows.Count > 1)
            {
                row = TotalResult(xls, row, totalRows, "по цеху");
                row = row + 2;
            }

            xls[2, row].SetValue("Нормировщик");
            xls[4, row].SetValue(ClientView.Current.GetUser()[ChargesParameters.ShortNameUser].Value.ToString());


            xls.SetColumnWidth(1, 4);
            xls.SetColumnWidth(2, 20);
            xls.SetColumnWidth(3, 6.71);
            xls.SetColumnWidth(4, 10.57);
            xls.SetColumnWidth(5, 6.71);
            xls.SetColumnWidth(6, 14);
            xls.SetColumnWidth(7, 14);
            xls.SetColumnWidth(8, 14);
            xls.SetColumnWidth(9, 14);
            xls.SetColumnWidth(10, 12);
            xls.SetColumnWidth(11, 12);

        }

        private List<int> WriteTypePayment(Xls xls, Dictionary<string, List<TableData>> reportData, LogDelegate logDelegate, ProgressBar progressBar, int progressStep, ref int step, ref int row)
        {
            var rowProfession = row;
            var rowsTotalResult = new List<int>();

            foreach (var profession in reportData)
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }
                int i = 1;
                rowProfession = row;

                foreach (var data in profession.Value.OrderBy(t => t.FIO))
                {
                    WriteRow(xls, logDelegate, row, i, data);
                    row++;
                    i++;
                }
                if (profession.Value.Where(t => t.PartTimeWorker).ToList().Count > 0)
                {
                    row = TotalResult(xls, row, rowProfession, "совместители");
                }
                else
                {
                    if (profession.Value.Where(t => t.Pupil).ToList().Count > 0)
                    {
                        row = TotalResult(xls, row, rowProfession, "ученики " + profession.Key);
                    }
                    else
                    {
                        row = TotalResult(xls, row, rowProfession, profession.Key);
                    }
                }
                rowsTotalResult.Add(row);
                row = row + 2;
            }
            return rowsTotalResult;
        }

        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, TableData data)
        {
            logDelegate(String.Format("Запись объекта: {0} {1} {2}", data.FIO, data.Profession, data.Category));

            xls[1, row].SetValue(i.ToString());
            xls[2, row].SetValue(data.FIO);
            xls[3, row].NumberFormat = "@";
            xls[3, row].SetValue(data.NumberTab);
            xls[4, row].SetValue(data.Profession);
            xls[5, row].SetValue(data.Category);
            xls[6, row].SetValue(data.FulfilledHours);
            xls[7, row].SetValue(data.OverTimeFulfilledHours);
            xls[8, row].SetValue(data.DaysOffFulfilledHours);
            xls[9, row].SetValue(data.ActualFulFilledHours);
            xls[10, row].SetValue(data.DevelopedStandartHours);
            xls[11, row].SetFormula(@"= ОКРУГЛ(" + xls[10, row].Address + " * 100 / (" + xls[6, row].Address + "+" + xls[7, row].Address + "+" + xls[8, row].Address + "); 1)");
            xls[1, row, 11, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static int TotalResult(Xls xls, int row, int rowProfession, string profession)
        {
            xls[1, row].SetValue("Итого " + profession + ":");
            xls[1, row, 2, 1].MergeCells = true;

            for (int j = 6; j <= 10; j++)
            {
                xls[j, row].SetFormula("=СУММ(" + xls[j, rowProfession].Address + ":" + xls[j, row - 1].Address + ")");
            }
            xls[11, row].SetFormula(@"= ОКРУГЛ(" + xls[10, row].Address + " * 100 / (" + xls[6, row].Address + "+" + xls[7, row].Address + "+" + xls[8, row].Address + "); 1)");

            xls[1, row, 11, 1].Interior.Color = Color.LightGray;
            xls[1, row, 11, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);         
            return row;
        }

        private static int TotalResult(Xls xls, int row, List<int> rows, string profession)
        {
            xls[1, row].SetValue("Итого " + profession + ":");
            xls[1, row, 2, 1].MergeCells = true;

            for (int j = 6; j <= 10; j++)
            {
                List<string> adresses = new List<string>();
                foreach (var rowRes in rows)
                {
                    adresses.Add(xls[j, rowRes].Address);
                }

                xls[j, row].SetFormula("=СУММ(" + string.Join(" + ", adresses) + ")");
            }
            xls[11, row].SetFormula(@"= ОКРУГЛ(" + xls[10, row].Address + " * 100 / (" + xls[6, row].Address + "+" + xls[7, row].Address + "+" + xls[8, row].Address + "); 1)");
            xls[1, row, 11, 1].Interior.Color = Color.LightGray;
            xls[1, row, 11, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            return row;
        }



        private Dictionary<string, List<TableData>> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");
            Dictionary<string, List<TableData>> reportData = new Dictionary<string, List<TableData>>();

            string year = reportParameters.Year.ToString();
            string month = reportParameters.MonthInt + " - " + reportParameters.Month;
            string department = reportParameters.Department;
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
                    List<ReferenceObject> foldersDepartment = folders.Where(t => t.ToString() == department && foldersMonth.Contains(t.Parent)).ToList();

                    if (foldersDepartment.Count > 0)
                    {
                        List<ReferenceObject> chargesByProfession = foldersDepartment[0].Children.OrderBy(t => GetProfession(t.Links.ToOne[ChargesParameters.ChargesGuid.LinkToWorker].
                        LinkedObject[ChargesParameters.WorkerGuid.Profession].Value.ToString())).ToList();

                        if (chargesByProfession.Count() > 0)
                        {
                            var progressStep = chargesByProfession.Count / 50;
                            int step = 0;
                            string profession = string.Empty;
                            var charges = new List<TableData>();
                            int i = 0;
                            string fio = string.Empty;

                            foreach (ReferenceObject obj in chargesByProfession.OrderBy(t=> GetProfession(t.Links.ToOne[ChargesParameters.ChargesGuid.LinkToWorker].
                        LinkedObject[ChargesParameters.WorkerGuid.Profession].Value.ToString())))
                            {
                                step++;
                                if (step >= progressStep)
                                {
                                    progressBar.PerformStep();
                                    step = 0;
                                }
                                var worker = obj.Links.ToOne[ChargesParameters.ChargesGuid.LinkToWorker].LinkedObject;

                                if (i != 0 && profession != GetProfession(worker[ChargesParameters.WorkerGuid.Profession].Value.ToString()))
                                {
                                    try
                                    {
                                        if (profession != "распр." && profession != "налад." && profession != "клад." && profession != "аппарат." && profession != "опер. о.о.")
                                            reportData.Add(profession, charges.OrderBy(t=>t.FIO).ToList());
                                        charges = new List<TableData>();
                                    }
                                    catch
                                    {
                                        MessageBox.Show("Ошибка в записи " + fio + " " + profession, "Исправьте запись в справочнике Рабочие!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }

                                var charge = new TableData();
                               
                                charge.FIO = worker[ChargesParameters.WorkerGuid.Name].Value.ToString();
                                charge.Profession = GetProfession(worker[ChargesParameters.WorkerGuid.Profession].Value.ToString());
                                charge.Category = worker[ChargesParameters.WorkerGuid.Category].Value.ToString();
                                charge.FulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.FullFilledHours].Value);
                                charge.ActualFulFilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.ActualFullFilledHours].Value);
                                charge.OverTimeFulfilledHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.OverTimeFullFilledHours].Value);
                                charge.DaysOffFulfilledHours =  Convert.ToDouble(obj[ChargesParameters.ChargesGuid.DaysOffFulFilledHours].Value) + Convert.ToDouble(obj[ChargesParameters.ChargesGuid.DaysOffFulFilledHours2].Value);
                                charge.Payment = worker[ChargesParameters.WorkerGuid.Payment].Value.ToString();
                                var number = Convert.ToDouble(worker[ChargesParameters.WorkerGuid.TimeBoardNumber].Value).ToString().Trim();
                                if (number.Length >= 4)
                                    charge.NumberTab = number;
                                else
                                {
                                    if (number.Length == 3)
                                        charge.NumberTab = "0"+number;
                                    else
                                    {
                                        if (number.Length == 2)
                                            charge.NumberTab = "00" + number;
                                        else
                                        {
                                            if (number.Length == 1)
                                                charge.NumberTab = "000" + number;
                                            else
                                                charge.NumberTab = "0000";
                                        }
                                    }
                                }
                                charge.DevelopedStandartHours = Convert.ToDouble(obj[ChargesParameters.ChargesGuid.DevelopedStandardHours].Value);
                                charge.Pupil = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.Pupil].Value);
                                charge.PartTimeWorker = Convert.ToBoolean(obj[ChargesParameters.ChargesGuid.PartTimeWorker].Value);

                                charge.Worker = worker;
                                charges.Add(charge);
                                logDelegate(String.Format("Загрузка объекта: {0} {1} {2}", charge.FIO, charge.Profession, charge.Category));
                                profession = GetProfession(worker[ChargesParameters.WorkerGuid.Profession].Value.ToString());
                                fio = charge.FIO;
                                i++;
                            }
                            if (profession != "распр." && profession != "налад." && profession != "клад.")
                                reportData.Add(profession, charges);
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
