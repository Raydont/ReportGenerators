using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace GroupingOIReport
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
                    selectionForm.Init(context);

                    if (selectionForm.ShowDialog() == DialogResult.OK)
                    {
                        GroupingOIReport.MakeReport(context, selectionForm.reportParameters);
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
    public class WaitCursorHelper : IDisposable
    {
        private bool _waitCursor;

        public WaitCursorHelper(bool useWaitCursor)
        {
            _waitCursor = System.Windows.Forms.Application.UseWaitCursor;
            System.Windows.Forms.Application.UseWaitCursor = useWaitCursor;
        }

        public void Dispose()
        {
            System.Windows.Forms.Application.UseWaitCursor = _waitCursor;
        }
    }
    public class GroupingOIReport
    {
        public IndicationForm m_form = new IndicationForm();
        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            GroupingOIReport report = new GroupingOIReport();
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
            m_form.setStage(IndicationForm.Stages.DataAcquisition);
            int errorId = 1;
            string nameDevice = string.Empty;
            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);//.Where(t=>t.Device != null).ToList(); 

            try
            {
                if (reportData.Count(t => t.NotFoundedObject.Count > 0) > 0)
                {
                    errorId = 2;
                    var notFoundObjectsForm = new NotFoundObjectForm();
                    notFoundObjectsForm.SortedData.AddRange(reportData);
                    m_form.Visible = false;
                    int i = 0;
                    errorId = 378;
                    // MessageBox.Show(string.Join("\r\n", reportData.Select(t => t.Device)));
                    errorId = 3;
                    foreach (var device in reportData.OrderBy(t => t.Device[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                    {
                        errorId = 4;
                        foreach (var obj in device.NotFoundedObject.OrderBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()))
                        {
                            errorId = 5;
                            i++;
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(i.ToString());
                            errorId = 6;
                            itmx[0].SubItems.Add(obj.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString());
                            errorId = 7;
                            itmx[0].SubItems.Add(obj.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString());
                            errorId = 8;
                            itmx[0].SubItems.Add(device.Device[Guids.NomenclatureParameters.NameGuid].Value + " " + device.Device[Guids.NomenclatureParameters.DenotationGuid].Value);
                            errorId = 9;
                            itmx[0].Tag = device;
                            errorId = 10;
                            notFoundObjectsForm.listViewNotFoundObjects.Items.AddRange(itmx);
                            errorId = 11;
                        }
                    }

                    // GroupsReferenceNames groups;
                    errorId = 12;
                    foreach (var firstSection in GroupsReferenceNames.supplyGroup)
                    {
                        errorId = 13;
                        notFoundObjectsForm.comboBoxSections.Items.Add(firstSection.Key.ToString());
                    }
                    errorId = 14;
                    notFoundObjectsForm.comboBoxSections.Text = notFoundObjectsForm.comboBoxSections.Items[0].ToString();
                    errorId = 15;
                    //  Clipboard.SetText(string.Join("\r\n", reportData.Where(t => t.FindedObj == false).Select(t => t.Name + " " + t.Denotation).OrderBy(t => t).ToList()));
                    if (notFoundObjectsForm.ShowDialog() == DialogResult.OK)
                    {
                        errorId = 16;
                        reportData = new List<DeviceByGroups>();
                        reportData.AddRange(notFoundObjectsForm.SortedData);
                        errorId = 17;
                        m_form.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("3.  Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки " + errorId + "\r\n" + ex.ToString(), "Внимание ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (reportData == null)
            {
                m_form.Close();
                return;
            }
            CoastData coastData = null;

            List<DeviceByGroups> sortReportData = SortObjects(m_writeToLog, reportData);
            if (reportParameters.IsCreateObjectCost)
            {
                m_writeToLog("=== Загрузка счетов фактур для получения стоимости объектов ===");
                List<NomenclatureObject> nomenclatureObjects = new List<NomenclatureObject>();
                foreach (var device in sortReportData)
                {
                    foreach (var group in device.Groups)
                    {
                        var nomObjects = group.Value.Select(t => t.NomObject).ToList().Where(t => !nomenclatureObjects.Contains(t)).ToList();
                        nomObjects.ForEach(t => nomenclatureObjects.Add(t));
                    }
                }
                coastData = new CoastData();
                coastData.SearchCostObject(nomenclatureObjects.Distinct().ToList(), m_writeToLog);
            }

            try
            {
                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                //Формирование отчета
                MakeSelectionByTypeReport(sortReportData, reportParameters, m_writeToLog, m_form.progressBar, coastData);

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            m_form.Dispose();
        }

        private static List<DeviceByGroups> SortObjects(LogDelegate m_writeToLog, List<DeviceByGroups> reportData)
        {
            m_writeToLog("=== Сортировка устройств ===");

            if (reportData.Count(t => t.Device.ToString().Contains("Воздухозаправщик 1Ю311")) > 1 &&
               reportData.Count(t => t.Device.ToString().Contains("Воздухозаправщик 1Ю311") && t.Device.ToString().Contains("ЗИП")) > 0 &&
               reportData.Count(t => t.Device.ToString().Contains("Азотозаправщик 1Ю312")) > 1 &&
               reportData.Count(t => t.Device.ToString().Contains("Азотозаправщик 1Ю312") && t.Device.ToString().Contains("ЗИП")) > 0)
            {
                var komplekt = reportData.Where(t => t.Device.ToString().Contains("Азотозаправщик 1Ю312") && t.Device.ToString().Contains("ЗИП")).FirstOrDefault();
                for (int i = 0; i < komplekt.Groups.Values.Count; i++)
                {
                    var countKomplekt = komplekt.Groups.Values.ToList()[i].Count();
                    List<NomSpecificationItems> removeObject = new List<NomSpecificationItems>();
                    for (int j = 0; j < countKomplekt; j++)
                    {
                        if (!komplekt.Groups.Values.ToList()[i].ToList()[j].NomObject.ToString().Contains("Фильтр"))
                        {
                            removeObject.Add(komplekt.Groups.Values.ToList()[i].ToList()[j]);
                            //   MessageBox.Show("Удаляю " + komplekt.Groups.Values.ToList()[i].ToList()[j].NomObject);
                        }
                    }
                    foreach (var obj in removeObject)
                    {
                        komplekt.Groups.Values.ToList()[i].Remove(obj);
                    }
                }

                reportData.Remove(reportData.Where(t => t.Device.ToString().Contains("Азотозаправщик 1Ю312") && t.Device.ToString().Contains("ЗИП")).FirstOrDefault());
                reportData.Add(komplekt);
            }
            var sortRegex = new Regex(@"(\d{1}[А-Я]{1}\d{3}(-\d{1,})?)|((кабель|Кабель|КАБЕЛЬ|Заглушка|Шланг|Шина|Корпус шкафа|Распредкоробка шкафа)\s?\w{1,})|(ЭМ\d\z)");
            var sortReportData = reportData.OrderBy(t => sortRegex.Match(t.Device[Guids.NomenclatureParameters.NameGuid].Value.ToString().Trim()).Value.ToString()).ThenBy(t => t.Device.ToString()).ToList();


            var kabelRegex = new Regex(@"(?<=(^кабель|^Кабель|^КАБЕЛЬ)\s?(№?))\d{1,}");

            foreach (var obj in sortReportData)
            {
                if (obj.NameLoadList == string.Empty)
                {
                    obj.NameLoadList = obj.Device[Guids.NomenclatureParameters.NameGuid].Value.ToString();
                }
            }

            var listKabel = sortReportData.Where(t => kabelRegex.Match(t.Device[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Success).ToList();
            var listKabelSort = listKabel.OrderBy(t => Convert.ToInt32(kabelRegex.Match(t.Device[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value)).ToList();

            if (listKabel.Count > 0)
            {
                var indexKabel = sortReportData.IndexOf(listKabel[0]);

                foreach (var kabel in listKabel)
                {
                    sortReportData.Remove(kabel);
                }

                sortReportData.InsertRange(indexKabel, listKabelSort);
            }
            return sortReportData;
        }

        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            // xls[col, row].Interior.Color = Color.LightGray;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].WrapText = true;

            return col + 1;
        }

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row, int sizeFont, string head)
        {
            xls[1, row].SetValue(head);
            xls[1, row].Font.Bold = true;
            xls[col, row].Font.Size = sizeFont;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 4, 1].MergeCells = true;

            return row + 1;
        }

        private int PasteHeader(Xls xls, List<string> headers, int col, int row)
        {
            headers.ForEach(t => col = InsertHeader(t, col, row, xls));
            SetBordersXls(xls, 1, row, headers.Count, 1);
            return row + 1;
        }

        private static void SetBordersXls(Xls xls, int col, int row, int countCol, int countRow)
        {
            xls[col, row, countCol, countRow].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                               XlBordersIndex.xlEdgeBottom,
                                                                                               XlBordersIndex.xlEdgeLeft,
                                                                                               XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static Dictionary<string, int> OldHeaderByWidthRow = new Dictionary<string, int>()
        {
            { "Наименование устройств", 54 },
            { "Количество", 6 },
            { "Примечание", 12 },
            { "Срок обеспечения по заказу", 12 }
        };


        private static Dictionary<string, int> NewHeaderByWidthRow = new Dictionary<string, int>()
        {
            { "Наименование ПКИ", 54 },
            { "Обозначение ПКИ", 20 },
            { "Количество в КД, шт", 7 },
            { "Количество к приобретению, шт", 7},
            { "Срок обеспечения по заказу", 12 },
            { "Дата завершения работ по заказу", 12 },
            { "Инженер", 31 },
            { "Примечание", 25 }
        };

        private static Dictionary<string, int> CoastHeadersByWidthRow = new Dictionary<string, int>()
        {           
            { "Наименование сопоставленного объекта", 40 },
            { "Контрагент", 40 },
            { "Счет-фактура", 25 },
            { "Цена", 10 },
            { "Стоимость", 10 }
        };

        private void MakeSelectionByTypeReport(List<DeviceByGroups> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar, CoastData coastData = null)
        {
            progressBar.Maximum = reportData.Sum(t => t.Groups.Sum(k => k.Value.Count)) * 2;

            int errorId = 0;
            var xlsFiles = new List<Xls>();
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;

            try
            {
                // Запись данных по группам и устройствам
                foreach (var group in GroupsReferenceNames.supplyGroup.Keys)
                {
                    errorId = 1;
                    xls.AddWorksheet(group, true);
                    int row = 1;
                    int col = 1;
                    row = PasteHeadName(xls, reportParameters, col, row, 13, "Позаказная карточка");
                    row = PasteHeadName(xls, reportParameters, col, row, 13, string.Format("заказ № {0} наряд-заказ № {1}", reportParameters.OrderNumber, reportParameters.SpecificationOrder));
                    SetBordersXls(xls, 1, row - 2, 4, 2);

                    row = PasteHeader(xls, OldHeaderByWidthRow.Keys.ToList(), col, row);

                    foreach (var device in reportData)
                    {
                        var currentGroup = device.Groups.Where(t => t.Key == group).FirstOrDefault().Value;
                        if (currentGroup != null)
                        {
                            var deviceRow = row;
                            if (currentGroup.Count > 0)
                            {
                                WriteDeviceName(xls, row, device, OldHeaderByWidthRow.Count);
                                row++;
                            }

                            foreach (var obj in currentGroup.OrderBy(t => t.NomObject.ToString()))
                            {
                                WriteRow(xls, logDelegate, row, obj, deviceRow, progressBar);
                                row++;
                            }

                            if (currentGroup.Count > 0)
                            {
                                for (int j = 0; j < 2; j++)
                                {
                                    SetBordersXls(xls, 1, row, OldHeaderByWidthRow.Count, 1);
                                    row++;
                                }
                            }
                        }
                    }
                    PrintSetting(xls, OldHeaderByWidthRow.Values.ToList(), row, xls.Worksheet.Name, true);
                }
                errorId = 2;
                //==================Вставка вкладок с общей суммой для каждого ПКИ=============
                DeviceByGroups listSumPKI = new DeviceByGroups();
                foreach (var group in GroupsReferenceNames.supplyGroup.Keys)
                {
                    foreach (var device in reportData)
                    {
                        var currentGroup = device.Groups.FirstOrDefault(t => t.Key == group).Value;
                        if (currentGroup != null)
                        {
                            foreach (var obj in currentGroup.OrderBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()))
                            {
                                var currentNameGroup = listSumPKI.Groups.Keys.FirstOrDefault(t => t == group);
                                if (currentNameGroup == null)
                                {
                                    var newNomItem = new NomSpecificationItems(obj.NomObject, device.Count * obj.Count, obj.Remarks, obj.Engineer);
                                    var newNomItems = new List<NomSpecificationItems>();
                                    newNomItems.Add(newNomItem);
                                    listSumPKI.Groups.Add(group, newNomItems);
                                }
                                else
                                {
                                    var nomSpecItem = listSumPKI.Groups[currentNameGroup].Where(t => t.NomObject.SystemFields.Id == obj.NomObject.SystemFields.Id).FirstOrDefault();
                                    if (nomSpecItem != null)
                                    {
                                        nomSpecItem.AddRemarks(obj.Remarks);
                                        nomSpecItem.Count += device.Count * obj.Count;
                                    }
                                    else
                                    {
                                        NomSpecificationItems newNomItem = new NomSpecificationItems(obj.NomObject, device.Count * obj.Count, obj.Remarks, obj.Engineer);
                                        listSumPKI.Groups[currentNameGroup].Add(newNomItem);
                                    }
                                }
                            }
                        }
                    }
                }
                errorId = 3;
                if (reportParameters.IsResultManyFiles)
                {
                    xls.SelectWorksheet("Лист1");
                    xls.Worksheet.Delete();
                    xlsFiles.Add(xls);
                }
                errorId = 4;
                Dictionary<string, int> headers = new Dictionary<string, int>();
                if (reportParameters.IsResultManyFiles)
                {
                    NewHeaderByWidthRow.ToList().ForEach(t => headers.Add(t.Key, t.Value));
                }
                else
                {
                    OldHeaderByWidthRow.ToList().ForEach(t => headers.Add(t.Key, t.Value));
                }

                if (reportParameters.IsCreateObjectCost)
                {
                    foreach (var header in CoastHeadersByWidthRow)
                    {
                        headers.Add(header.Key, header.Value);
                    }
                }
              

                foreach (var group in GroupsReferenceNames.supplyGroup.Keys)
                {
                    errorId = 5;
                    var currentGroup = listSumPKI.Groups.Where(t => t.Key == group).FirstOrDefault().Value;
                    if (currentGroup == null || currentGroup.Count == 0)
                    {
                        continue;
                    }

                    if (reportParameters.IsResultManyFiles)
                    {
                        xls = new Xls();
                        xls.Application.DisplayAlerts = false;
                    }
                    xls.AddWorksheet(group + "(Сумма)", true);
                    int row = 1;
                    int col = 1;

                    if (!reportParameters.IsResultManyFiles)
                    {
                        row = PasteHeadName(xls, reportParameters, col, row, 13, "Позаказная карточка");
                        row = PasteHeadName(xls, reportParameters, col, row, 13, string.Format("заказ № {0} наряд-заказ № {1}", reportParameters.OrderNumber, reportParameters.SpecificationOrder));
                        SetBordersXls(xls, 1, row - 2, 4, 2);
                    }
                    row = PasteHeader(xls, headers.Keys.ToList(), col, row);

                    if (currentGroup != null)
                    {
                        try
                        {
                            if (group == GroupsReferenceNames.supplyGroup.Keys.ToList()[3])
                            {
                                row = WriteGroup3(xls, currentGroup, logDelegate, row, progressBar, coastData, reportParameters);
                            }
                            else
                            {
                                if (group == GroupsReferenceNames.supplyGroup.Keys.ToList()[2])
                                {
                                    row = WriteGroup2(xls, currentGroup, logDelegate, row, progressBar, coastData, reportParameters);
                                }
                                else
                                {

                                    if (group == GroupsReferenceNames.supplyGroup.Keys.ToList()[1])
                                    {
                                        row = WriteGroup1(xls, currentGroup, logDelegate, row, progressBar, coastData, reportParameters);
                                    }
                                    else
                                    {
                                        foreach (var obj in currentGroup.OrderBy(t => t.NomObject.ToString()))
                                        {
                                            row = WriteRow(xls, logDelegate, row, obj, progressBar, coastData, reportParameters);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("34 Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки " + errorId + "\r\n" + currentGroup.Count + "  " + listSumPKI.Device + "\r\n" + ex,
                                "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        for (int j = 0; j < 3; j++)
                        {
                            SetBordersXls(xls, 1, row + j, GetTableCountCol(reportParameters), 1);
                        }
                        row = row + 2;
                    }
                    var currentSheetName = xls.Worksheet.Name;
                    if (reportParameters.IsResultManyFiles)
                    {
                        AddEngineers(xls, row - 4, currentSheetName);
                        xls.SelectWorksheet("Лист1");
                        xls.Worksheet.Delete();
                        xlsFiles.Add(xls);
                    }
                    PrintSetting(xls, headers.Values.ToList(), row, currentSheetName, !reportParameters.IsResultManyFiles);
                }

                errorId = 6;
                if (!reportParameters.IsResultManyFiles)
                {
                    xls.SelectWorksheet("Лист1");
                    xls.Worksheet.Delete();
                    xlsFiles.Add(xls);
                }
                errorId = 7;
            }
            catch (Exception ex)
            {
                MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (var xlsFile in xlsFiles)
                {
                    errorId = 8;
                    xlsFile.Application.DisplayAlerts = true;
                    errorId = 9;
                    xlsFile.Visible = true;
                }
            }
        }

        private void WriteAllData(Xls xls, 
            DeviceByGroups listSumPKI, 
            List<DeviceByGroups> reportData,
            Dictionary<string, int> headers,
            ReportParameters reportParameters, 
            LogDelegate logDelegate, 
            ProgressBar progressBar, 
            CoastData coastData = null)
        {
            int errorId = 0;
            errorId = 5;
            var currentGroup = listSumPKI.Groups.FirstOrDefault().Value;
            if (currentGroup == null || currentGroup.Count == 0)
            {
                return;
            }

            if (reportParameters.IsResultManyFiles)
            {
                xls = new Xls();
                xls.Application.DisplayAlerts = false;
            }
            xls.AddWorksheet("Общая сумма", true);
            int row = 1;
            int col = 1;

            if (!reportParameters.IsResultManyFiles)
            {
                row = PasteHeadName(xls, reportParameters, col, row, 13, "Позаказная карточка");
                row = PasteHeadName(xls, reportParameters, col, row, 13, string.Format("заказ № {0} наряд-заказ № {1}", reportParameters.OrderNumber, reportParameters.SpecificationOrder));
                SetBordersXls(xls, 1, row - 2, 4, 2);
            }
            row = PasteHeader(xls, headers.Keys.ToList(), col, row);

            if (currentGroup != null)
            {
                try
                {
                    foreach (var obj in currentGroup.OrderBy(t => t.NomObject.ToString()))
                    {
                        row = WriteRow(xls, logDelegate, row, obj, progressBar, coastData, reportParameters);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("34 Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки " + errorId + "\r\n" + currentGroup.Count + "  " + listSumPKI.Device + "\r\n" + ex,
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                for (int j = 0; j < 3; j++)
                {
                    SetBordersXls(xls, 1, row + j, GetTableCountCol(reportParameters), 1);
                }
                row = row + 2;
            }
            var currentSheetName = xls.Worksheet.Name;
            PrintSetting(xls, headers.Values.ToList(), row, currentSheetName, !reportParameters.IsResultManyFiles);
        }
        private int WriteGroup1(Xls xls, List<NomSpecificationItems> currentGroup, LogDelegate logDelegate, int row, ProgressBar progressBar, CoastData coastData, ReportParameters reportParameters)
        {
            //Регулярное выражение для выявления номиналов 
            string namePattern = @"(?<=(\w{1,})\s{1,})\S{1,}\s{1,}\S{1,}";
          
            Regex regexName = new Regex(namePattern);

                 //Записываем объекты без типа в наименовании
            var sortCurrentGroupMC = currentGroup.Where(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString().ToLower().Contains("микросхема")).
                OrderBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).ToList();

            //Удаляем объекты без типа из списка
            foreach (var obj in sortCurrentGroupMC)
            {
                currentGroup.Remove(obj);
            }

            //Сортировка
            var sortCurrentGroup = currentGroup.OrderBy(t => regexName.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString()).
                 ThenBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).ToList();
          
            sortCurrentGroup.AddRange(sortCurrentGroupMC);

       //     Clipboard.SetText(string.Join("\r\n", sortCurrentGroup.Select(t => regexName.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString() + " | " + t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString())));
          //  MessageBox.Show(string.Join("\r\n", sortCurrentGroup.Select(t => regexName.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString() + " | " + t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString())));

            //Запись
            foreach (var obj in sortCurrentGroup)
            {
                row = WriteRow(xls, logDelegate, row, obj, progressBar, coastData, reportParameters);
            }
            return row;
        }


        private int WriteGroup2(Xls xls, List<NomSpecificationItems> currentGroup, LogDelegate logDelegate, int row, ProgressBar progressBar, CoastData coastData, ReportParameters reportParameters)
        {
            //Регулярное выражение для выявления номиналов 
            string nomianalPattern = @"(?<=-)((\d+)[.,]?\d*)(?=\s?(пФ|мкФ|Ом|кОм|МОм|/))";
            Regex regexNominal = new Regex(nomianalPattern);

            //Регулярное выражение для выявление типов
            string typePattern = @"(?<=(Конденсатор|Резистор)\s).*(?=-\d+[.,/]?\d*\s?(пФ|мкФ|Ом|кОм|МОм))";
            Regex regexType = new Regex(typePattern);

            //Регулярное выражение для выделения размерности
            string sizePattern = @"(мкФ|пФ|Ом|кОм|МОм)";
            Regex regexSize = new Regex(sizePattern);

            //Записываем объекты без типа в наименовании
            var sortCurrentGroup = currentGroup.Where(t => !regexType.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Success).
                OrderBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).ToList();

            //Удаляем объекты без типа из списка
            foreach (var obj in sortCurrentGroup)
            {
                currentGroup.Remove(obj);
            }

            //Сортировка
            sortCurrentGroup.AddRange(currentGroup.OrderBy(t => regexType.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString()).
                ThenBy(t => SortBySize(regexSize.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString())).
                ThenBy(t => Convert.ToDouble(regexNominal.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString().
                Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator).
                Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator))).ToList());

            string prewType = string.Empty;
            //Запись
            foreach (var obj in sortCurrentGroup)
            {
                obj.Type = regexType.Match(obj.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString();
                if (prewType != obj.Type && !reportParameters.IsResultManyFiles)
                {
                    WriteTypeGroup(xls, row, obj, reportParameters);
                    row++;
                }
                row = WriteRow(xls, logDelegate, row, obj, progressBar, coastData, reportParameters);
                prewType = obj.Type;
            }
            return row;
        }

        public int WriteGroup3(Xls xls, List<NomSpecificationItems> currentGroup, LogDelegate logDelegate, int row, ProgressBar progressBar, CoastData coastData, ReportParameters reportParameters)
        {
            //Регулярное выражение для выявления номиналов 
            string nomianalPattern = @"(?<=-)((\d+)[.,]?\d*)(?=\s?(пФ|мкФ|Ом|кОм|МОм|/))";
            Regex regexNominal = new Regex(nomianalPattern);

            //Регулярное выражение для выявление типов
            string typePattern = @"(?<=(Резистор|Блок|Набор резисторов|Терморезистор|Поглотитель|Варистор)\s).*(?=-\d+[.,/]?\d*\s?(пФ|мкФ|Ом|кОм|МОм))";
            Regex regexType = new Regex(typePattern);

            //Регулярное выражение для выделения размерности
            string sizePattern = @"(мкФ|пФ|Ом|кОм|МОм)";
            Regex regexSize = new Regex(sizePattern);

            //Регулярное выражение для выделения буквы А,Б,В,С
            string letterPattern = @"(?<=(-1,0-|-1.0-))(А|A|Б|В|B|C|С)\s";
            Regex regexLetter = new Regex(letterPattern);

            //Записываем объекты без типа в наименовании
            var sortCurrentGroup = currentGroup.Where(t => !regexType.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Success).
                OrderBy(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).ToList();

            //Удаляем объекты без типа из списка
            foreach (var obj in sortCurrentGroup)
            {
                currentGroup.Remove(obj);
            }

            //Сортировка
            sortCurrentGroup.AddRange(currentGroup.OrderBy(t => regexType.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString()).
                ThenBy(t => SortByLetter(regexLetter.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString().Trim())).
                ThenBy(t => SortBySize(regexSize.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString())).            
                ThenBy(t => Convert.ToDouble(regexNominal.Match(t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString().
                Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator).
                Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator))).ToList());

            string prewType = string.Empty;
            //Запись
            foreach (var obj in sortCurrentGroup)
            {
                obj.Type = regexType.Match(obj.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()).Value.ToString();

                var matchLetter = regexLetter.Match(obj.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString());
                if (matchLetter.Success)
                    obj.Letter = matchLetter.Value.ToString();

                if (prewType != obj.Type + obj.Letter && !reportParameters.IsResultManyFiles)
                {
                    WriteTypeGroup(xls, row, obj, reportParameters);
                    row++;
                }
                row = WriteRow(xls, logDelegate, row, obj, progressBar, coastData, reportParameters);
                prewType = obj.Type + obj.Letter;
            }

            return row;
        }

        public int SortBySize(string size)
        {
            return size.ToLower().Trim() switch
            {
                "oм" => 0,
                "ом" => 0,
                "om" => 0,
                "оm" => 0,
                "ком" => 1,
                "мом" => 2,
                "пф" => 1,
                "мкф" => 2,
                _ => 0
            };
        }

        public int SortByLetter(string letter)
        {
            return letter switch
            {
                "c" => 1,
                "с" => 1,
                _ => 0
            };
        }

        private static void PrintSetting(Xls xls, List<int> widthColumns, int row, string sheetName, bool printCommunication = false)
        {
            xls.SelectWorksheet(sheetName);
            for (int i = 1; i <= widthColumns.Count; i++)
            {
                xls.SetColumnWidth(i, widthColumns[i - 1]);
            }
            xls[1, 1, widthColumns.Count, row].WrapText = true;

            if (printCommunication)
            {
                xls.Worksheet.PageSetup.LeftMargin = xls.Application.InchesToPoints(0.984251968503937);
                xls.Worksheet.PageSetup.RightMargin = xls.Application.InchesToPoints(0.393700787401575);
                xls.Worksheet.Application.PrintCommunication = true;
            }
            if (!printCommunication)
            {
                xls.Worksheet.Cells.Locked = false;
                xls[1, 1, 3, row].Locked = true;
                xls.Worksheet.Protect("омтс", true, true, true,
                    Type.Missing, Type.Missing, true, true, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            }
        }
        private void WriteRow(Xls xls, LogDelegate logDelegate, int row, NomSpecificationItems data, int deviceRow, ProgressBar progressBar)
        {
            var beginRow = row;
            var name = data.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = data.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            if (denotation != string.Empty)
                name += " " + denotation;

            logDelegate(String.Format("Запись объекта: " + name));
            xls[1, row].SetValue(name.Trim());

            xls[2, row].SetFormula("=ПРОИЗВЕД(" + xls[2, deviceRow].Address + ";" + data.Count + ")");
            if (data.CountError)
                xls[2, row].Font.Color = Color.Red;

            xls[3, row].SetValue(string.Join("\r\n", data.Remarks.Where(t => t != string.Empty)));
            xls[1, row].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[2, row].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[3, row].VerticalAlignment = XlVAlign.xlVAlignTop;

            if (data.Error)
            {
                xls[1, row, 4, 1].Font.Color = Color.Red;
            }
            progressBar.PerformStep();
     
            SetBordersXls(xls, 1, row, 4, 1);
        }

        private static int WriteRow(Xls xls, LogDelegate logDelegate, int row, NomSpecificationItems data, ProgressBar progressBar, CoastData coastData, ReportParameters reportParameters)
        {
            var name = data.NomObject[Guids.NomenclatureParameters.NameGuid].GetString().Trim();
            var denotation = data.NomObject[Guids.NomenclatureParameters.DenotationGuid].GetString().Trim();
            var fullName = (name + " " + denotation).Trim();
            var remarks = string.Join("\r\n", data.Remarks.Where(t => t != string.Empty));
            logDelegate(String.Format("Запись объекта: " + fullName));
            var col = 1;
            var colCount = 1;

            if (reportParameters.IsResultManyFiles)
            {
                xls[1, row].SetValue(name);
                xls[2, row].SetValue(denotation);
                xls[3, row].SetValue(data.Count);
                xls[7, row].SetValue(data.Engineer);
                xls[8, row].SetValue(remarks);
                col = 8;
                colCount = 3;
            }
            else
            {
                xls[1, row].SetValue(fullName);
                xls[2, row].SetValue(data.Count);
                xls[3, row].SetValue(remarks);
                col = 4;
                colCount = 2;
            }
           
            xls[1, row, col, 1].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[colCount, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            if (data.Error)
            {
                for (int j = 1; j < col; j++)
                {
                    xls[j, row].Font.Color = Color.Red;
                }
            }
            if (data.CountError)
            {
                xls[colCount, row].Font.Color = Color.Red;
            }
            progressBar.PerformStep();
            var rowAfterAddedCoast = WriteCoastData(xls, col, colCount, row, data, coastData, reportParameters);
            rowAfterAddedCoast++;
            return rowAfterAddedCoast;
        }

        private static int GetTableCountCol(ReportParameters reportParameters)
        {
            var defaultCol = reportParameters.IsResultManyFiles ? NewHeaderByWidthRow.Count : OldHeaderByWidthRow.Count; 
            var colCount = reportParameters.IsCreateObjectCost ? defaultCol + CoastHeadersByWidthRow.Count : defaultCol;
            return colCount;
        }

        private static int WriteCoastData(Xls xls, int col, int colCount, int row, NomSpecificationItems data, CoastData coastData, ReportParameters reportParameters)
        {
            col++;
            if (reportParameters.IsCreateObjectCost && coastData != null)
            {
                var objectCoast = coastData.Objects.FirstOrDefault(t => t.NomObj.SystemFields.Id == data.NomObject.SystemFields.Id);
                if (objectCoast != null && objectCoast.StoreObjects.Count > 0)
                {
                    for (int i = 0; i < objectCoast.StoreObjects.Count; i++)
                    {
                        xls[col, row + i].SetValue(objectCoast.StoreObjects[i].Name);
                        var invoice = objectCoast.StoreObjects[i].InvoiceObject;
                        if (invoice != null)
                        {
                            xls[col + 1, row + i].SetValue(invoice.ContractorInvoice);
                            xls[col + 2, row + i].SetValue(invoice.Number + " от " + invoice.Date.ToShortDateString());
                            var pki = invoice.pkiObjects.FirstOrDefault(t => t.StoreObject != null && t.StoreObject.SystemFields.Id == objectCoast.StoreObjects[i].RefObj.SystemFields.Id);
                            if (pki != null)
                            {
                                xls[col + 3, row + i].SetValue(pki.Price);
                                xls[col + 4, row + i].SetFormula("=ПРОИЗВЕД(" + xls[colCount, row + i].Address + ";" + xls[col + 3, row + i].Address + ")");
                            }
                        }
                    }
                    SetBordersXls(xls, 1, row, col + 4, objectCoast.StoreObjects.Count);
                    if (objectCoast.StoreObjects.Count > 1)
                    {
                        xls[1, row, col + 4, objectCoast.StoreObjects.Count].Font.Color = Color.Red;
                    }
                    xls[col, row, col, objectCoast.StoreObjects.Count].VerticalAlignment = XlVAlign.xlVAlignTop;
                    return row + objectCoast.StoreObjects.Count - 1;
                }
                else
                {
                    SetBordersXls(xls, 1, row, GetTableCountCol(reportParameters), 1);
                }
            }
            else
            {
                SetBordersXls(xls, 1, row, GetTableCountCol(reportParameters), 1);
            }

            return row;
        }

        private static void WriteDeviceName(Xls xls, int row, DeviceByGroups device, int countColumn)
        {
            var name = device.Device[Guids.NomenclatureParameters.NameGuid].GetString().Trim();
            var denotation = device.Device[Guids.NomenclatureParameters.DenotationGuid].GetString().Trim();
            var fullName = name + " " + denotation;

            if (name.Trim().ToLower() != device.NameLoadList.Trim().ToLower())
            {
                name = device.NameLoadList;
            }

            xls[1, row].SetValue(!String.IsNullOrEmpty(device.Remark) && name.ToLower().Contains("кабель") ? fullName + " " + device.Remark : fullName);  
            xls[2, row].SetValue(device.Count.ToString());
          
            for (int i = 1; i < 5; i++)
            {
                xls[i, row].Font.Bold = true;
                xls[i, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            }
            SetBordersXls(xls, 1, row, countColumn, 1);
        }

        private static void WriteTypeGroup(Xls xls, int row, NomSpecificationItems item, ReportParameters reportParameters)
        {
            var type = item.Type;

            if (item.Letter != string.Empty)
            {
                type += " с буквой " + item.Letter;
            }

            xls[1, row].SetValue(type);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            SetBordersXls(xls, 1, row, GetTableCountCol(reportParameters), 1);
        }

        private List<DeviceByGroups> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");
            List<DeviceByGroups> devices = new List<DeviceByGroups>();

            if (reportParameters.AddCodeDevices)
            {
                var addedDevices = new List<NomSpecificationItems>();

                foreach (var nomObj in reportParameters.NomObjects)
                {
                    logDelegate(String.Format("Поиск шифрованных устройств в : " + nomObj.NomObject));

                    var namesParents = new List<string>();

                    List<NomSpecificationItems> codeChildren = new List<NomSpecificationItems>();
                    codeChildren.AddRange(GetAssemblyChildren(reportParameters, nomObj.NomObject, nomObj.Count, nomObj.Number, namesParents));
                
                    if (codeChildren != null)
                    {                       
                        foreach (var obj in codeChildren)
                        {
                            if (reportParameters.AddAllCodeDevice)
                            {
                                if (addedDevices.Count(t => t.NomObject == obj.NomObject) == 0)
                                {         
                                    addedDevices.Add(obj);
                                }
                                else
                                {
                                    addedDevices.Where(t => t.NomObject == obj.NomObject).First().Count += obj.Count;
                                }
                            }
                            else
                            {
                                if (reportParameters.NomObjects.Count(t => t.NomObject == obj.NomObject) == 0)
                                {
                                    if (addedDevices.Count(t => t.NomObject == obj.NomObject) == 0)
                                    {
                                        addedDevices.Add(obj);
                                    }
                                }
                            }
                        }
                    }
                }

                if (addedDevices.Count > 0)
                {
                    reportParameters.NomObjects.AddRange(addedDevices.OrderBy(t => t.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()));
                }
            }

            foreach (var nomObj in reportParameters.NomObjects)
            {
                logDelegate(String.Format("Загрузка объектов из СЕ: " + nomObj.NomObject));
                var device = new DeviceByGroups(reportParameters, nomObj, logDelegate);
          
                devices.Add(device);
            }

            foreach (var device in devices)
            {
                var parents = device.Device.Parents.Where(t => devices.Count(r => r.Device == t) > 0).ToList();
                if (parents.Count > 0)
                {
                    List<string> remarks = new List<string>();
                    foreach (var parent in parents)
                    {
                        var hLink = parent.GetChildLink(device.Device); //Иерархия номенклатуры    
                        var remark = hLink[NomenclatureHierarchyLink.FieldKeys.Remarks].Value.ToString();
                        if (remark != string.Empty)
                             remarks.Add(remark);                
                    }
                    if (remarks.Count > 0)
                        device.Remark = string.Join("\r\n", remarks);
                }
            }

            FindDublicates(devices);

            return devices;
        }

        static private void FindDublicates(List<DeviceByGroups> data)
        {
            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (String.Equals(data[mainID].Device, data[slaveID].Device))
                    {
                        data[mainID].Count += data[slaveID].Count;
                        data.RemoveAt(slaveID);                      
                        slaveID--;
                    }
                }
            }
        }

        public List<NomSpecificationItems> GetAssemblyChildren(ReportParameters reportParameters, NomenclatureHierarchyLink childHlinkObject, NomenclatureObject parentObject, double countParent, string number, List<string> namesParents)
        {
            var deviceRegex = new Regex(@"(\d{1}[А-Я]{1}\d{3}(-\d{1,})?)");
            int errorId = 0;
       
            var listObjects = new List<NomSpecificationItems>();
            try
            {
                errorId = 11;
                var nameChild = childHlinkObject.ChildObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
                double amount = childHlinkObject.Amount * countParent;

                if ((childHlinkObject.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Assembly) || childHlinkObject.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Komplekt)) &&
                    //  childObject[Guids.NomenclatureParameters.CodeGuid].Value.ToString().Trim() != string.Empty && 
                    (reportParameters.NomObjects.Select(t => t.NomObject).Contains(childHlinkObject.ChildObject) || deviceRegex.Match(nameChild).Success) &&
                    !nameChild.ToLower().Contains("упаковка") &&
                    !nameChild.ToLower().Contains("документ") &&
                    parentObject != null)
                {
                    errorId = 1;
                    
                    if (parentObject != null)
                    {
                        errorId = 2;

                      //  amount = Convert.ToDouble(childHlinkObject[NomenclatureHierarchyLink.FieldKeys.Amount].Value) * countParent;
                      //  MessageBox.Show(childHlinkObject + " - " + childHlinkObject[NomenclatureHierarchyLink.FieldKeys.Amount].Value + " * " + countParent);
                    }
                    errorId = 3;

                    if (listObjects.Count(t => t.NomObject == childHlinkObject.ChildObject) == 0)
                    {
                        errorId = 4;
                        var codeDevices = new NomSpecificationItems((NomenclatureObject)childHlinkObject.ChildObject, amount, string.Empty, string.Empty, number);
                        codeDevices.NameParents = string.Join(" -> ", namesParents);
                        listObjects.Add(codeDevices);
                    }
                }
                errorId = 5;

                var childrenChildObject = childHlinkObject.ChildObject.Children;
                if (childrenChildObject.Count() > 0)
                {
                    errorId = 6;
                    namesParents.Add(childHlinkObject.ChildObject.ToString());
                }

                foreach (NomenclatureHierarchyLink childNHL in childrenChildObject.GetHierarchyLinks())
                {
                    errorId = 7;

                   // MessageBox.Show("countParent " + childHlinkObject + " - " + childHlinkObject[NomenclatureHierarchyLink.FieldKeys.Amount].Value + " * " + countParent);

                    if (parentObject != null)
                    {
                        errorId = 8;
                 
                     //   amount = Convert.ToDouble(childHlinkObject[NomenclatureHierarchyLink.FieldKeys.Amount].Value) * countParent;

                   //     MessageBox.Show("countParent " + childHlinkObject + " - " + childHlinkObject[NomenclatureHierarchyLink.FieldKeys.Amount].Value + " * " + countParent);
                    }
                    errorId = 9;
                    List<string> newNamesParents = new List<string>();
                    newNamesParents.AddRange(namesParents);

                    var children = GetAssemblyChildren(reportParameters, childNHL, (NomenclatureObject) childHlinkObject.ChildObject, amount, number, newNamesParents);
                    errorId = 10;
                    if (children == null)
                    {
                        continue;
                    }
                    foreach (var childParent in children)
                    {
                        errorId = 11;
                        if (listObjects.Count(t => t.NomObject == childParent.NomObject) == 0)
                        {
                        //    MessageBox.Show(childParent.ToString() + " " + childParent.Count);
                            listObjects.Add(childParent);
                        }
                        else
                        {
                            listObjects.Where(t => t.NomObject == childParent.NomObject).FirstOrDefault().Count += childParent.Count;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("99. Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return listObjects;
        }

        public List<NomSpecificationItems> GetAssemblyChildren(ReportParameters reportParameters, NomenclatureObject childObject, double countParent, string number, List<string> namesParents)
        {
            var deviceRegex = new Regex(@"(\d{1}[А-Я]{1}\d{3}(-\d{1,})?)");

            var listObjects = new List<NomSpecificationItems>();
            try
            {
                var childrenChildObject = childObject.Children;
                if (childrenChildObject.Count() > 0)
                {
                    namesParents.Add(childObject.ToString());
                }

                foreach (NomenclatureHierarchyLink childNHL in childrenChildObject.GetHierarchyLinks().Where(t=>t.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Assembly) ||
                                                                                                                t.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Komplekt)))
                {
                  //  double amount = childNHL.Amount;
                    List<string> newNamesParents = new List<string>();
                    newNamesParents.AddRange(namesParents);

                    var children = GetAssemblyChildren(reportParameters, childNHL, childObject, countParent, number, newNamesParents);

                    if (children == null)
                    {
                        continue;
                    }
                    foreach (var childParent in children)
                    {
                        if (listObjects.Count(t => t.NomObject == childParent.NomObject) == 0)
                        {
                            listObjects.Add(childParent);
                        }
                        else
                        {
                            listObjects.Where(t => t.NomObject == childParent.NomObject).FirstOrDefault().Count += childParent.Count;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("9. Возникла непредвиденная ошибка. Обратитесь в отдел 911." + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return listObjects;
        }
        private void AddEngineers(Xls xls, int countRows, string sheetName)
        {
            xls.AddWorksheet("Инженеры", true);
            ReferenceInfo usersReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Users);
            Reference usersReference = usersReferenceInfo.CreateReference();
            var engineers = usersReference.Find(Guids.GroupAndUsers.РольИнженерОМТС).Children;

            for (int i = 0; i < engineers.Count(); i++)
            {
                xls[1, i + 1].Value = engineers[i].ToString();
            }
            xls.SelectWorksheet(sheetName);
            xls[7, 2, 1, countRows].Validation.Delete();
            xls[7, 2, 1, countRows].Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, "=Инженеры!$A$1:$A$" + engineers.Count());
            xls[7, 2, 1, countRows].Validation.InCellDropdown = true;
        }
    }
}
