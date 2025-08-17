using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Common;
using TFlex.Reporting;
using static ObjectsOIWithCoastStore.Guids;

namespace ObjectsOIWithCoastStore
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
                    StoreOIReport.MakeReport(context, selectionForm.reportParameters);
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
    public class StoreOIReport
    {
        public IndicationForm m_form = new IndicationForm();
        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            StoreOIReport report = new StoreOIReport();
            LogDelegate m_writeToLog;

            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
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

            var reportNomData = new List<NomObject>();
            var reportListDevicesData = new Dictionary<DeviceWithKoef, List<NomObject>>();
            var reportStoreData = new List<StoreObject>();
            if (reportParameters.SelectedDevice)
            {
                reportNomData = ReadData(reportParameters.Device, 1, m_writeToLog, m_form.progressBar);

                if (reportNomData == null)
                {
                    m_form.Close();
                    return;
                }
            }
            else
            {
                if (reportParameters.SelectedStoreObjects || reportParameters.ObjectsReferenceFailure)
                {
                    reportStoreData.AddRange(reportParameters.StoreObjects);
                    if (reportStoreData == null)
                    {
                        m_form.Close();
                        return;
                    }
                }
                else
                {
                    if (reportParameters.SelectedDevices)
                    {
                        foreach (var device in reportParameters.Devices)
                        {
                            reportNomData.AddRange(ReadData(device.Key, device.Value, m_writeToLog, m_form.progressBar));
                            FindDublicates(reportNomData);
                        }

                        if (reportNomData == null)
                        {
                            m_form.Close();
                            return;
                        }
                    }
                    else
                    {
                        if (reportParameters.SelectedDevicesWithKoef)
                        {
                            foreach (var device in reportParameters.DevicesWithKoef)
                            {
                                var nomObjects = ReadData(device.NomObj, 1, m_writeToLog, m_form.progressBar);
                                reportListDevicesData.Add(device, nomObjects);
                            }

                            if (reportNomData == null)
                            {
                                m_form.Close();
                                return;
                            }
                        }
                        else
                        {
                            if (reportParameters.SelectedNomObjects)
                            {
                                reportNomData.AddRange(reportParameters.NomObjects);
                            }
                            else
                            { 
                                if (reportParameters.SelectedDevicesDifferentFiles)
                                {
                                    int i = 0;
                                    var xlsCost = new Xls();
                                    xlsCost.AddWorksheet("ПКИ", false);
                                    xlsCost.SaveAs(reportParameters.SelectedPath + @"\Итог.xlsx");
                                    int row = 1;

                                    foreach (var device in reportParameters.Devices)
                                    {
                                        i++;

                                        reportNomData = ReadData(device.Key, device.Value, m_writeToLog, m_form.progressBar);

                                        if (reportNomData == null || reportNomData.Count == 0)
                                        {
                                            foreach (var str in reportParameters.DevicesByRow.Where(t => t.NomObj.SystemFields.Id == device.Key.SystemFields.Id).ToList())
                                            {
                                                xlsCost[1, str.Row].Value = device.Key.ToString();
                                                xlsCost[2, str.Row].Value = str.Count;
                                                xlsCost[3, str.Row].Value = 0;
                                                xlsCost[3, str.Row].NumberFormat = @"#,##0.00";
                                                var formula = "=ПРОИЗВЕД(" + xls[2, str.Row].Address + ";" + xls[3, str.Row].Address + ")";
                                                xlsCost[4, str.Row].SetFormula(formula);
                                                xlsCost[4, str.Row].NumberFormat = @"#,##0.00";
                                            }
                                            continue;
                                        }
                                        m_form.Text = "Обработка " + i + " из " + reportParameters.Devices.Count;
                                        //  var cost = MakeReportSeveralFiles(xlsCost, m_writeToLog, reportNomData, reportParameters, device.Key);

                                        m_writeToLog("Запись отчета для " + device);
                                        m_form.progressBar.Value = 0;
                                        xlsCost.SelectWorksheet("ПКИ");
                                        //Формирование отчета
                                        row = MakeSelectionByTypeReport(xlsCost, reportNomData.Where(t => t.NomObj.Class.IsInherit(Guids.NomenclatureTypes.OtherItem)).OrderBy(t => t.Name).ToList(),
                                            reportParameters, m_writeToLog, m_form.progressBar, device.Key, row);
                                      

                                        reportNomData = new List<NomObject>();
                                        xlsCost.SelectWorksheet("Лист1");

                                        var strDevices = reportParameters.DevicesByRow.Where(t => t.NomObj.SystemFields.Id == device.Key.SystemFields.Id).ToList();

                                        foreach (var str in strDevices)
                                        {
                                            xlsCost[1, str.Row].Value = device.Key.ToString();
                                            xlsCost[2, str.Row].Value = str.Count;
                                            xlsCost[3, str.Row].Value = @"='ПКИ'!" + xls[8,row].Address;
                                            xlsCost[3, str.Row].NumberFormat = @"#,##0.00";
                                            var formula = "=ПРОИЗВЕД(" + xls[2, str.Row].Address + ";" + xls[3, str.Row].Address + ")";
                                            xlsCost[4, str.Row].SetFormula(formula);
                                            xlsCost[4, str.Row].NumberFormat = @"#,##0.00";
                                        }

                                        row++;
                                        row++;
                                        xlsCost.Save();
                                    }

                                    xlsCost.Application.DisplayAlerts = true;
                                    xlsCost.Save();
                                    xlsCost.Close();

                                    m_form.Dispose();
                                    return;
                                }
                            }
                        }
                    }
                }
            }

            try
            {
                xls.Application.DisplayAlerts = false;
                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                if (reportParameters.SelectedDevice || reportParameters.SelectedDevices || reportParameters.SelectedNomObjects)
                {
                    //Формирование отчета
                    MakeSelectionByTypeReport(xls, reportNomData.Where(t => t.NomObj.Class.IsInherit(Guids.NomenclatureTypes.OtherItem)).OrderBy(t => t.Name).ToList(),
                        reportParameters, m_writeToLog, m_form.progressBar, null, 1);
                }
                else
                {
                    if (reportParameters.SelectedDevicesWithKoef)
                        MakeSelectionByTypeReport(xls, reportListDevicesData, reportParameters, m_writeToLog, m_form.progressBar);
                    else
                        MakeSelectionByTypeReport(xls, reportStoreData, reportParameters, m_writeToLog, m_form.progressBar);
                }

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

        private double MakeReportSeveralFiles(Xls xls, LogDelegate m_writeToLog, List<NomObject> reportNomData, ReportParameters reportParameters, ReferenceObject device)
        {
            double cost = 0;
           

            return cost;
        }

        static private void FindDublicates(List<NomObject> data)
        {
            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (String.Equals(data[mainID].NomObj.SystemFields.Id, data[slaveID].NomObj.SystemFields.Id))
                    {
                        data[mainID].Count += data[slaveID].Count;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
        }

        public Dictionary<int, Invoice> invoicesBySCHFK = new Dictionary<int, Invoice>();

        public Invoice SearchInvoice(StoreObject storeObj, Reference invoiceReference,  ReferenceInfo invoiceReferenceInfo)
        {
            List<Invoice> invoices = new List<Invoice>();
            if (invoicesBySCHFK.ContainsKey(storeObj.RefObj.SystemFields.Id))
            {
                var schfk = invoicesBySCHFK[storeObj.RefObj.SystemFields.Id];

                if (schfk != null)
                {
                    return schfk;
                }
            }
            
            var filter = new TFlex.DOCs.Model.Search.Filter(invoiceReferenceInfo);

            filter.Terms.AddTerm("[Поставленные изделия].[ПКИ]->[ID]", ComparisonOperator.Equal, storeObj.RefObj.SystemFields.Id);
            var refObjects = invoiceReference.Find(filter).ToList();
           // foreach (var refObj in refObjects)
             //   invoices.Add(new Invoice(refObj));

            invoices.AddRange(refObjects.Select(t=>new Invoice(t)).ToList());

            if (invoices.Count == 0)
                return null;
            else
            {
                var findedInvoice = invoices.OrderByDescending(t => t.Date).FirstOrDefault();
                invoicesBySCHFK.Add(storeObj.RefObj.SystemFields.Id, findedInvoice);

                return findedInvoice;
            }
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
            col = InsertHeader("№", col, row, xls);

            if (reportParameters.ObjectsReferenceFailure)
            {
                col = InsertHeader("СЗ", col, row, xls);
            }
            else
                col = InsertHeader("Наименование T-FLEX DOCs", col, row, xls);

            col = InsertHeader("Наименование Парус", col, row, xls);
            col = InsertHeader("Номенклатурный номер", col, row, xls);
            col = InsertHeader("Ед. изм.", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Цена (руб.) без НДС", col, row, xls);
            col = InsertHeader("Сумма (руб.) без НДС", col, row, xls);
            col = InsertHeader("Обоснование", col, row, xls);
            col = InsertHeader("Наименование поставщика Парус", col, row, xls);
            col = InsertHeader("Наименование поставщика T-FLEX DOCs", col, row, xls);
            col = InsertHeader("ИНН", col, row, xls);
            xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            return row;
        }

        public int PasteHeaderListDevices(Xls xls, int col, int row, bool addedCol)
        {
            col = InsertHeader("№", col, row, xls);
            col = InsertHeader("Наименование", col, row, xls);
            col = InsertHeader("Обозначение", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            if (addedCol)
            {
                col = InsertHeader("Коэффициент", col, row, xls);
                col = InsertHeader("Сумма", col, row, xls);
                xls[1, row, 6, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
            else
                xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            row++;
            return row;
        }

        public List<NomObject> GetStoreData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            List<NomObject> objects = new List<NomObject>();
            return objects;
        }

        public List<NomObject> ReadData(NomenclatureObject nomObject, int summaryCount, LogDelegate logDelegate, ProgressBar progressBar)
        {
            List<NomObject> objects = new List<NomObject>();

            var nomReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomReferenceInfo.CreateReference();

            var hierarchyLinks = nomObject.Children.RecursiveLoadHierarchyLinks();
            var hierarchy = new Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>>();
            var objectsDictionary = new Dictionary<int, NomenclatureObject>();
            objectsDictionary[nomObject.SystemFields.Id] = nomObject;

            var корневыеОбъектыНеВключаемыйВРасчет = hierarchyLinks.Where(t => (bool)t[NomHierarchyLink.Fields.NotIncludeCalculateSum]).ToList();
            var объектыНеВключаемыйВРасчет = new List<ComplexHierarchyLink>();
            корневыеОбъектыНеВключаемыйВРасчет.ForEach(t => объектыНеВключаемыйВРасчет.AddRange(t.ChildObject.Children.RecursiveLoadHierarchyLinks()));

            foreach (var hierarchyLink in hierarchyLinks.Where(t => !объектыНеВключаемыйВРасчет.Contains(t)))
            {
                var nhl = (NomenclatureHierarchyLink)hierarchyLink;
                objectsDictionary[hierarchyLink.ChildObjectId] = (NomenclatureObject)hierarchyLink.ChildObject;

                Dictionary<Guid, ComplexHierarchyLink> parentHierarchy = null;
                if (!hierarchy.TryGetValue(hierarchyLink.ParentObjectId, out parentHierarchy))
                {
                    parentHierarchy = new Dictionary<Guid, ComplexHierarchyLink>();
                    hierarchy[hierarchyLink.ParentObjectId] = parentHierarchy;
                }
                parentHierarchy[hierarchyLink.Guid] = hierarchyLink;
            }

            var countObjects = new Dictionary<int, double>();
            CalculateCount(nomObject.SystemFields.Id, 1, hierarchy, countObjects);

            var objectTypes = new List<Guid>();
            objectTypes.Add(Guids.NomenclatureTypes.OtherItem);

            foreach (var nomObj in objectsDictionary.Values.Where(t => objectTypes.Any(t1 => t.Class.IsInherit(t1)))
                .OrderBy(t => t.Denotation.GetString()).ThenBy(t => t.Name.GetString()))
            {
                var doc = new NomObject(nomObj);
                doc.Count = countObjects[nomObj.SystemFields.Id];
                objects.Add(doc);
            }

            List<NomObject> objectsWithCount = new List<NomObject>();

            foreach (var obj in objects)
            {
                var newObj = obj;
                newObj.Count = newObj.Count * summaryCount;
                objectsWithCount.Add(newObj);
            }

            return objectsWithCount;
        }

        public static void CalculateCount(int parentId, double count,
         Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>> hierarchy,
         Dictionary<int, double> countObjects)
        {
            if (countObjects.ContainsKey(parentId))
            {
                countObjects[parentId] += count;
            }
            else
            {
                countObjects[parentId] = count;
            }

            if (hierarchy.ContainsKey(parentId))
            {
                foreach (var link in hierarchy[parentId].Values)
                {
                    var remark = link[NomHierarchyLink.Fields.Comment].Value.ToString();
                    var countNHL = Convert.ToDouble(link[NomHierarchyLink.Fields.Count].Value);
                    if (countNHL != Math.Truncate(countNHL))
                    {
                        countNHL = CountSelectionRezistor((NomenclatureObject)link.ChildObject, remark);
                    }
                    var linkCount = countNHL * count;
                    CalculateCount(link.ChildObjectId, linkCount, hierarchy, countObjects);
                }
            }
        }

        private static int CountSelectionRezistor(NomenclatureObject nomObj, string remark)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            int errorId = 0;
            Regex regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                errorId = 1;
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    errorId = 2;
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value + " " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    errorId = 3;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                    errorId = 4;
                }
                errorId = 5;
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
                errorId = 6;
            }
            catch
            {
                MessageBox.Show("Обратитесь в отдел 911. Сформированный отчет будет некорректен. Код ошибки 7. Не могу посчитать количество резистора\r\n" + nomObj + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId + "\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            return countRezist;
        }

        private int MakeSelectionByTypeReport(Xls xls, List<NomObject> reportData,  ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar, ReferenceObject device, int row)
        {
            //Задание начальных условий для progressBar
  
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(reportData.Count));
            if (progressBar.Maximum < (reportData.Count))
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
            if (reportParameters.SelectedDevice)
            {
                PasteHeadName(xls, reportParameters, col, row, 14, reportParameters.Device.ToString());
            }
            else
            {
                if (reportParameters.SelectedDevices)
                {
                    PasteHeadName(xls, reportParameters, col, row, 14, "Расчет стоимости ПКИ для списка устройств " + device);
                    row++;
                    row = PasteHeaderListDevices(xls, col, row, false);
                    for (int i = 0; i < reportParameters.Devices.Count; i++)
                    {
                        col = 1;
                        xls[col, row].SetValue(i + 1);
                        col++;
                        xls[col, row].SetValue(reportParameters.Devices.Keys.ToList()[i][Guids.NomenclatureParameters.NameGuid].Value.ToString());
                        col++;
                        xls[col, row].SetValue(reportParameters.Devices.Keys.ToList()[i][Guids.NomenclatureParameters.DenotationGuid].Value.ToString());
                        col++;
                        xls[col, row].SetValue(reportParameters.Devices.Values.ToList()[i]);
                        xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                        row++;
                    }
                }
                else
                {
                    if (reportParameters.SelectedDevicesDifferentFiles)
                    {
                        PasteHeadName(xls, reportParameters, col, row, 14,device.ToString());
                    }
                }
            }
            row++;
            col = 1;
            row = PasteHeader(xls, col, row, reportParameters);
            var beginRow = row;
            int rowCount = 1;

            var invoiceReferenceInfo = ReferenceCatalog.FindReference("Счета-фактуры");
            var invoiceReference = invoiceReferenceInfo.CreateReference();

            invoiceReference.LoadSettings.AddAllParameters();
            var invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.LinkInvoicesToContractor);
            invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.ListObjectDeliveredItems);
            invoiceRelation.AddAllParameters();


            var invoiceRelation2 = invoiceRelation.AddRelation(Guids.DeliveredItems.LinkNomStore);
            invoiceRelation2.AddAllParameters();

            try
            {
                errorId = 1;
                for (int i = 0; i < reportData.Count; i++)
                {

                   // var begData1 = DateTime.Now;


                 //   var timesp1 = DateTime.Now - begData1;
                 //   var timesp2 = DateTime.Now - begData1;
                 //   var timesp3 = DateTime.Now - begData1;
                    //Заполнение параметров объекта Номенклатуры
                    WriteRow(xls, reportData[i], logDelegate, row, i);
                 //   timesp1 = DateTime.Now - begData1;
                    errorId = 2;
                    rowCount = row;
                    progressBar.PerformStep();
                    for (int j = 0; j < reportData[i].StoreObjects.Count; j++ )
                    {
                        errorId = 3;
                      //  var begData2 = DateTime.Now;
                        //Заполнение параметров объекта склада
                        WriteRow(xls, reportData[i].StoreObjects[j], logDelegate, row);
                     //   timesp2 = DateTime.Now - begData2;
                      //  var begData3 = DateTime.Now;
                        var lastInvoices = SearchInvoice(reportData[i].StoreObjects[j], invoiceReference, invoiceReferenceInfo);//invoices.Where(t => t.pkiObjects.Count(p => p.RefObject.SystemFields.Id == storeObj.RefObj.SystemFields.Id) > 0).OrderByDescending(t => t.Date).FirstOrDefault();
                     //   timesp3 = DateTime.Now - begData3;
                        if (lastInvoices != null)
                        {
                            errorId = 4;
                            //Заполнение параметров объекта счета факутры
                            WriteRow(xls, lastInvoices, reportData[i].StoreObjects[j], logDelegate, row, rowCount);
                          

                        }
                      
                        if (reportData[i].StoreObjects.Count > 1 && reportData[i].StoreObjects.Count - 1 != j)
                        {
                            for (int k = 1; k < 13; k++)
                            {
                                errorId = 5;
                                xls[k, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                                xls[k, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                for (int r = 0; r < 2; r++)
                                {
                                    xls[k, row+r].Interior.Color = Color.IndianRed;
                                }
                            }
                            row++;
                        }
                    }
                    for (int k = 1; k < 13; k++)
                    {
                        xls[k, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                        xls[k, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;                        
                    }
                    row++;

                  //  MessageBox.Show("для " + reportData[i] + "\r\n" + timesp1 + "\r\n" + timesp2 + "\r\n" + timesp3);
                }
                errorId = 6;
                PrintSetting(xls);
                xls[1, 1, 12, row].WrapText = true;
                var formula = "=СУММ(" + xls[8, beginRow].Address + ":" + xls[8, row-1].Address + ")";
                xls[8, row].SetFormula(formula);
                xls[8, row].NumberFormat = @"#,##0.00";
                return row;
            }
            catch (Exception ex)
            {
                MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return 0;
        }

        private void MakeSelectionByTypeReport(Xls xls, Dictionary<DeviceWithKoef, List<NomObject>> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            var invoiceReferenceInfo = ReferenceCatalog.FindReference("Счета-фактуры");
            var invoiceReference = invoiceReferenceInfo.CreateReference();

            invoiceReference.LoadSettings.AddAllParameters();
            var invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.LinkInvoicesToContractor);
            invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.ListObjectDeliveredItems);
            invoiceRelation.AddAllParameters();

            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(reportData.Sum(t=>t.Value.Count)));
            if (progressBar.Maximum < Convert.ToDouble(reportData.Sum(t => t.Value.Count)))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            int errorId = 0;
            int col = 1;
            int row = 1;
            if (reportParameters.SelectedDevice)
            {
                PasteHeadName(xls, reportParameters, col, row, 14, reportParameters.Device.ToString());
            }
            else
            {
                if (reportParameters.SelectedDevicesWithKoef)
                {
                    PasteHeadName(xls, reportParameters, col, row, 14, "Расчет стоимости ПКИ для списка устройств с коэффициентами");
                    row++;
                    row = PasteHeaderListDevices(xls, col, row, reportParameters.SelectedDevicesWithKoef);

                    for (int i = 0; i < reportData.Count; i++)
                    {
                        col = 1;
                        xls[col, row].SetValue(i + 1);
                        col++;
                        xls[col, row].SetValue(reportData.Keys.ToList()[i].NomObj[Guids.NomenclatureParameters.NameGuid].Value.ToString());
                        col++;
                        xls[col, row].SetValue(reportData.Keys.ToList()[i].NomObj[Guids.NomenclatureParameters.DenotationGuid].Value.ToString());
                        col++;
                        xls[col, row].SetValue(reportData.Keys.ToList()[i].Count);
                        col++;
                        xls[col, row].SetValue(reportData.Keys.ToList()[i].Koef);
                        xls[1, row, 6, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                             XlBordersIndex.xlEdgeBottom,
                                                                                             XlBordersIndex.xlEdgeLeft,
                                                                                             XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
                        row++;
                    }
                }
            }
          
          //  int rowCount = 1;
            try
            {
                errorId = 1;
                for (int l = 0; l < reportData.Count; l++)
                {
                    errorId = 129;
                    //Заполнение параметров объекта Номенклатуры
                    row++;
                    col = 1;
                   
                    WriteRowDevice(xls, reportData.Keys.ToList()[l], logDelegate, row, l);
                    row++;
                    row = PasteHeader(xls, col, row, reportParameters);
                    var prewRow = row;
                    for (int i = 0; i < reportData.Values.ToList()[l].Count; i++)
                    {
                       // MessageBox.Show(string.Join("\r\n" , reportData.Values.ToList()[l][i].);
                        errorId = 29;
                        WriteRow(xls, reportData.Values.ToList()[l][i], logDelegate, row, i);
                        errorId = 2;
                        int rowCount = row;
                        progressBar.PerformStep();
                        for (int j = 0; j < reportData.Values.ToList()[l][i].StoreObjects.Count; j++)
                        {
                            errorId = 3;
                            //Заполнение параметров объекта склада
                            WriteRow(xls, reportData.Values.ToList()[l][i].StoreObjects[j], logDelegate, row);

                            var lastInvoices = SearchInvoice(reportData.Values.ToList()[l][i].StoreObjects[j], invoiceReference, invoiceReferenceInfo);
                            if (lastInvoices != null)
                            {
                                errorId = 4;
                                //Заполнение параметров объекта счета факутры
                                WriteRow(xls, lastInvoices, reportData.Values.ToList()[l][i].StoreObjects[j], logDelegate, row, rowCount);
                            }
                            errorId = 478;

                            if (reportData.Values.ToList()[l][i].StoreObjects.Count > 1 && reportData.Values.ToList()[l][i].StoreObjects.Count - 1 != j)
                            {
                                for (int k = 1; k < 13; k++)
                                {
                                    errorId = 5;
                                    xls[k, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                                    xls[k, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                                    for (int r = 0; r < 2; r++)
                                    {
                                        xls[k, row + r].Interior.Color = Color.IndianRed;
                                    }
                                }
                                row++;
                            }
   
                            errorId = 479;
                        }
                        for (int k = 1; k < 13; k++)
                        {
                            xls[k, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                            xls[k, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        }
                        row++;
                    }
                    errorId = 482;
                   // var formula = "=СУММ(" + xls[8, prewRow].Address + ":" + xls[8, row].Address + ")";
                //    xls[8, row].SetFormula("=СУММ(" + xls[8, prewRow].Address + ":" + xls[8, row].Address + ")");
                    xls[6, l + 3].SetFormula("=ПРОИЗВЕД(СУММ(" + xls[8, prewRow].Address + ":" + xls[8, row].Address + ");" + xls[5, l + 3].Address + ";" + xls[4, l + 3].Address + ")");
                    //xls[8, row].SetFormula(formula);
                 //   xls[7, row].SetValue("Итого");
                  //  xls[6, row, 2, 1].MergeCells = true;
                   // xls[6, row, 3, 1].Font.Bold = true;
                    //xls[6, row, 3, 1].Font.Size = 14;
                    row++;
                }

                errorId = 6;
                PrintSetting(xls);
                xls[1, 1, 12, row].WrapText = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MakeSelectionByTypeReport(Xls xls, List<StoreObject> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(reportData.Count));
            if (progressBar.Maximum < (reportData.Count))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            var invoiceReferenceInfo = ReferenceCatalog.FindReference("Счета-фактуры");
            var invoiceReference = invoiceReferenceInfo.CreateReference();

            invoiceReference.LoadSettings.AddAllParameters();
            var invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.LinkInvoicesToContractor);
            invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.ListObjectDeliveredItems);
            invoiceRelation.AddAllParameters();


            int errorId = 0;
            int col = 1;
            int row = 1;
            if (reportParameters.ObjectsReferenceFailure)
                PasteHeadName(xls, reportParameters, col, row, 14, "Отчет по объектам справочника Реестр отказов КРЭ и ПКИ за период c " + 
                    reportParameters.BeginPeriod.ToShortDateString() + " по " + reportParameters.EndPeriod.ToShortDateString());
            else
                PasteHeadName(xls, reportParameters, col, row, 14, "Отчет для загруженного списка изделий");
            row++;
            col = 1;
            row = PasteHeader(xls, col, row, reportParameters);
            int rowCount = 1;

            try
            {
                for (int i = 0; i < reportData.Count; i++)
                {
                    progressBar.PerformStep();
                    //Заполнение параметров объекта склада
                    WriteRow(xls, reportData[i], logDelegate, row, i, reportParameters);
                    rowCount = row;



                    var lastInvoices = SearchInvoice(reportData[i], invoiceReference, invoiceReferenceInfo);//invoices.Where(t => t.pkiObjects.Count(p => p.RefObject.SystemFields.Id == storeObj.RefObj.SystemFields.Id) > 0).OrderByDescending(t => t.Date).FirstOrDefault();
                    if (lastInvoices != null)
                    {
                        //Заполнение параметров объекта счета факутры
                        WriteRow(xls, lastInvoices, reportData[i], logDelegate, row, rowCount);
                    }
                    for (int k = 1; k < 13; k++)
                    {
                        xls[k, row].VerticalAlignment = XlVAlign.xlVAlignTop;
                        xls[k, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    }
                    row++;
                }

                PrintSetting(xls);
                xls[1, 1, 12, row].WrapText = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("4. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorId.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row, int sizeFont, string head)
        {
            xls[1, row].SetValue(head);
            xls[1, row].Font.Bold = true;
            xls[col, row].Font.Size = sizeFont;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 12, 1].MergeCells = true;

            row++;

            return row;
        }

        private static void PrintSetting(Xls xls)
        {
            xls.SetColumnWidth(1, 6);
            xls.SetColumnWidth(2, 30);
            xls.SetColumnWidth(3, 30);
            xls.SetColumnWidth(4, 11);
            xls.SetColumnWidth(5, 5);
            xls.SetColumnWidth(6, 7);
            xls.SetColumnWidth(7, 11);
            xls.SetColumnWidth(8, 11);
            xls.SetColumnWidth(9, 13);
            xls.SetColumnWidth(10, 15);
            xls.SetColumnWidth(11, 23);
            xls.SetColumnWidth(12, 11);
            xls.Worksheet.Application.PrintCommunication = true;
        }

        private static void WriteRow(Xls xls, NomObject data, LogDelegate logDelegate, int row, int number)
        {
            logDelegate(String.Format("Запись объекта номенклатуры: " + data.Name));

            xls[1, row].SetValue((number + 1));
            xls[2, row].SetValue(data.Name);
            xls[6, row].SetValue(data.Count);
            xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteRowDevice(Xls xls, DeviceWithKoef data, LogDelegate logDelegate, int row, int number)
        {
            logDelegate(String.Format("Запись устройства: " + data.NomObj));

          //  xls[1, row].SetValue((number + 1));
            xls[2, row].SetValue(data.NomObj.ToString());
            xls[2, row].Font.Size = 12;
            xls[2, row].Font.Bold = true;
            xls[2, row, 2, 1].MergeCells = true;
            
            //xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
            //                                                                                         XlBordersIndex.xlEdgeBottom,
            //                                                                                         XlBordersIndex.xlEdgeLeft,
            //                                                                                         XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteRow(Xls xls, StoreObject data, LogDelegate logDelegate, int row)
        {
            logDelegate(String.Format("Запись объекта склада: " + data.Name));

            xls[3, row].SetValue(data.Name);
            xls[4, row].SetValue(data.Article);
            xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteRow(Xls xls, StoreObject data, LogDelegate logDelegate, int row, int number, ReportParameters reportParameters)
        {
            logDelegate(String.Format("Запись объекта склада: " + data.Name));
            xls[1, row].SetValue(number+1);
            if (reportParameters.ObjectsReferenceFailure)
            {
                xls[2, row].SetValue(data.NumberSZ + " от " + data.DateSZ.ToShortDateString());
                xls[3, row].SetValue(data.Name);
                xls[4, row].SetValue(data.Article);
                xls[6, row].SetValue(data.Count);


                xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                                         XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
            else
            {

                if (data.NomObjects.Count > 0)
                {

                    xls[2, row].SetValue(string.Join("\r\n", data.NomObjects.Select(t => t.ToString())));
                }
                xls[3, row].SetValue(data.Name);
                xls[4, row].SetValue(data.Article);
                xls[6, row].SetValue(data.Count);


                xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                                         XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
        }

        private static void WriteRow(Xls xls, Invoice data, StoreObject storeObject, LogDelegate logDelegate, int row, int rowCount)
        {
            var numberDate = data.Number + " " + data.Date.ToShortDateString();
            var pki = data.pkiObjects.Where(t => t.StoreObject != null && t.StoreObject.SystemFields.Id == storeObject.RefObj.SystemFields.Id).FirstOrDefault();
            logDelegate(String.Format("Запись счета факутуры: " + numberDate));

            if (pki != null)
            {
                int errorID = 0;
                try
                {
                    xls[7, row].SetValue(pki.Price);
                    errorID = 1;
                    xls[7, row].NumberFormat = @"#,##0.00";
                    errorID = 2;
                    var formula = "=ПРОИЗВЕД(" + xls[7, row].Address + ";" + xls[6, rowCount].Address + ")";
                    errorID = 3;
                    xls[8, row].SetFormula(formula);
                    errorID = 4;
                    xls[8, row].NumberFormat = @"#,##0.00";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("3. Возникла непредвиденная ошибка. Обратитесь в отдел 911. Код ошибки  " + errorID.ToString() + "\r\n" + ex, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            xls[9, row].SetValue(numberDate);
            if (data.ContractorObject != null)
            {
                xls[11, row].SetValue(data.ContractorObject.ToString());
                xls[12, row].SetValue(data.ContractorObject[Guids.Contractors.INN].Value.ToString());
            }
            xls[10, row].SetValue(data.ContractorInvoice);

            xls[1, row, 12, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

    }
}
