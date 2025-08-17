using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using Microsoft.Office.Interop.Excel;
using CalcualtionOtherProducts;
using System.Drawing;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Search;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model.References.Nomenclature;

namespace CalculationOtherProducts
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
            selectionForm.Init(context);
            if (selectionForm.ShowDialog() == DialogResult.OK)
            {
                CalculationOtherProducts calcOI = new CalculationOtherProducts();
                var reportParameters = selectionForm.MakeReport();

                var nomenclatureReference = new NomenclatureReference();
                reportParameters.RootObject = (NomenclatureObject)nomenclatureReference.Find(context.ObjectsInfo[0].ObjectId);

                calcOI.MakeReport(context, reportParameters);
            }
        }
    }

    public delegate void LogDelegate(string line);

    public class CalculationOtherProducts
    {

        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            CalculationOtherProducts report = new CalculationOtherProducts();
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

            CalculationOtherProducts report = new CalculationOtherProducts();

            m_form.setStage(IndicationForm.Stages.DataAcquisition);


            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

            m_writeToLog(String.Format("Загружено {0} объектов", reportData.Count.ToString()));

            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;
                //  xls.Open(context.ReportFilePath);

                m_form.setStage(IndicationForm.Stages.DataProcessing);
                m_writeToLog("=== Обработка данных ===");

                m_form.progressBar.PerformStep();

                var structuredData = ProcessData(reportData, m_writeToLog, m_form.progressBar, reportParameters.RootObject);

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportParameters, m_writeToLog, m_form.progressBar, structuredData);

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

                xls.Application.DisplayAlerts = true;
                xls.Visible = true;
            }

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№", col, row, xls);
            col = InsertHeader("Наименование", col, row, xls);
            col = InsertHeader("Ед. изм.", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Цена, руб.", col, row, xls);
            col = InsertHeader("Сумма, руб.", col, row, xls);
            col = InsertHeader("Обоснование", col, row, xls);

            // xls.SetRowHeight(row, 45);

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
            xls[col, row].SetValue(reportParameters.RootObject.ToString());
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row, 7, 1].MergeCells = true;

            row = row + 2;

            return row;
        }

        static public List<TableData> SumDublicates(List<TableData> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ShortName.Trim() == data[slaveID].ShortName.Trim())
                    {
                        data[mainID].Count += data[slaveID].Count;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
        }

        private void MakeSelectionByTypeReport(Xls xls,
            ReportParameters reportParameters,
            LogDelegate logDelegate,
            ProgressBar progressBar,
            Dictionary<string, List<StructureTableData>> structuredData)
        {
            double progressStep = 0;
            try
            {
                //Задание начальных условий для progressBar
                progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(structuredData.Sum(t => t.Value.Sum(k => k.Data.Count))) * 2);

                if (progressBar.Maximum < structuredData.Sum(t => t.Value.Sum(k => k.Data.Count)) * 2)
                {
                    progressBar.Step = 1;
                }
                else
                {
                    progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
                }
            }
            catch
            {
                MessageBox.Show("Ошибка прогрессбара " + Convert.ToDouble(structuredData.Sum(t => t.Value.Sum(k => k.Data.Count))));
            }

            double step = 0;



            int i = 0;


            int row = 1;
            int col = 1;
            string department = string.Empty;
            row = PasteHeadName(xls, reportParameters, col, row);
            HeaderTable(xls, reportParameters, col, row);

            xls[1, row + 1, 1, 11].Select();
            xls.Application.ActiveWindow.FreezePanes = true;
            row++;


            foreach (var firstSection in structuredData)
            {
                bool firstOccurrenceSections = true;

                var countObject = 0;
                bool dataExists = false;

                foreach (var secondSections in firstSection.Value)
                {
                    countObject += secondSections.Data.Count;

                    if (countObject > 0)
                    {
                        dataExists = true;
                        break;
                    }
                }

                if (firstOccurrenceSections && dataExists)
                {
                    xls[2, row].SetValue(firstSection.Key);
                    xls[2, row].Font.Underline = true;
                    xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[2, row, 6, 1].MergeCells = true;
                    InsertFrame(xls, row);
                    row++;
                }

                foreach (var secondSections in firstSection.Value)
                {
                    bool firstOccurrenceItem = true;

                    if (firstOccurrenceItem && secondSections.Data.Count > 0)
                    {
                        if (secondSections.NameSections != null && secondSections.NameSections != string.Empty)
                        {
                            xls[2, row].SetValue(secondSections.NameSections.ToString());
                            xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            xls[2, row, 6, 1].MergeCells = true;
                            InsertFrame(xls, row);
                            row++;
                        }
                    }

                    foreach (var obj in secondSections.Data.OrderBy(t => t.ShortName).ToList())
                    {
                        i++;
                        WriteRow(xls, logDelegate, row, i, obj.ShortName, obj.MeasureUnit, obj.Count);
                        row++;

                        step += progressStep;
                        if (step >= 1)
                        {
                            progressBar.PerformStep();
                            step = step - progressBar.Step;
                        }
                    }
                    firstOccurrenceItem = false;
                }
                firstOccurrenceSections = false;

            }
            xls.AutoWidth();

            // Clipboard.SetText(string.Join("\r\n", reportData.Where(t => t.FindedObj == false).Select(t => t.Name + " " + t.Denotation).OrderBy(t => t).ToList()));
            //  MessageBox.Show("Не нашел " + reportData.Where(t => t.FindedObj == false).ToList().Count() + " объектов");
        }

        private Dictionary<string, List<StructureTableData>> ProcessData(List<TableData> reportData, LogDelegate logDelegate, ProgressBar progressBar, ReferenceObject rootObject)
        {
            double progressStep = 0;
            try
            {
                //Задание начальных условий для progressBar
                progressStep = Convert.ToDouble(progressBar.Maximum) / (Convert.ToDouble(reportData.Count * 2));

                if (progressBar.Maximum < Convert.ToDouble(reportData.Count * 2))
                {
                    progressBar.Step = 1;
                }
                else
                {
                    progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
                }
            }
            catch
            {
                MessageBox.Show("Ошибка прогрессбара " + Convert.ToDouble((reportData.Count * 2)));
            }

            double step = 0;



            Dictionary<string, List<StructureTableData>> structuredData = new Dictionary<string, List<StructureTableData>>();

            logDelegate(String.Format("Загружено {0} объектов", reportData.Count.ToString()));

            ReferenceInfo referenceCategoriesInfo = ReferenceCatalog.FindReference(Guids.References.CategoriesOtherItems);
            Reference categoriesOtherItemsReference = referenceCategoriesInfo.CreateReference();
            //   categoriesOtherItemsReference.LoadSettings.AddAllParameters();
            //   var relation = categoriesOtherItemsReference.LoadSettings.AddRelation(Guids.CategoryOtherItems.LinkOtherItems);
            //   relation.AddAllParameters();
            //  relation = categoriesOtherItemsReference.LoadSettings.AddRelation(Guids.CategoryOtherItems.LinkOtherItems);
            //  relation.AddParameters(Guids.CategoryOtherItems.Name);

            logDelegate("Загрузка объектов из справочника Категории прочих изделий");

            categoriesOtherItemsReference.Objects.Load();

            var sections = new List<FirstSection>();
            logDelegate("Загружено " + categoriesOtherItemsReference.Objects.ToList().Count + " объектов");
            //   Clipboard.SetText(string.Join("\n\r", categoriesOtherItemsReference.Objects.Select(t=>t.ToString())));

            foreach (var category in categoriesOtherItemsReference.Objects.OrderBy(t => Convert.ToInt32(t[Guids.CategoryOtherItems.SerialIndex].Value)).ToList())
            {
                //Объекты справочника Категории прочих изделий
                var objectsCategory = category.Children.Where(t => t.Class.IsInherit(Guids.CategoryOtherItems.ObjectTypes)).ToList();
                //Папки справочника Категории прочих изделий
                var foldersCategory = category.Children.Where(t => t.Class.IsInherit(Guids.CategoryOtherItems.FolderTypes)).ToList();

                FirstSection firstSection = new FirstSection();
                //Наименование раздела = наименование папки
                firstSection.Name = category[Guids.CategoryOtherItems.Name].Value.ToString();
                
                if (objectsCategory.Count > 0)
                {
                    //First так как в папках содержится один объект
                    firstSection.OINamesTextNorm = objectsCategory.First()[Guids.CategoryOtherItems.OtherItemsNames].Value.ToString().Split(',').Select(t=> textNorm.GetNormalForm(t)).ToList();

                    foreach (var objCat in objectsCategory.OrderBy(t => Convert.ToInt32(t[Guids.CategoryOtherItems.SerialIndex].Value)))
                    {
                        firstSection.OtherItems.AddRange(GetOtherItems(objCat, objCat[Guids.CategoryOtherItems.Name].Value.ToString(),
                        objCat[Guids.CategoryOtherItems.RegularExpression].Value.ToString()).OrderBy(t => t.NameTextNorm).ToList());

                        
                    }
                  //  var oifs = string.Join("\r\n", firstSection.OtherItems.Select(t => t.Name + " " + t.Denotation));
                 //   Clipboard.SetText(oifs);
                   // MessageBox.Show("прочие для " + firstSection.Name);
                }
                else
                {
                    foreach (var folder in foldersCategory)
                    {
                        //Объекты справочника Категории прочих изделий
                        var objectsSecondCategory = folder.Children.Where(t => t.Class.IsInherit(Guids.CategoryOtherItems.ObjectTypes)).ToList();
                        SecondSection secondSection = new SecondSection();
                        secondSection.Name = folder[Guids.CategoryOtherItems.Name].Value.ToString();

                        if (objectsSecondCategory.Count > 0)
                        {
                            //First так как в папках содержится один объект
                            secondSection.OINamesTextNorm = objectsSecondCategory.First()[Guids.CategoryOtherItems.OtherItemsNames].Value.ToString().Split(',').Select(t => textNorm.GetNormalForm(t)).ToList();
                            foreach (var objCat in objectsSecondCategory.OrderBy(t => Convert.ToInt32(t[Guids.CategoryOtherItems.SerialIndex].Value)))
                            {
                                secondSection.OtherItems.AddRange(GetOtherItems(objCat, objCat[Guids.CategoryOtherItems.Name].Value.ToString(),
                                objCat[Guids.CategoryOtherItems.RegularExpression].Value.ToString()).OrderBy(t => t.NameTextNorm).ToList());
                            }
                            firstSection.SecondSection.Add(secondSection);
                        }
                    }
                }
                sections.Add(firstSection);


               // List<StructureOtherItems> listStructureOI = new List<StructureOtherItems>();

               // if (objectsCategory.Count > 0)
                {
                 //   StructureOtherItems emptyStructOI = new StructureOtherItems();
                //    emptyStructOI.NameSections = "";

                  //  foreach (var objCat in objectsCategory.OrderBy(t => Convert.ToInt32(t[Guids.CategoryOtherItems.SerialIndex].Value)))
                    {
                       // var emptySection = oi.GetObjects(Guids.CategoryOtherItems.LinkOtherItems);
                      //  emptyStructOI.OtherItems.AddRange(GetOtherItems(oi, oi[Guids.CategoryOtherItems.Name].Value.ToString(), oi[Guids.CategoryOtherItems.RegularExpression].Value.ToString()).OrderBy(t => t.NameTextNorm).ToList());
                    }

                  //  listStructureOI.Add(emptyStructOI);
                }

                //if (otherItems.Count == 0 && folders.Count == 0)
                //{
                //    StructureOtherItems emptyStructOI = new StructureOtherItems();
                //    emptyStructOI.NameSections = "";
                //    emptyStructOI.OtherItems.Add(new OtherItem("Пусто", "Пусто", "Пусто", "Пусто", "Пусто"));
                //    listStructureOI.Add(emptyStructOI);
                //}


              //  foreach (var folder in foldersCategory)
                {
                    //StructureOtherItems structOI = new StructureOtherItems();
                    //structOI.NameSections = folder[Guids.CategoryOtherItems.Name].Value.ToString();
                    //var oi = GetOtherItems(folder, string.Empty, string.Empty);

                    //structOI.OtherItems.AddRange(oi);
                    
                    //listStructureOI.Add(structOI);
                }
               // sections.Add(category[Guids.CategoryOtherItems.Name].Value.ToString(), listStructureOI);
            }


            //Регулярное выражение для формирования сокращенного наименования
            var pattern = @"(?<=^[А-ЯA-Zа-яa-z]{3,}\s(([А-ЯA-Zа-яa-z]{3,}\s)?))([А-ЯA-Z]{1,3}\d{0,2}([А-ЯA-Z]{1,2})?-(\d{1,3}))|(\d{1,4}[А-ЯA-Z]{1,2}\d{1,3}([А-ЯA-Z]{0,2}?))|((\d?)[А-ЯA-Z]{1,5}\d{1,5})|([А-ЯA-Z]{2,3}-[А-ЯA-Z]{2}(\d{1,3})?)|([А-ЯA-Z]{1}(\d{1}?)(\s?)(,?)(\d{1,3}))";
            Regex regex = new Regex(pattern);

            //     MessageBox.Show(string.Join("\r\n*", sections.Select(t=>string.Join("\r\n|", t.Value))));
            //   Clipboard.SetText(string.Join("\r\n*", sections.Select(t => string.Join("\r\n|", t.Value))));
            var listObjects = string.Join("\r\n",reportData.Select(t => t.Name + " " + t.Denotation + " " + t.NameTextNorm));
          //  Clipboard.SetText(listObjects);
          //  MessageBox.Show(listObjects);
           // MessageBox.Show(string.Join("\r\n", sections.Select(t => t.Name + "\r\n" + string.Join("\r\n", t.SecondSection.Select(r => "       -" + r.Name)))));


            string nameSection = string.Empty;
            //Первый уровень заголовков
  
           


            foreach (var firstSection in sections)
            {
                List<TableData> sortData = new List<TableData>();
                var data = new List<StructureTableData>();

                foreach (var secondSections in firstSection.SecondSection)
                {
                    var findedData = AddTableData(progressBar, progressStep, ref step, pattern, reportData, null, secondSections);
                    data.Add(findedData);
                    try
                    {
                        //if (findedData.Data.Count > 0)
                        //{
                           
                        //    //   break;
                        //}
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Код ошибки № 11. Заголовок " + secondSections.Name + " Обратитесь в отдел 911 \r\n", "Внимание ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }



                if (firstSection.SecondSection.Count == 0)
                {
                    var findedData = AddTableData(progressBar, progressStep, ref step, pattern, reportData, firstSection, null);
                    data = new List<StructureTableData>();
                    data.Add(findedData);
                    try
                    {
                        if (findedData.Data.Count > 0)
                        {
                            structuredData.Add(firstSection.Name, data);
                            // break;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Код ошибки № 12. Заголовок " + firstSection.Name + " Обратитесь в отдел 911 \r\n", "Внимание ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (data.Count > 0)
                        structuredData.Add(firstSection.Name, data);
                }
            }


          //  var result = string.Join("\r\n", structuredData.Select(t => "1***" + t.Key + "\r\n" + string.Join("\r\n2++", t.Value.Select(r => "3---" + r.NameSections + "\r\n" + string.Join("\r\n4##", r.Data.Select(v => v.Name + " " + v.Denotation))))));
          //  Clipboard.SetText(result);
           // MessageBox.Show("res");

            var notFoundObjects = reportData.Where(t => t.FindedObj == false).ToList();

            if (notFoundObjects.Count() > 0)
            {
                var notFoundObjectsForm = new NotFoundObjectsForm();
                notFoundObjectsForm.Tag = rootObject;
                m_form.Visible = false;
                int i = 0;

                foreach (var obj in notFoundObjects)
                {
                    var match = Regex.Match(obj.Name, pattern);
                    obj.ShortName = match.Value.ToString();

                    if (obj.ShortName.Trim() == string.Empty)
                        obj.ShortName = obj.Name + " " + obj.Denotation;
                }

                foreach (var obj in SumDublicates(notFoundObjects.OrderBy(t => t.ShortName).ToList()))
                {
                    i++;
                    var itmx = new ListViewItem[1];
                    itmx[0] = new ListViewItem(i.ToString());
                    itmx[0].SubItems.Add(obj.Name + " " + obj.Denotation);
                    itmx[0].SubItems.Add(obj.ShortName);
                    itmx[0].Tag = obj;

                    notFoundObjectsForm.listViewNotFoundObjects.Items.AddRange(itmx);
                }

                foreach (var firstSection in sections)
                {
                    notFoundObjectsForm.comboBoxSections.Items.Add(firstSection.Name.ToString());

                    foreach (var secondSections in firstSection.SecondSection.Select(t => t.Name.Trim()).Where(t => t.ToString() != string.Empty))
                    {
                        notFoundObjectsForm.comboBoxSections.Items.Add(secondSections);
                    }
                }
                notFoundObjectsForm.comboBoxSections.Text = notFoundObjectsForm.comboBoxSections.Items[0].ToString();


                notFoundObjectsForm.StructuredData = structuredData;

                //  Clipboard.SetText(string.Join("\r\n", reportData.Where(t => t.FindedObj == false).Select(t => t.Name + " " + t.Denotation).OrderBy(t => t).ToList()));
                if (notFoundObjectsForm.ShowDialog() == DialogResult.OK)
                {
                    m_form.Visible = true;

                    //  var rrr =  notFoundObjectsForm.StructuredData.OrderBy(t=>t.Value.OrderBy(c=>c.Data);

                    return notFoundObjectsForm.StructuredData;
                }
            }
            //   Clipboard.SetText(string.Join("\r\n", reportData.Where(t => t.FindedObj == false).Select(t => t.Name + " " + t.Denotation).OrderBy(t => t).ToList()));
            
            return structuredData;
        }

        private StructureTableData AddTableData(ProgressBar progressBar, double progressStep, ref double step, string pattern,  List<TableData> reportData, FirstSection firstSection, SecondSection secondSections)
        {
            StructureTableData structData = new StructureTableData();
            List<TableData> sortData = new List<TableData>();
            List<OtherItem> coincidenceList;
            string nameSection = string.Empty;
            foreach (var obj in reportData)
            {
                if (firstSection != null)
                {
                    coincidenceList = firstSection.OtherItems.Where(t => t.NameTextNorm.Trim() == obj.NameTextNorm.Trim() ||
                            (firstSection.OINamesTextNorm.Count(n=>obj.NameTextNorm.IndexOf(n) == 0)>0)).ToList();
                    nameSection = firstSection.Name;
                }
                else
                {
                    coincidenceList = secondSections.OtherItems.Where(t => t.NameTextNorm == obj.NameTextNorm ||
                     (secondSections.OINamesTextNorm.Count(n => obj.NameTextNorm.IndexOf(n) == 0) > 0)).ToList();
                    nameSection = secondSections.Name;
                }

                if (coincidenceList.Count > 0)
                {
                    //   MessageBox.Show(nameSection + "\r\n" + obj);
                    step += progressStep;
                    if (step >= 1)
                    {
                        progressBar.PerformStep();
                        step = step - progressBar.Step;
                    }
                    Match match = null;
                    var regularExpression = coincidenceList.FirstOrDefault().RegularExpression;
                    if (regularExpression != string.Empty)
                    {
                        //Считываем регулярное выражение из группы
                        match = Regex.Match(obj.Name, regularExpression);
                        //     MessageBox.Show(coincidenceList.FirstOrDefault().RegularExpression);
                        obj.ShortName = match.Value.ToString();
                        obj.RegularExpression = regularExpression;
                    }
                    else
                    {
                        match = Regex.Match(obj.Name, pattern);
                        obj.ShortName = match.Value.ToString();
                        obj.RegularExpression = pattern;
                    }



                    if (obj.ShortName.Trim() == string.Empty)
                        obj.ShortName = obj.Name + " " + obj.Denotation;

                    obj.FindedObj = true;
                    sortData.Add(obj);

                }


                  
            }

            bool addSection = false;

            //Суммирование дубликатов
            foreach (var srtObj in SumDublicates(sortData))
            {
                if (firstSection == null)
                {
                    structData.NameSections = nameSection;
                }
                structData.Data.Add(srtObj);
                addSection = true;
            }

            if (!addSection)
            {
                if (firstSection == null)
                {
                    structData.NameSections = nameSection;
                }
                structData.Data = new List<TableData>();
            }

            //if (sortData.Count > 0)
            // data.Add(structData);
            //MessageBox.Show(structData.NameSections + "\r\n" +string.Join("\r\n", structData.Data.Select(t=>t.Name + " " + t.Denotation)));
            return structData;
        }

        private static void InsertFrame(Xls xls, int row)
        {
            xls[1, row, 7, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }



        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, string name,string measureUnit, double count)
        {
            logDelegate(String.Format("Запись объекта: {0}", name));

            xls[1, row].SetValue(i.ToString());
            xls[2, row].SetValue(name);
            //xls[3, row].NumberFormat = "@";
            xls[3, row].SetValue(measureUnit);
            xls[4, row].SetValue(count);

            InsertFrame(xls, row);
        }


  


        TextNormalizer textNorm = new TextNormalizer();

        public List<OtherItem> GetOtherItems(ReferenceObject refObj, string singularName, string regularExpression)
        {
            var otherItems = new List<OtherItem>();

            if (refObj.Class.IsInherit(Guids.OtherItems.ObjectTypes))
            {
                var name = refObj[Guids.OtherItems.Name].Value.ToString();
                var denotation = refObj[Guids.OtherItems.Denotation].Value.ToString();
                var nameTextNorm = textNorm.GetNormalForm(name + denotation);

                OtherItem otherItem = new OtherItem(name, denotation, nameTextNorm, singularName, regularExpression);
                otherItems.Add(otherItem);
            }

            if (refObj.Class.IsInherit(Guids.OtherItems.FolderTypes)) 
            {
                foreach (var el in refObj.Children)
                {
                    otherItems.AddRange(GetOtherItems(el, singularName, regularExpression));
                }
            }

            if (refObj.Class.IsInherit(Guids.CategoryOtherItems.ObjectTypes))
            {
                var linkOtherItems = refObj.GetObjects(Guids.CategoryOtherItems.LinkOtherItems);
                 
                foreach (var oi in linkOtherItems)
                {                   
                    otherItems.AddRange(GetOtherItems(oi, refObj[Guids.CategoryOtherItems.Name].Value.ToString(), 
                        refObj[Guids.CategoryOtherItems.RegularExpression].Value.ToString()));          
                }
            }

            if (refObj.Class.IsInherit(Guids.CategoryOtherItems.FolderTypes))
            {
                foreach (var child in refObj.Children)
                {
                    otherItems.AddRange(GetOtherItems(child, string.Empty, string.Empty));
                }
            }

            return otherItems;
        }


        public List<TableData> GetChildren(ReferenceObject rootObject, ReportParameters reportParameters, double countRootObject, LogDelegate logDelegate)
        {
            var listObjects = new List<TableData>();
            int errorID = 0;
            try
            {

                foreach (NomenclatureHierarchyLink childHierarchyLink in rootObject.Children.GetHierarchyLinks())
                {

                    var childCount = Convert.ToDouble(childHierarchyLink[NomHierarchyLink.Fields.Count].Value.ToString());
                    errorID = 3;
                    var childRemark = childHierarchyLink[NomHierarchyLink.Fields.Comment].Value.ToString();
                    errorID = 4;

                    if (childCount != Math.Truncate(childCount))
                        childCount = CountSelectionRezistor(childRemark);
                    errorID = 5;
                    listObjects.AddRange(GetChildren(childHierarchyLink.ChildObject, reportParameters, childCount, logDelegate));
                    errorID = 6;
                    if ((reportParameters.OtherItems && childHierarchyLink.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.OtherItems)) ||
                        (reportParameters.Assembly && childHierarchyLink.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Assembly)) ||
                        (reportParameters.StandartItems && childHierarchyLink.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.StandartItems)) ||
                        (reportParameters.Materials && childHierarchyLink.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Materials)) ||
                        (reportParameters.Details && childHierarchyLink.ChildObject.Class.IsInherit(Guids.NomenclatureTypes.Details)))
                    {
                        errorID = 7;

                        var name = childHierarchyLink.ChildObject[ParametersInfo.Name].Value.ToString();
                        var denotation = childHierarchyLink.ChildObject[ParametersInfo.Denotation].Value.ToString();
                        errorID = 8;
                        var obj = new TableData(name,
                            denotation,
                            childHierarchyLink[NomHierarchyLink.Fields.MeasureUnit].Value.ToString(),
                            Convert.ToDouble(childHierarchyLink[NomHierarchyLink.Fields.Count].Value.ToString()),
                            textNorm.GetNormalForm(name + denotation),
                            string.Empty,
                            childHierarchyLink[NomenclatureHierarchyLink.FieldKeys.Remarks].Value.ToString());
                        errorID = 9;
                        listObjects.Add(obj);

                        logDelegate(String.Format("Считывание объекта: {0} {1}", obj.Name, obj.Denotation));
                    }
                }
            }
            catch
            {
                MessageBox.Show("Код ошибки " + errorID + ". Обратиетсь в отдел 911", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return listObjects;
        }

        private int CountSelectionRezistor(string remark)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            Regex regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value + " " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                }
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
            }
            catch
            {
                return 0;
                // System.Windows.Forms.MessageBox.Show("7. Не могу посчитать количество резистора\r\n" + nomObj + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId + "\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return countRezist;
        }

        private List<TableData> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            List<TableData> reportData = new List<TableData>();
            logDelegate("=== Загрузка данных ===");
            //------------------------------------------------------------------------------------
            ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(Guids.References.Nomenclature);
            Reference nomenclatureReference = referenceNomenclatureInfo.CreateReference();

            nomenclatureReference.LoadSettings.AddParameters(Guids.NomenclatureParameters.Name, Guids.NomenclatureParameters.Denotation);
            List<TableData> listObjects = new List<TableData>();

            if (reportParameters.UseCurrentObject)
            {
                listObjects.AddRange(GetChildren(reportParameters.RootObject, reportParameters, 1, logDelegate));
            }
            else
            {
                if(reportParameters.Devices != null && reportParameters.Devices.Count > 0)
                {
                    foreach (var rootObject in reportParameters.Devices)
                    {
                        listObjects.AddRange(GetChildren(rootObject.NomObject, reportParameters, rootObject.Count, logDelegate));
                    }                
                }
                else
                {
                    MessageBox.Show("Загруженный список устройств отсутствует. Отчет будет пустым.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            return listObjects;
        }

        //public static Dictionary<string, string> SemiconductorDevices = new Dictionary<string, string>()
        //{
        //     {"Диод","Диоды"},
        //     {"Варикап", "Варикапы"},
        //     {"Стабилитрон","Стабилитроны"},
        //     {"Диодная матрица","Диодные матрицы"},
        //     {"Модуль","Модули"},
        //     {"Диодная сборка","Диодные сборки"},
        //     {"Транзистор","Транзисторы"},
        //     {"Транзисторная матрица","Транзисторные матрицы"},
        //     {"Транзисторная сборка","Транзисторные сборки"},
        //     {"Тиристор","Тиристоры"}
        //};

        //public static Dictionary<string, string> DevicesPiezoelectric = new Dictionary<string,string>()
        //{
        //     {"Резонатор","Резонаторы"},
        //     {"Генератор","Генераторы"}
        //};

        //public static Dictionary<string, string> ResistorsCondensers = new Dictionary<string,string>()
        //{
        //     {"Резистор","Резисторы"},
        //     {"Конденсатор","Конденсаторы"}
        //};

        //public static Dictionary<string, string> TransformersThrottles = new Dictionary<string, string>()
        //{
        //     {"Трансформатор","Трансформаторы"},
        //     {"Дроссель","Дроссели"},
        //     {"Линия задержки","Линии задержки"}
        //};

        //public static Dictionary<string, string> ProductsSwitching = new Dictionary<string, string>()
        //{
        //     {"Реле","Реле"},
        //     {"Контакт магнитоуправляемый герметизированный","Контакт магнитоуправляемый герметизированный"},
        //     {"Контактор","Контакторы"},
        //     {"Выключатель","Выключатели"},
        //     {"Микропереключатель","Микропереключатели"},
        //     {"Переключатель","Переключатели"}
        //};

        //public static Dictionary<string,string> Connectors = new Dictionary<string, string>()
        //{
        //     {"Вилка","Вилки"},
        //     {"Розетка","Розетки"},
        //     {"Контакт","Контакт"},
        //     {"Переход","Переход"},
        //     {"Гнездо","Гнездо"},
        //     {"Штепсель","Штепсель"},
        //     {"Вставка плавкая","Вставки плавкие"},
        //     {"Держатель вставок плавких","Держатели вставок плавких"}
        //};

        //public static Dictionary<string,string> FunctionalDevices = new Dictionary<string, string>()
        //{
        //     {"Источники вторичного электропитания","Источники вторичного электропитания"}
        //};

        //public static Dictionary<string, string> ProductsFerrite = new Dictionary<string,string>()
        //{
        //     {"сердечник","Сердечники"}
        //};

        //public static Dictionary<string, Dictionary<string, string>> Sections = new Dictionary<string, Dictionary<string,string>>()
        //{
        //     {"Электрорадиоизделия (ЭРИ)", null},
        //     {"Изделия СВЧ", null},
        //     {"Микросхемы", null},
        //     {"Полупроводниковые приборы", SemiconductorDevices},
        //     {"Оптоэлектронные приборы", null},
        //     {"Индикаторы знакосинтезирующие", null},
        //     {"Приборы пьезоэлектрические и фильтры электромеханические", DevicesPiezoelectric},
        //     {"Резисторы и конденсаторы", null},
        //     {"Трансформаторы, дроссели, линии задержки", TransformersThrottles},
        //     {"Изделия коммутационные", ProductsSwitching},
        //     {"Соединители", Connectors},
        //     {"Машины электрические малой мощности", null},
        //     {"Функциональные устройства",FunctionalDevices},
        //     {"Источники света электрические", null},
        //     {"Изделия из ферритов и магнитодиэлектриков", ProductsFerrite},
        //     {"Микросборки и многокристальные модули", null},
        //     {"ПКИ (кроме крупных дорогостоящих)", null},
        //     {"Инструмент и оснастка", null},
        //     {"Прочие ТМЦ", null},
        //     {"Дорогостоящие ПКИ", null}
        //};

        //private static TFlex.DOCs.Model.Search.Filter createTypeFilter(ReferenceInfo referenceChargesInfo, Reference nomenclatureReference, Guid guidType)
        //{          
        //    ClassObject classObject = referenceChargesInfo.Classes.Find(guidType);
        //    //Создаем фильтр
        //    TFlex.DOCs.Model.Search.Filter filter = new TFlex.DOCs.Model.Search.Filter(referenceChargesInfo);
        //    //Добавлем условие поиска - «Тип = Папка»
        //    filter.Terms.AddTerm(nomenclatureReference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
        //    //Получаем список объектов, в качестве условия поиска – сформированный фильтр   
        //    return filter;
        //}
    }
}
