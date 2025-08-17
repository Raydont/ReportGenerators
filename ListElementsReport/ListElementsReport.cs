using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using TFlex;
using TFlex.Model;
using TFlex.Drawing;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.Model.Model2D;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.References.Nomenclature;
using Application = TFlex.Application;
using Globus.DOCs.Technology.Reports;
using Globus.DOCs.Technology.Reports.CAD;
using TFlex.DOCs.Model.References.Files;
using System.Threading;
using ListElementReport;
using System.Diagnostics;

namespace ListElementsReport
{
    public class CadGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get { return ".grb"; }
        }

        public delegate void LogDelegate(string line);

        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context,
                                           System.Windows.Forms.IWin32Window owner)
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
                //var textNorm = new TextNormalizer();
                //var krum1 = textNorm.GetNormalForm("Дроссель Д263ВСС КРЮМО.475.013ТУ");
                //var krum2 = textNorm.GetNormalForm("Дроссель Д263ВСС КРЮМ0.475.013ТУ");
                //Clipboard.SetText(krum1 + "\r\n" + krum2);
                //MessageBox.Show(krum1 + "\r\n" + krum2);
                selectionForm.Init(context);
                
                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    ListElementsReport.MakeReport(context, selectionForm.reportParameters);
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


        public class ListElementsReport : IDisposable
        {
            public static readonly Guid NomenclatureReference = new Guid ("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static readonly Guid TypeDetailNomReference = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
            public static readonly Guid TypeAssemblyNomReference = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
            public static readonly Guid TypeOtherItemsNomReference = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
            public static readonly Guid TypeMaterialNomReference = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");
            public static readonly Guid TypeStandartItemsNomReference = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
            public static readonly Guid DenotationNomReference = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            public static readonly Guid OtherItemsReference = new Guid("5cb5931a-3329-466a-ad46-13f07c4b3d7a");
            public static readonly Guid TypeOtherItems = new Guid("bd33ab80-c510-45ed-83c9-db4e62117a57");
            public static readonly Guid NameOtherItemsGuid = new Guid("4d3a600e-0dd2-4741-b045-c37b78fe88ae");
            public static readonly Guid DenotationItemsGuid = new Guid("27fe7568-2650-453e-a34c-a431ea0f5f00");
            public static readonly Guid CommentItemsGuid = new Guid("b796b9f5-1a43-4a1e-b958-f8a505b2dcf7");
            public static readonly Guid NomenclatureNameGuid = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            public static readonly Guid NomenclatureDenotationGuid = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");

            public IndicationForm m_form = new IndicationForm();
            static StringBuilder globalLog = new StringBuilder();

            public void Dispose()
            {
            }
            public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
            {
                globalLog = new StringBuilder();
                ListElementsReport report = new ListElementsReport();
                report.Make(context, reportParameters);
            }

            private IReportGenerationContext _context;
            public List<ElementPki> findedElements = new List<ElementPki>();
            public void Make(IReportGenerationContext context, ReportParameters reportParameters)
            {
                findedElements = new List<ElementPki>();
                _context = context;
                context.CopyTemplateFile();    // Создаем копию шаблона

                m_form.Visible = true;
                LogDelegate m_writeToLog;
                m_writeToLog = new LogDelegate(m_form.writeToLog);

                m_form.setStage(IndicationForm.Stages.Initialization);
                m_writeToLog("=== Инициализация ===");

                bool writeAll = false;

                //if (context.UserName.ToString().Contains("gudov"))
                //{
                //    writeAll = true;
                //    MessageBox.Show("Тестовый режим. Все элементы вне зависимости от ошибок будут записаны в Перечень");
                //}

                try
                {
                    using (ListElementsReport report = new ListElementsReport())
                    {
                        m_form.setStage(IndicationForm.Stages.DataAcquisition);
                        m_writeToLog("=== Получение данных ===");
                        List<ElementPki> objectsForSpecification = new List<ElementPki>();
                        var elements = report.ReadData(reportParameters, m_writeToLog, writeAll, out objectsForSpecification);
                        if (elements == null)
                        {
                            m_form.Dispose();
                            return;
                        }
                       
                        m_form.setStage(IndicationForm.Stages.CheckData);
                        m_writeToLog("=== Проверка данных ===");
                        CheckElementsPKI(reportParameters, m_writeToLog, elements, writeAll);
                        m_form.setStage(IndicationForm.Stages.DataProcessing);
                        m_writeToLog("=== Обработка данных ===");
                        AddAdmission(findedElements);

                        List<ElementPki> detailsNomenclature = new List<ElementPki>();
                       
                        var procData = report.ProcessingData(reportParameters,findedElements, m_writeToLog, writeAll, out detailsNomenclature);

                        // var наименованиеФайла = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\rd.txt";
                        // File.WriteAllText(наименованиеФайла, string.Join("\r\n", reportData.Select(t => t.Name + " " + t.PosDenotation)));
                        // Process.Start(наименованиеФайла);

                        //Формирование намиенований и обозначений объектов для ТОЧЕК и КОНТАКТОВ 
                        findedElements = new List<ElementPki>();
                        CheckElementsPKI(reportParameters, m_writeToLog, objectsForSpecification, writeAll);

                        //Clipboard.SetText(string.Join("\n", procData.Select(t => t.Name + " " + t.Denotation)));
                        //MessageBox.Show("4 " +string.Join("\n", procData.Select(t=>t.Name + " " +t.Denotation)));
                        AddAdmission(findedElements);

                        m_form.setStage(IndicationForm.Stages.ReportGenerating);
                        m_writeToLog("=== Формирование отчета === ");
                        var reportData = SummSelRezist(procData);

                       // var наименованиеФайла = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\rd.txt";
                       // File.WriteAllText(наименованиеФайла, string.Join("\r\n", reportData.Select(t => t.Name + " " + t.PosDenotation)));
                       // Process.Start(наименованиеФайла);

                        MakeSelectionByTypeReport(reportData, reportParameters, m_writeToLog);

                        var linkedObjects = SearchLinkedObjInNomenclature(reportData, detailsNomenclature, m_writeToLog, findedElements);
                        //Clipboard.SetText(string.Join("\r\n", linkedObjects.OrderBy(t=>t.Name + t.Denotation).Select(t=>t.Name + " " + t.Denotation + " " + t.ClassObject + " " + t.PosDenotation + " " + t.Count)));
                        //MessageBox.Show("Прикрепленные объекты");
                        m_writeToLog("=== Формирование файла *.str ===");                                
                        MakeFileStr(reportData.Where(t => !t.Type.ToLower().Contains("konstr")).ToList(), linkedObjects, reportParameters, objectsForSpecification, m_writeToLog);

                        m_form.setStage(IndicationForm.Stages.Done);
                        m_writeToLog("=== Завершение работы ===");
                        if (globalLog.ToString() != string.Empty)
                        {
                            Clipboard.SetText(globalLog.ToString());
                            MessageBox.Show(globalLog.ToString(), "Лог ошибок", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                            System.Diagnostics.Process.Start(context.ReportFilePath);
                       
                    }
                }
                catch (Exception e)
                {
                    string message = String.Format("ОШИБКА: {0}", e.ToString());
                    MessageBox.Show(message, "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                }
                m_form.Dispose();
            }

            private void CheckElementsPKI(ReportParameters reportParameters, LogDelegate m_writeToLog, List<ElementPki> elements, bool writeAll)
            {
                var writedElements = new List<string>();
                StringBuilder strB = new StringBuilder();
                TextNormalizer textNormal = new TextNormalizer();

                m_writeToLog("Преобразование к нормальной форме БД Access " + reportParameters.dBElements.Count() + " объектов");
                foreach (var el in reportParameters.dBElements)
                {
                    try
                    {
                        el.NormalPartNumber = textNormal.GetNormalForm(el.PartNumber);
                      
                    }
                    catch
                    {

                    }
                }

                StringBuilder notFindedFilesObj = new StringBuilder();
;
                //Сравнение NormalPartNumber из БД и Type из считанного файла                
                foreach (var el in elements)
                {
                    try
                    {
                        TextNormalizer textNorm = new TextNormalizer();
                     //   var type = textNorm.GetNormalForm(el.Type.ToLower().Trim());

                      //  var findedEl = reportParameters.dBElements.Where(t => textNorm.GetNormalForm(t.PartNumber) == type).ToList();

                        var findedEl = reportParameters.dBElements.Where(t => t.PartNumber.ToLower().Trim() == el.Type.ToLower().Trim()).ToList();

                        if (findedEl.Count > 0)
                        {
                            el.dbElements.AddRange(findedEl);
                            findedElements.Add(el);
                            // MessageBox.Show("Объекты из файла сопоставленные с БД Access\n" + findedElements.Count().ToString());
                        }
                        else
                        {
                            if (writeAll)
                            {
                                DataBaseElements pki = new DataBaseElements();
                                List<DataBaseElements> pkis = new List<DataBaseElements>();
                                pki.Comment = el.Type + " " + el.Value + el.Admission + " " + el.TU;
                                pki.NotWrite = true;
                                pkis.Add(pki);
                                el.dbElements.AddRange(pkis);
                                findedElements.Add(el);
                            }
                            else
                            {
                                notFindedFilesObj.Append(el.PosDenotation + " " + el.Type + " " + el.Value + el.Admission + " " + el.TU);
                                notFindedFilesObj.AppendLine();
                            }
                        }
                    }
                    catch
                    {
                        // MessageBox.Show(el.Type.ToLower().Trim() + " " + el.Name);
                        globalLog.AppendLine("Ошибка при сравнении с БД Access " + el.Type.ToLower().Trim() + " " + el.Name);
                    }

                }

                //Вывод пользователю сообщения со списком несопоставленных объектов 
                if (notFindedFilesObj.Length > 0)
                {
                    Clipboard.Clear();
                    Clipboard.SetText(notFindedFilesObj.ToString());

                    //MessageBox.Show("Объекты из файла несопоставленные с БД Access\n" + notFindedFilesObj.ToString(), "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    globalLog.AppendLine("Объекты из файла несопоставленные с БД Access\n" + notFindedFilesObj.ToString());
                    Clipboard.SetText(globalLog.ToString());
                }
                //Формирование имён считанных объектов
                SetNames(reportParameters);           
            }

            //Суммирование подборочных резисторов в перенне
            private List<ElementPki> SummSelRezist(List<ElementPki> reportData)
            {
                var prewData = reportData[0];
                List<ElementPki> reportDataWithSortSelObj = new List<ElementPki>();
                List<ElementPki> listDeleteSelRez = new List<ElementPki>();
                //Объединяем подборочные резисторы, елси их параметры равны
                for (int i = 1; i < reportData.Count; i++)
                {
                    if (prewData.Name == reportData[i].Name
                        && prewData.Denotation == reportData[i].Denotation
                        && reportData[i].Comment == prewData.Comment
                        && reportData[i].SelectionObject
                        && prewData.SelectionObject)
                    {
                        prewData.Count += reportData[i].Count;
                        prewData.PosDenotation += "," + reportData[i].PosDenotation;
                        if (!string.IsNullOrEmpty(reportData[i].Remark))
                        {
                            if (!string.IsNullOrEmpty(prewData.Remark))
                            {
                                prewData.Remark += "," + reportData[i].Remark;
                            }
                            else
                            {
                                prewData.Remark = reportData[i].Remark;
                            }  
                        }
                        listDeleteSelRez.Add(reportData[i]);
                    }
                    else
                    {
                        reportDataWithSortSelObj.Add(prewData);
                        prewData = reportData[i];
                    }
                }

                reportDataWithSortSelObj.Add(prewData);


                foreach (var delObj in listDeleteSelRez)
                {
                    reportData.Remove(delObj);
                }
                return reportData;
            }

            private void SetNames(ReportParameters reportParameters)
            {
                foreach (var el in findedElements)
                {
                    el.TU = el.dbElements[0].TU;
                    if (el.dbElements[0].NotWrite)
                    {
                        el.Name = el.dbElements[0].Comment;
                        continue;
                    }

                    TextNormalizer textNorm = new TextNormalizer();

                    var name = textNorm.GetNormalForm(el.dbElements[0].Comment + " " + el.dbElements[0].Description);

                    if (/*name.Contains(textNorm.GetNormalForm("С2-29")) &&*/ el.dbElements.Count > 0 )
                    {
                        if (el.Tks != null && el.Tks.Trim() != string.Empty)
                        {
                            el.dbElements[0].OptionalParameter = el.Tks;                          
                        }
                        else
                        {
                            if (reportParameters.SelectTKS && el.dbElements[0].OptionalParameter == "-1,0-А")
                                el.dbElements[0].OptionalParameter = "*";
                        }
                    }
 
                    if (el.dbElements[0].Value.Contains("*") || el.dbElements[0].Value.Contains("Value"))//(el.Value != null && el.Value != string.Empty)//Пустое значение или Value из БД!!! Если конкретное значение из файла 
                    {
                      //  if ( el.Value == string.Empty)
                        {

                            if (el.dbElements[0].Admission != string.Empty && !el.dbElements[0].Admission.Contains("*"))
                            {
                                SetNamePKI(el, el.dbElements[0].Admission);
                            }
                            else
                            {
                                SetNamePKI(el, el.Admission);
                            }
                        }

                        if (el.Name.Contains(" пФОм"))
                        {
                            el.Name = el.Name.Replace(" пФОм", "");
                        }
                        //else
                        //{
                        //    if (el.PosDenotation.Contains("R21"))
                        //        MessageBox.Show("4 SetNames " + el.Value);
                        //    //  if (el.PosDenotation == "C1")
                        //    //      MessageBox.Show(el.PosDenotation + "\n(" + el.Value + ")\n" + el.Comment + el.Type);
                        //    el.Name = el.dbElements[0].Comment + " " + el.dbElements[0].Description + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                        //}
                    }
                    else
                    {

                        if (el.dbElements[0].Value.ToString().Trim() == string.Empty)
                        {
                            if (el.dbElements[0].Admission != string.Empty && !el.dbElements[0].Admission.Contains("*"))
                            {
                                el.Name = el.dbElements[0].Comment + " " + el.dbElements[0].Description + el.dbElements[0].Admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                            }
                            else
                            {
                                el.Name = el.dbElements[0].Comment + " " + el.dbElements[0].Description + el.Admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                            }                 
                        }
                        else
                        {
                            var description = !string.IsNullOrEmpty(el.dbElements[0].Description) ? el.dbElements[0].Description + "-" : string.Empty;
                            if (el.dbElements[0].Admission != string.Empty && !el.dbElements[0].Admission.Contains("*"))
                            {
                                el.Name = el.dbElements[0].Comment + " " + description + el.dbElements[0].Value.ToString() + el.dbElements[0].Admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                            }
                            else
                            {
                                el.Name = el.dbElements[0].Comment + " " + description + el.dbElements[0].Value.ToString() + el.Admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                            }
                        }

                      
                    }

                }
            } 

            private static void SetNamePKI(ElementPki el, string admission)
            {
                // Для резисторов C2-33 и С2-29 без номиналов. Записываем в результирующий список без номинала
                TextNormalizer textNorm = new TextNormalizer();
                el.TU = el.dbElements[0].TU;
                var name = textNorm.GetNormalForm(el.dbElements[0].Comment + " " + el.dbElements[0].Description);
                if (/*el.dbElements[0].RangeValue.Where(t=>t.Trim().Contains("*")).ToList().Count > 0 &&*/el.Value == null && (name.Contains(textNorm.GetNormalForm("С2-33"))|| name.Contains(textNorm.GetNormalForm("С2-29"))))
                {                  
                    el.Name = el.dbElements[0].Comment + " " + el.dbElements[0].Description + " ..." + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    el.TU = el.dbElements[0].TU;
                    el.RezWithOutNominal = true;
                    return;
                }
                // Для резисторов C2-33 приформировываем дополнительное буквенное обозначение
                if ((name.Contains(textNorm.GetNormalForm("С2-33"))))
                {
                    var description = !string.IsNullOrEmpty(el.dbElements[0].Description) ? el.dbElements[0].Description + "-" : string.Empty;
                    if ((el.dbElements[0].Comment.ToLower().Trim().Contains("резистор") || el.dbElements[0].Comment.ToLower().Trim().Contains("блок")) && el.Value != null)
                    {
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value.Replace("пФ", "") + admission + RezistC233(el.Value.Replace("пФ", ""), admission, el.dbElements[0].OptionalParameter) + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    }
                    else
                    {
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value + admission + RezistC233(el.Value.Replace("пФ", ""), admission, el.dbElements[0].OptionalParameter) + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    }

                    el.TU = el.dbElements[0].TU;
                
                   return;
                }

                // Для резисторов Р1-112 приформировываем дополнительное буквенное обозначение
                if ((name.Contains(textNorm.GetNormalForm("Р1-112"))))
                {
                    var description = !string.IsNullOrEmpty(el.dbElements[0].Description) ? el.dbElements[0].Description + "-" : string.Empty;
                    if (el.dbElements[0].Comment.ToLower().Trim().Contains("резистор") && el.Value != null)
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value.Replace("пФ", "") + admission + RezistR1112(el.Value.Replace("пФ", "")) + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    else
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value + admission + RezistR1112(el.Value.Replace("пФ", "")) + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;

                    el.TU = el.dbElements[0].TU;

                    return;
                }



                if (el.dbElements[0].Comment.ToLower().Trim().Contains("конденсатор") && el.Value != null)
                {
                    var description = !string.IsNullOrEmpty(el.dbElements[0].Description) ? el.dbElements[0].Description + "-" : string.Empty;
                    // Если номинал коненсатора не соответствует указанному диапазону выводим сообщение об ошибке
                    if (!el.dbElements[0].NormalRangeValue.Contains(textNorm.GetNormalForm(el.Value.Replace("Ом", ""))) && !el.dbElements[0].RangeValue.Contains("*"))
                    {
                        el.Name = el.dbElements[0].Comment + " " + description + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                        var notCorrectName = el.dbElements[0].Comment + " " + description + el.Value.Replace("Ом", "") + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                        el.TU = el.dbElements[0].TU;
                        string messageText = "Объект " + el.PosDenotation + " " + notCorrectName + " имеет неверное значение.\nРазрешенный диапазон значений:\n" + string.Join("\n\r", el.dbElements[0].RangeValue.Select(t => t.Trim()).ToList());              
                        Clipboard.Clear();
                        Clipboard.SetText(messageText);
                     //   MessageBox.Show(messageText, "Неверное значение!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        globalLog.AppendLine(messageText);
                        Clipboard.SetText(globalLog.ToString());

                    }
                    else
                    {
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value.Replace("Ом", "") + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                        el.TU = el.dbElements[0].TU;
                    }
                }
                else
                {
                    var description = !string.IsNullOrEmpty(el.dbElements[0].Description) ? el.dbElements[0].Description + "-" : string.Empty;
                    if ((el.dbElements[0].Comment.ToLower().Trim().Contains("резистор") || el.dbElements[0].Comment.ToLower().Trim().Contains("блок")) && el.Value != null)
                    {      
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value.Replace("пФ", "") + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    }
                    else
                    {
                        el.Name = el.dbElements[0].Comment + " " + description + el.Value + admission + el.dbElements[0].OptionalParameter + " " + el.dbElements[0].TU;
                    }

                    el.TU = el.dbElements[0].TU;                 
                }

               
            }

            //Для резисторов С2-33
            //Ом переводить в кОМ
            private static string RezistC233(string value, string admission, string optionalParameters)
            {
                if (optionalParameters.Trim() != string.Empty)
                {
                    return string.Empty;
                }
                var patternValue = @"(\d*(\.|\,)\d*)|\d*";
                var nominal = Regex.Match(value, patternValue).Value.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                nominal = nominal.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                double val = 0;
                try
                {
                    val = Convert.ToDouble(nominal);
                }
                catch
                {
                   globalLog.AppendLine("Некорректный номинал у " + value);
                }
                
                 

                if ((value.ToLower().Trim().Contains("мом") || value.ToLower().Trim().Contains("ком")) && (admission.Contains("5") || admission.Contains("10")))
                {
                    return " А-Д-В";
                }
                else
                {
                    if (val < 1 && value.ToLower().Trim().Contains("ом") && !value.ToLower().Trim().Contains("мом") && !value.ToLower().Trim().Contains("ком") && (admission.Contains("5") || admission.Contains("10")))
                    {
                        return " А-Ж-В";
                    }
                    if  ((value.ToLower().Trim().Contains("ком") || value.ToLower().Trim().Contains("ом")) && (admission.Contains("1") || admission.Contains("2")) && !admission.Contains("10") && !admission.Contains("20"))
                    {
                        return " А-В-В";                    
                    }
                    return " А-Д-В";
                }
            }

            //Для резисторов Р1-112
            private static string RezistR1112(string value)
            {
                var patternValue = @"(\d*(\.|\,)\d*)|\d*";

                var nominal = Regex.Match(value, patternValue).Value.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                nominal = nominal.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                double val = 0;
                try
                {
                    val = Convert.ToDouble(nominal);
                }
                catch
                {
                    globalLog.AppendLine("Некорректный номинал у " + value);
                }
               

                string constPart = string.Empty;

                if (val <= 0.91 && val >= 0.05 && value.ToLower().Trim().Contains("ом") && !value.ToLower().Trim().Contains("ком") && !value.ToLower().Trim().Contains("мом"))
                {
                    constPart = "-М-А";
                }

                if (val <= 91 && val >= 1 && value.ToLower().Trim().Contains("ом") && !value.ToLower().Trim().Contains("ком") && !value.ToLower().Trim().Contains("мом"))
                {
                    constPart = "-Л-А";
                }

                if ((val <= 999 && val >= 100 && value.ToLower().Trim().Contains("ом")) || (val >= 0.1 && val <= 100 && value.ToLower().Trim().Contains("ком")))
                {
                    constPart = "-К-А";
                }

                if ((val <= 999 && val >= 110 && value.ToLower().Trim().Contains("ком")) || (val <= 22 && val >= 0.1 && value.ToLower().Trim().Contains("мом")))
                {
                    constPart = "-Л-А";
                }

                if (val >= 22 && value.ToLower().Trim().Contains("мом"))
                {
                    constPart = "-Н-А";
                }

                return constPart;
            }

            //Заголовок группы элементов В ПЭ
            private Dictionary<List<string>, string> listHeading = new Dictionary<List<string>, string>()
            {
                {new List<string> { "Резистор", "Терморезистор" }, "Резисторы"},
                {new List<string> { "Конденсатор" }, "Конденсаторы"},
                {new List<string> { "Вилка", "Розетка", "Контакт (Розетки", "Контакт (Вилки" }, "Соединители" }
            };

            private int getSortByDigit(string name)
            {
                var regexDigitals = new Regex(@"\d{1,}");
                var matchDigit = regexDigitals.Match(name.Trim());

                try
                {
                    return matchDigit.Success ? Convert.ToInt32(matchDigit.Value.ToString()) : 0;
                }
                catch
                {
                    return 0;
                }
            }

            private void MakeSelectionByTypeReport(List<ElementPki> reportData, ReportParameters reportParameters, LogDelegate m_writeToLog)
            {
                reportData = reportData.Where(t => t.Name != string.Empty).ToList();

                var applicationSessionSetup = new ApplicationSessionSetup
                {
                    Enable3D = true,
                    ReadOnly = false,
                    PromptToSaveModifiedDocuments = false,
                    EnableMacros = true,
                    EnableDOCs = true,
                    DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version13,
                    ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
                };

                if (!Application.InitSession(applicationSessionSetup))
                {
                    return;
                }
                Application.FileLinksAutoRefresh = Application.FileLinksRefreshMode.DoNotRefresh;

                TFlex.Model.Document doc;
                var fileReference = new FileReference();

                if (reportParameters.BeginPage == 1)
                {
                    doc = Application.OpenDocument(_context.ReportFilePath);
                }
                else
                {
                    //Если пользователь указал параметр "Отсчет со страницы" не равный 1, то загружаем соответствующий шаблон
                    FileReferenceObject fileObjectSecondPage = fileReference.FindByRelativePath(
                        @"Служебные файлы\Шаблоны отчётов\Перечень элементов ГОСТ 2.701-2008\Перечень элементов ГОСТ 2.701-2008 вторая и последующие страницы.grb");
                    fileObjectSecondPage.GetHeadRevision(_context.ReportFilePath.Replace("Перечень элементов ГОСТ 2.701-2008.grb", "Перечень элементов ГОСТ 2.701-2008 вторая и последующие страницы.grb"));

                    doc = Application.OpenDocument(_context.ReportFilePath.Replace("Перечень элементов ГОСТ 2.701-2008.grb", "Перечень элементов ГОСТ 2.701-2008 вторая и последующие страницы.grb")); ;
                }

                doc.BeginChanges("Заполнение данными");
                var reportText = doc.GetTextByName("Table");
                reportText.BeginEdit();
                var contentTable = reportText.GetFirstTable();

                m_writeToLog("Количество элементов отчета: " + reportData.Count);

                if (contentTable.HasValue)
                {
                    CadTable cadTable;
                    if (reportParameters.BeginPage == 1)
                    {
                        cadTable = new CadTable(contentTable.Value, reportText, 23, 29);
                    }
                    else
                    {
                        cadTable = new CadTable(contentTable.Value, reportText, 29, 29);
                    }
                    int countStr = 0;

                    if (reportData.Count > 0)
                    {

                        if (reportParameters.IsManyDevice)
                        {
                            var reportDataByDevices = reportData.GroupBy(t => t.DenotationDevice + t.NameDevice).ToDictionary(t => t.Key, t => t.ToList()).OrderBy(t => getSortByDigit(t.Value.First().NameDevice));

                            foreach (var reportDataByDevice in reportDataByDevices)
                            {

                                var emptyRows = cadTable.CountEmptyRowsOnPage();
                                if (emptyRows < 15)
                                    cadTable.NewPage();
                                else
                                {
                                    cadTable.CreateRow();
                                    cadTable.CreateRow();
                                }
                                var rowHeadDevice = cadTable.CreateRow();
                                rowHeadDevice.AddText(reportDataByDevice.Value.First().NameDevice.Replace("+", "≠"));
                                rowHeadDevice.AddText(reportDataByDevice.Value.First().DenotationDevice, CadTableCellJust.Center, TextStyle.ItalicUnderline);
                                WriteRow(reportParameters, m_writeToLog, cadTable, reportDataByDevice.Value);

                                //var наименованиеФайла = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\rd.txt";
                                //File.WriteAllText(наименованиеФайла, string.Join("\r\n", reportDataByDevice.Value.Select(t => t.Name + " " + t.PosDenotation + " " + t.Type)));
                                //Process.Start(наименованиеФайла);
                            }
                        }
                        else
                        {
                            WriteRow(reportParameters, m_writeToLog, cadTable, reportData);
                        }

                    }
                    
                    cadTable.Apply();
                }
                reportText.EndEdit();

                // вставка листа регистрации изменений
                var changelogPage = new Page(doc);
                changelogPage.Rectangle = new Rectangle(0, 0, 210, 297);
                // временная страница для удаления фрагментов
                var tempPage = new Page(doc);
                foreach (Fragment fragment in changelogPage.GetFragments())
                {
                    fragment.Page = tempPage;
                }
                doc.DeletePage(tempPage, new DeleteOptions(true) { DeletePageObjects = true });

               

                FileReferenceObject fileObject = fileReference.FindByRelativePath(
                    @"Служебные файлы\Шаблоны отчётов\Перечень элементов ГОСТ 2.701-2008\Лист регистрации изменений ГОСТ 2.503.grb");               

                fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
                var changelogFragment = new Fragment(new FileLink(doc, fileObject.LocalPath, true));
                changelogFragment.Page = changelogPage;

                var pages = new List<Page>();
                foreach (Page page in doc.GetPages())
                {
                    pages.Add(page);
                }
                for (int i = 0; i < pages.Count; i++)
                {
                    Page page = pages[i];

                    foreach (Fragment fragment in page.GetFragments())
                    {
                        foreach (FragmentVariableValue variable in fragment.GetVariables())
                        {
                            // Удаление связи с переменной документа
                            variable.AttachedVariable = null;
                            switch (variable.Name)
                            {
                                case "$name":
                                    variable.TextValue = reportParameters.NameDevice;
                                    break;
                                case "$denotation":
                                    variable.TextValue = reportParameters.DenotationDevice;
                                    break;
                                case "$oboznach":
                                    variable.TextValue = reportParameters.DenotationDevice;
                                    break;
                                case "$author":
                                    variable.TextValue = reportParameters.AuthorBy;
                                    break;
                                case "$musteredBy":
                                    variable.TextValue = reportParameters.MusteredBy;
                                    break;
                                case "$nControlBy":
                                    variable.TextValue = reportParameters.NControlBy;
                                    break;
                                case "$approvedBy":
                                    variable.TextValue = reportParameters.ApprovedBy;
                                    break;
                                case "$authorDate":
                                    if (reportParameters.AuthorCheck)
                                        variable.TextValue = reportParameters.AuthorByDate.ToShortDateString();
                                    break;
                                case "$musteredByDate":
                                    if (reportParameters.MusteredBCheck)
                                        variable.TextValue = reportParameters.MusteredByDate.ToShortDateString();
                                    break;
                                case "$nControlByDate":
                                    if (reportParameters.NControlByCheck)
                                        variable.TextValue = reportParameters.NControlByDate.ToShortDateString();
                                    break;
                                case "$approvedByDate":
                                    if (reportParameters.ApprovedByCheck)
                                        variable.TextValue = reportParameters.ApprovedByDate.ToShortDateString();
                                    break;
                                case "$numberList":
                                    variable.TextValue = (i + reportParameters.BeginPage).ToString();
                                    break;
                                case "$list":
                                    variable.TextValue = (i + reportParameters.BeginPage).ToString();
                                    break;
                                case "$listov":
                                    variable.TextValue = (pages.Count).ToString();
                                    break;
                                case "$firstUse":
                                    variable.TextValue = reportParameters.FirstUse;
                                    break;
                                case "$litera1":
                                    variable.TextValue = reportParameters.Litera1.ToString();
                                    break;
                                case "$litera2":
                                    variable.TextValue = reportParameters.Litera2.ToString();
                                    break;
                                case "$numberSol":
                                    variable.TextValue = reportParameters.NumberSol;
                                    break;
                                case "$code":
                                    variable.TextValue = reportParameters.Code;
                                    break;
                            }
                        }
                    }
                }

                //Page firstPages = null;

                //foreach(Page page in doc.Pages)
                //{
                //    firstPages = page;
                //    break;
                //}

                //doc.DeletePage(firstPages, new DeleteOptions(true) { DeletePageObjects = true });

                doc.EndChanges();
                if (reportParameters.BeginPage == 1)
                    doc.Save();
                else
                    doc.SaveAs(_context.ReportFilePath);
                doc.Close();
            }

            private void WriteRow(ReportParameters reportParameters, LogDelegate m_writeToLog, CadTable cadTable, List<ElementPki> reportDataByDevice)
            {
                var prewHead = "***";
                var prewLetterPos = string.Empty;
                bool pasteHead = false;
                foreach (var pki in reportDataByDevice)
                {
                    int errorId = 0;
                    try
                    {
                        foreach (var head in listHeading)
                        {
                            errorId = 1;
                            if (head.Key.Any(t => pki.Name.Contains(t) && !string.IsNullOrEmpty(pki.TU)) && head.Key.All(t => t != prewHead) && prewLetterPos.Trim() != pki.LetterPosDenotation.Trim())
                            {
                                if (reportDataByDevice.First() != pki && head.Key.All(t => t != "Конденсатор"))
                                {
                                    var emptyRows = cadTable.CountEmptyRowsOnPage();
                                    if (emptyRows < 15)
                                        cadTable.NewPage();
                                    else
                                    {
                                        cadTable.CreateRow();
                                        cadTable.CreateRow();
                                        errorId = 2;
                                    }
                                }
                                if (head.Key.Any(t => t == "Конденсатор"))
                                {
                                    cadTable.CreateRow();
                                    cadTable.CreateRow();
                                }
                                errorId = 3;
                                if (head.Key.Any(t => pki.Name.Contains(t)))
                                {
                                    var condRez = reportDataByDevice.Where(t => head.Key.Any(s => t.Name.Contains(s))).ToList();
                                    var tuCondRezs = condRez.GroupBy(t => t.TU).Select(t => t.First());

                                    foreach (var cR in tuCondRezs)
                                    {
                                        var rowHead = cadTable.CreateRow();
                                        rowHead.AddText("");

                                        Regex regexType = new Regex(@"(^[А-ЯA-Z]{1,2}\d{0,2}([А-ЯA-Z]{1})?-(\d{1,3}))|(^\d{0,1}[А-ЯA-Z]{3}\d{0,2}[А-ЯA-Z]{0,1}\d{0,2})");
                                        errorId = 4;
                                        var nameElementWithOutKey = cR.Name;
                                        head.Key.ForEach(t => nameElementWithOutKey = nameElementWithOutKey.Replace(t, string.Empty));
                                        var typeCondRezs = regexType.Match(nameElementWithOutKey.Trim()).Value;

                                        if (cR.Name.Contains("Терморезистор"))
                                        {
                                            string headTerm;
                                            if (reportDataByDevice.Where(t => t.Name.Contains("Терморезистор")).ToList().Count > 1)
                                                headTerm = "Терморезисторы";
                                            else
                                                headTerm = "Терморезистор";
                                            rowHead.AddText(headTerm + " " + typeCondRezs + " " + cR.TU, CadTableCellJust.Left, TextStyle.ItalicUnderline);
                                            continue;
                                        }
                                        rowHead.AddText(head.Value + " " + typeCondRezs + " " + cR.TU, CadTableCellJust.Left, TextStyle.ItalicUnderline);
                                    }
                                    errorId = 5;
                                }
                                else
                                {
                                    cadTable.CreateRow();
                                    var rowHead = cadTable.CreateRow();
                                    rowHead.AddText("");
                                    rowHead.AddText(head.Value, CadTableCellJust.Center, TextStyle.ItalicUnderline);
                                }
                                errorId = 6;

                                cadTable.CreateRow();
                                prewHead = head.Key.First();
                                pasteHead = true;
                                break;
                            }
                            else
                                pasteHead = false;
                        }
                        errorId = 7;
                        if (!pasteHead && prewLetterPos.Trim() != pki.LetterPosDenotation.Trim())
                        {
                            for (int i = 0; i < reportParameters.CountHeadRows; i++)
                            {
                                cadTable.CreateRow();
                            }
                        }
                        errorId = 8;

                        var listPosDenotation = pki.ModifyPosDenotation();

                        //Вставка * в Поз обознач подборочных резисторов
                        if (pki.SelectionObject)
                        {
                            for (int i = 0; i < listPosDenotation.Count; i++)
                            {
                                listPosDenotation[i] = listPosDenotation[i].Replace(",", "*,").Replace("-", "*-");
                                if (listPosDenotation[i].Last() != ',')
                                    listPosDenotation[i] += "*";
                            }
                        }
                        errorId = 9;

                        if (CountLineName(prewHead, pki) > cadTable.CountEmptyRowsOnPage() || listPosDenotation.Count > cadTable.CountEmptyRowsOnPage())
                        {
                            for (int i = 0; i < cadTable.CountEmptyRowsOnPage(); i++)
                            {
                                cadTable.CreateRow();
                            }
                        }

                        errorId = 10;
                        var row = cadTable.CreateRow();

                        if (listPosDenotation.Count == 1)
                        {
                            row.AddText(listPosDenotation[0], textStyle: TextStyle.Italic);
                        }
                        else
                        {
                            if (listPosDenotation.Count == 0)
                            {
                                row.AddText("");
                            }
                            else
                            {
                                for (int i = 0; i < listPosDenotation.Count; i++)
                                {
                                    if (i == 0)
                                    {
                                        row.AddText(listPosDenotation[i], textStyle: TextStyle.Italic);
                                        WriteRow(cadTable, prewHead, pki, row, listPosDenotation.Count > 1 ? listPosDenotation[i] : string.Empty);

                                        m_writeToLog("Запись: " + pki.Name);
                                        prewLetterPos = pki.LetterPosDenotation;
                                    }
                                    else
                                    {
                                        var newRow = cadTable.CreateRow();
                                        newRow.AddText(listPosDenotation[i], textStyle: TextStyle.Italic);
                                        WriteRow(cadTable, prewHead, pki, newRow, listPosDenotation[i]);
                                    }
                                }
                                continue;
                            }
                        }
                        errorId = 11;
                        WriteRow(cadTable, prewHead, pki, row);
                        m_writeToLog("Запись: " + pki.Name);
                        prewLetterPos = pki.LetterPosDenotation;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка!!!" + pki.Name + " " + errorId + ". Обратитесь в отдел 911. Тел. 26-00", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }

            //пара
            //метр Короткое имя справочника "Группы и пользователи"
            private static readonly Guid UserShortNameGuid = new Guid("76a97c36-f2d6-49ad-abb0-f5fdc91c071b");

            private void MakeFileStr(List<ElementPki> reportData, List<ElementPki> linkedObjects, ReportParameters reportParameters, List<ElementPki> objectsForSpecification, LogDelegate m_writeToLog)
            {
                // Clipboard.SetText(string.Join("\r\n", reportData.Select(obj=> obj.Name + ";" + obj.Denotation + ";" + obj.Type + ";" + obj.Count + ";" + obj.PosDenotation)));
                //  MessageBox.Show(reportData.Count + string.Join("\r\n", reportData.OrderBy(t=>t.Name).Select(obj => obj.Name + ";" + obj.Denotation + ";" + obj.Type + ";" + obj.Count + ";" + obj.PosDenotation)));
                //StringBuilder strB = new StringBuilder();

                //foreach (var obj in objectsForSpecification)
                //{
                //    strB.AppendLine(obj.PosDenotation + " " + obj.Type + " " + obj.Value + " " + obj.Admission);
                //}

                //if (strB.ToString() != string.Empty)
                //{
                //    Clipboard.SetText(strB.ToString());
                //    MessageBox.Show(strB.ToString());
                //}

                //foreach (var lo in linkedObjects)
                //{
                //    strB.AppendLine(lo.PosDenotation + " " + lo.Name + " " + lo.Denotation + " " + lo.Comment);
                //}
                //   Clipboard.SetText(strB.ToString());
                //  MessageBox.Show(strB.ToString());

                //strB = new StringBuilder();
                //  linkedObjects.AddRange(objectsForSpecification);
                // var objectsForSp = FindObjectForSpecificationInNomenclature(objectsForSpecification, m_writeToLog);
                //ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(NomenclatureReference);

                //foreach (var objForSp in objectsForSp)
                //    {
                //        if (objForSp.NomObj != null)
                //        {

                //            ClassObject classObject = referenceNomenclatureInfo.Classes.Find(objForSp.NomObj.Class.Id);
                //            objForSp.Type = classObject.ToString();
                //        }
                //    }

                // var kont = string.Join("\r\n", objectsForSp.Select(t => "Name " + t.Name + " nom " + t.NomObj  + " type " + t.Type));
                //      Clipboard.SetText(kont);
                //    MessageBox.Show(kont);

                //  var kont2 = string.Join("\r\n", objForSp.Select(t => "Name " + t.Name + " nom " + t.NomObj + " type " + t.Type));
                //  Clipboard.SetText("r" + kont2);
                //   MessageBox.Show("r" + kont2);
                var resultPkiElements = new List<SortPKIElements>();

                int errorid = 0;
                try
                {
                    errorid = 1;
                    List<ElementPki> fileData = new List<ElementPki>();
                    foreach (var obj in reportData)
                    {
                        if (obj.NomObj != null)
                        {
                            List<ElementPki> objectsForWrite = new List<ElementPki>();

                            if (obj.Type == "Сборочная единица")
                            {
                                objectsForWrite = fileData.Where(t => (t.Name + t.Denotation) == (obj.NomObj[ParametersInfo.Name].Value.ToString() + obj.NomObj[ParametersInfo.Denotation].Value.ToString())
                                    && t.Comment.ToLower().Trim() == obj.Comment.ToLower().Trim()).ToList();

                            }
                            else
                            {
                                objectsForWrite = fileData.Where(t => (t.Name + t.Denotation) == (obj.NomObj[NameOtherItemsGuid].Value.ToString() + obj.NomObj[DenotationItemsGuid].Value.ToString())
                                                               && t.Comment.ToLower().Trim() == obj.Comment.ToLower().Trim()).ToList();

                            }

                            if (objectsForWrite.Count > 0)
                            {
                                objectsForWrite.FirstOrDefault().PosDenotation += "," + obj.PosDenotation;
                                objectsForWrite.FirstOrDefault().Count += obj.Count;

                            }
                            else
                            {
                                if (obj.Type == "Прочее изделие")
                                {
                                    obj.OldName = obj.Name;
                                    obj.Name = obj.NomObj[NameOtherItemsGuid].Value.ToString();
                                    obj.Denotation = obj.NomObj[DenotationItemsGuid].Value.ToString();
                                    fileData.Add(obj);
                                }
                                else
                                {
                                    obj.Name = obj.NomObj[ParametersInfo.Name].Value.ToString();
                                    obj.Denotation = obj.NomObj[ParametersInfo.Denotation].Value.ToString();
                                    fileData.Add(obj);
                                }
                            }
                        }
                    }

                        if (linkedObjects != null)
                        fileData.AddRange(linkedObjects);

                    //if (objectsForSpecification != null)
                    //    fileData.AddRange(objectsForSpecification);


                    string filePath = string.Empty;
                    if (reportParameters.filePath.Contains(".atr"))
                        filePath = reportParameters.filePath.Replace(".atr", " " + ClientView.Current.GetUser()[UserShortNameGuid].Value + " " + DateTime.Now.ToString().Replace(":","-") + ".str");
                    if (reportParameters.filePath.Contains(".xls"))
                        filePath = reportParameters.filePath.Replace(".xls", " " + ClientView.Current.GetUser()[UserShortNameGuid].Value + " " + DateTime.Now.ToString().Replace(":", "-") + ".str");
                    if (reportParameters.filePath.Contains(".xlsx"))
                        filePath = reportParameters.filePath.Replace(".xlsx", " " + ClientView.Current.GetUser()[UserShortNameGuid].Value + " " + DateTime.Now.ToString().Replace(":", "-") + ".str");
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    List<SortPKIElements> sortFileData = new List<SortPKIElements>();
                    List<ElementPki> documentsData = new List<ElementPki>();
                    List<ElementPki> complexesData = new List<ElementPki>();
                    List<ElementPki> assemblyData = new List<ElementPki>();
                    List<ElementPki> detailsData = new List<ElementPki>();
                    List<ElementPki> standartItemsData = new List<ElementPki>();
                    List<ElementPki> otherItemsData = new List<ElementPki>();
                    List<ElementPki> materialsData = new List<ElementPki>();
                    List<ElementPki> komplektsData = new List<ElementPki>();
                    errorid = 4;
                    foreach (var obj in fileData)
                    {
                        if (obj.Type == "Документы")
                        {
                            documentsData.Add(obj);
                        }
                        if (obj.Type == "Комплекс")
                        {
                            complexesData.Add(obj);
                        }
                        if (obj.Type == "Сборочная единица")
                        {
                            assemblyData.Add(obj);
                        }
                        if (obj.Type == "Деталь")
                        {
                            detailsData.Add(obj);
                        }
                        if (obj.Type == "Стандартное изделие")
                        {
                            standartItemsData.Add(obj);
                        }
                        if (obj.Type == "Прочее изделие")
                        {
                            otherItemsData.Add(obj);
                        }
                        if (obj.Type == "Материал")
                        {
                            materialsData.Add(obj);
                        }
                        if (obj.Type == "Комплект")
                        {
                            komplektsData.Add(obj);
                        }
                    }
                    errorid = 5;
                    foreach (var obj in documentsData.OrderBy(t => t.Denotation).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    foreach (var obj in complexesData.OrderBy(t => t.Denotation).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    foreach (var obj in assemblyData.OrderBy(t => t.Denotation).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    foreach (var obj in detailsData.OrderBy(t => t.Denotation).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    foreach (var obj in standartItemsData.OrderBy(t => t.Name).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    errorid = 6;
                    List<SortPKIElements> otherItemsFileData = new List<SortPKIElements>();

                    foreach (var obj in otherItemsData.OrderBy(t => t.Name).ToList())
                    {

                        otherItemsFileData.Add(new SortPKIElements(obj));

                    }
                    otherItemsFileData = otherItemsFileData
                        .OrderBy(t => t.ElPKI.LetterPosDenotation)
                        .ThenBy(t => t.FirstPartName)
                        .ThenBy(t => t.LetterSecondPartNameKond)
                        .ThenBy(t => t.SecondPartName)
                        .ThenBy(t => t.SecondNameMcx)
                        .ThenBy(t => t.ThirdPartName)
                        .ThenBy(t => t.Measure)
                        .ThenBy(t => t.Nominal)
                        .ToList();

                        sortFileData.AddRange(otherItemsFileData);


                    errorid = 6;
                    foreach (var obj in materialsData.OrderBy(t => t.Name).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    foreach (var obj in komplektsData.OrderBy(t => t.Denotation).ToList())
                    {
                        sortFileData.Add(new SortPKIElements(obj));
                    }
                    errorid = 7;
                    List<string> lStr = new List<string>();
                    using (FileStream fs = File.Create(filePath))
                    {
                        errorid = 77;
                        foreach (var obj in sortFileData)
                        {
                            errorid = 777;
                            ElementPki pkiElement = new ElementPki();
                            List<string> nominals = new List<string>();
                            List<string> nominalsWithOutBlank = new List<string>();
                            List<ElementPki> selectionObjects = new List<ElementPki>();
                            errorid = 8;
                            string comment = string.Empty;
                            foreach (var pos in obj.ElPKI.ModifyPosDenotation())
                            {
                                comment += pos;
                            }
                            errorid = 9;
                            if (obj.ElPKI.Comment.Trim() != string.Empty && comment.Trim() != string.Empty)
                            {
                                // comment += "#" + obj.ElPKI.Comment.Replace(";", "#");

                                if (obj.ElPKI.SelectionObject)
                                {
                                    //Формируем список номиналов для подборочных объектов
                                    nominals.AddRange(SetCommentSelectionObject(obj.ElPKI.Comment).Split(';').ToList());
                                 //   comment = SetCommentSelectionObject(obj.ElPKI.Comment).Replace(";", "#");
                                }
                                else
                                {
                                    comment = obj.ElPKI.Comment.Replace(";", "#");
                                }
                            }
                            else
                            {
                                if (obj.ElPKI.Comment.Trim() != string.Empty)
                                {
                                    if (obj.ElPKI.SelectionObject)
                                    {
                                        //Формируем список номиналов для подборочных объектов
                                        nominals.AddRange(SetCommentSelectionObject(obj.ElPKI.Comment).Split(';').ToList());
                                       // comment = SetCommentSelectionObject(obj.ElPKI.Comment).Replace(";", "#");
                                    }
                                    else
                                    {
                                        comment = obj.ElPKI.Comment.Replace(";", "#");
                                    }
                                }
                            }
                            errorid = 10;
                            if (obj.ElPKI.Preparation.Trim() != string.Empty)
                            {
                             //   obj.ElPKI.Name += " (Заготовка для " + obj.ElPKI.Preparation.Trim() + ")";
                            }

                            string line = string.Empty;
                            resultPkiElements.Add(obj);
                            //Запись в файл не подборочных объектов
                            if (!obj.ElPKI.SelectionObject)
                            {
                                errorid = 11;
                                line = string.Format(@"{0};{1};{2};{3};{4}", obj.ElPKI.Name, obj.ElPKI.Denotation, obj.ElPKI.Type, obj.ElPKI.Count, comment);
                               // lStr.Add(line);
                                lStr.Add(Convert.ToBase64String(Encoding.UTF8.GetBytes(line)));
                               
                            }
                            else
                            {
                                errorid = 12;
                                // obj.ElPKI - основной подборочный объект                 
                                obj.ElPKI.Comment = comment;

                                pkiElement.Name = obj.ElPKI.Name;
                                pkiElement.Denotation = obj.ElPKI.Denotation;
                                pkiElement.Type = obj.ElPKI.Type;
                                pkiElement.Count = obj.ElPKI.Count;
                                pkiElement.Comment = obj.ElPKI.Comment.Replace(";", "#");

                                Regex regexDigitals = new Regex(@"(\d*(\.|\,)\d*)|\d*");
                                errorid = 13;
                                for (int i = 0; i < nominals.Count; i++)
                                {
                                    errorid = 14;
                                    var digit = regexDigitals.Match(nominals[i].Trim()).Value.ToString();
                                    var razm = nominals[i].Replace(digit, "");
                                    nominals[i] = digit.Trim() + " " + razm.Trim();
                                    nominalsWithOutBlank.Add(digit + razm);
                                }

                                //==============================================================================================
                                errorid = 15;
                                //Регулярное выражение для выявления номиналов 
                                string nomianalPattern = @"(?<=-)((\d+)[.,]?\d*)(\s?(пФ|мкФ|Ом|кОм|МОм))";
                                Regex regexNominal = new Regex(nomianalPattern);
                                //Номинал подборочного объекта
                                var beginNominal = regexNominal.Match(obj.ElPKI.Name).Value;

                                //Цикл по номиналам для создания нового объекта в соответствии с выявленным номиналом
                                for (int i = 0; i < nominals.Count; i++)
                                {
                                    try
                                    {
                                        errorid = 16;
                                        //Создание нового подборочного объекта
                                        ElementPki selObj = new ElementPki();
                                        //Формирование нового наименования подборочного объекта с заменой на новый номинал
                                        selObj.Name = obj.ElPKI.Name.Replace(beginNominal, nominals[i]);
                                        selObj.NameWithOutBlank = obj.ElPKI.Name.Replace(beginNominal, nominalsWithOutBlank[i]);
                                        selObj.Denotation = obj.ElPKI.Denotation;
                                        selObj.Type = obj.ElPKI.Type;
                                        selObj.Count = obj.ElPKI.Count;
                                        selObj.Comment = comment;
                                        selObj.SelectionObject = true;
                                        selectionObjects.Add(selObj);
                                    }
                                    catch
                                    {
                                       // MessageBox.Show("Что-то пошло не так " + nominals[i]);
                                    }
                                }
                                errorid = 17;
                                Regex regexDigital = new Regex(@"\d{1,}");
                                obj.ElPKI.MainSelectionObject = true;
                                selectionObjects.Add(obj.ElPKI);
                                //Сортировка подборочных объектов
                                selectionObjects = selectionObjects.OrderBy(t => t.Name).ThenBy(t => t.Comment).ThenBy(t => regexDigital.Match(t.Comment).Value.ToString().Trim()).ToList();
                                //List<ElementPKI> removeSelectionsObjects = new List<ElementPKI>();

                                //string objForSP = string.Join("\r\n", selectionObjects.Select(t => t.Name + " " + t.Comment));
                                //Clipboard.SetText(objForSP);
                                //MessageBox.Show(objForSP);
                                 

                                //for (int i = 1; i < selectionObjects.Count; i++)
                                //{
                                //    if (selectionObjects[i].Name == selectionObjects[i - 1].Name && selectionObjects[i].SelectionObject == selectionObjects[i - 1].SelectionObject/*&& selectionObjects[i].Type == selectionObjects[i - 1].Type*/ && selectionObjects[i].Comment != selectionObjects[i - 1].Comment)
                                //    {
                                //        selectionObjects[i].Count += selectionObjects[i - 1].Count;
                                //        selectionObjects[i].Comment += ", " + selectionObjects[i - 1].Comment;
                                //        removeSelectionsObjects.Add(selectionObjects[i - 1]);
                                //    }
                                //}
                                ////Удаление повторяющихся подборочных объектов
                                //foreach (var delObj in removeSelectionsObjects)
                                //{
                                //    selectionObjects.Remove(delObj);
                                //}

                                //string objForSP1 = string.Join("\r\n", selectionObjects.Select(t => t.Name + " " + t.Comment));
                                //Clipboard.SetText(objForSP1);
                                //MessageBox.Show(objForSP1);

                                errorid = 18;
                                foreach (var selObj in selectionObjects)
                                {
                                    if (selObj.MainSelectionObject)
                                        line = string.Format(@"{0};{1};{2};{3};{4}", selObj.Name, selObj.Denotation, selObj.Type + "***", selObj.Count, selObj.Comment);
                                    else
                                        line = string.Format(@"{0};{1};{2};{3};{4}", selObj.Name, selObj.Denotation, selObj.Type + "*", selObj.Count, selObj.Comment);
                                   // lStr.Add(line);
                                    lStr.Add(Convert.ToBase64String(Encoding.UTF8.GetBytes(line)));
                                }
                                //===============Завершение обработки подборочных объектов===================================
                            }     
                        }
                    }
                    //   Clipboard.SetText(string.Join("\r\n", lStr.ToList()));
                    //  MessageBox.Show( string.Join("\r\n", lStr.ToList()));
                    MakeFileExcel(resultPkiElements.Where(t => t.ElPKI.Type.Contains("Прочее")).OrderBy(t => t.ElPKI.Name).ToList(), reportParameters);
                    errorid = 19;
                    File.WriteAllLines(filePath, lStr.ToArray(), UnicodeEncoding.UTF8);

                }
                catch
                {
                    MessageBox.Show("Сформирвоать файл *.str нельзя код ошибки " + errorid + ". Обратитесь в отдел 911", "Файл *.str сформирован не будет", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }           
            }

            private void MakeFileExcel(List<SortPKIElements> pkiElements, ReportParameters reportParameters)
            {
                if (!reportParameters.IsLoadKre)
                {
                    return;
                }
                var xls = new Xls();
                int row = 1;

                xls[1, row].Value = "Наименование (T-FLEX DOCs)";
                xls[2, row].Value = "Обозначение (T-FLEX DOCs)";
                xls[3, row].Value = "Количество";
                xls[4, row].Value = "Наименование (1С:ERP)";
                xls[5, row].Value = "Обозначение (1С:ERP)";
                xls[6, row].Value = "Наименования равны";

                xls[1, row, 6, 1].SetBorders(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                                                                                              Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                                                                                              Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                                                                                              Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                                                                                              Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical);

                row++;

                foreach (var pkiElement in pkiElements)
                {
                    xls[1, row].Value = pkiElement.ElPKI.Name;
                    xls[2, row].Value = pkiElement.ElPKI.Denotation;
                    xls[3, row].Value = pkiElement.ElPKI.Count;
                    if (pkiElement.ElPKI.NomObj != null)
                    {
                        var эталонноеЗначение = ПолучитьЭталонныеЗначения(pkiElement.ElPKI.NomObj);
                        xls[4, row].Value = string.IsNullOrEmpty(эталонноеЗначение.Key) && string.IsNullOrEmpty(эталонноеЗначение.Value) ? "неотнормализован" : эталонноеЗначение.Key; 
                        xls[5, row].Value = string.IsNullOrEmpty(эталонноеЗначение.Key) && string.IsNullOrEmpty(эталонноеЗначение.Value) ? "неотнормализован" : эталонноеЗначение.Value;
                        xls[6, row].Value = pkiElement.ElPKI.Name == эталонноеЗначение.Key && pkiElement.ElPKI.Denotation == эталонноеЗначение.Value;
                    }
                    else
                    {
                        xls[4, row].Value = string.Empty;
                        xls[5, row].Value = string.Empty;
                        xls[6, row].Value = false;
                    }
                    xls[1, row, 6, 1].SetBorders(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                                                                                               Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                                                                                               Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                                                                                               Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight,
                                                                                               Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical);
                    row++;
                }
                xls.AutoWidth();
                var filePath = Path.GetDirectoryName(reportParameters.filePath) + 
                    Path.DirectorySeparatorChar + "Список КРЭ " + 
                    ClientView.Current.GetUser()[UserShortNameGuid].Value + " " + 
                    DateTime.Now.ToString().Replace(":", "-") + ".xlsx";
                xls.SaveAs(filePath);
                xls.Close();
            }

            private KeyValuePair<string, string> ПолучитьЭталонныеЗначения(ReferenceObject refObject)
            {
                var эталонноеНаименование = string.Empty;
                var эталонноеОбозначение = string.Empty;

                if (refObject.Class.IsInherit(Guids.SpecReferencesTypes.OtherItems) && refObject[Guids.OtherItemsParameters.Нормализован].GetString() == "да")
                {
                    эталонноеНаименование = refObject[Guids.OtherItemsParameters.НаименованиеЭталонное].GetString();
                    эталонноеОбозначение = refObject[Guids.OtherItemsParameters.ОбозначениеЭталонное].GetString();
                }

                if (refObject.Class.IsInherit(Guids.НоменклатураИИзделия.Типы.СтандартноеИзделие) && refObject[Guids.StandartItemsParameters.Нормализован].GetString() == "да")
                {
                    эталонноеНаименование = refObject[Guids.StandartItemsParameters.НаименованиеЭталонное].GetString();
                    эталонноеОбозначение = refObject[Guids.StandartItemsParameters.ОбозначениеЭталонное].GetString();
                }
                return new KeyValuePair<string, string>(эталонноеНаименование, эталонноеОбозначение);
            }

            private string SetCommentSelectionObject(string comment)
            {
                string newComment = string.Empty;
                List<string> comments = new List<string>();
                string patternDimension = @"Ом|кОм|МОм";
                Regex regexDimension = new Regex(patternDimension);

                //Разделяем комментарий по разделителю ';'
                var nominals = comment.Split(';').ToList();
                var dimensionMatches = regexDimension.Matches(comment);

                //Если количество номиналов для подборочных резисторов больше найденных размерностей
                if (nominals.Count > dimensionMatches.Count)
                {
                    int beginIndex = 0;
                    for (int i = 0; i < nominals.Count; i++)
                    {
                        if (regexDimension.Match(nominals[i]).Success)
                        {
                            var dimension = regexDimension.Match(nominals[i]).Value;
                            for (int j = beginIndex; j < i; j++)
                            {
                                if (!regexDimension.Match(nominals[j]).Success)
                                {
                                    nominals[j] = nominals[j] + " " + dimension;
                                }
                                else
                                {
                                    comments.Add(nominals[j]);
                                }
                            }
                            beginIndex = i + 1;
                        }

                    }
                }
                return string.Join(";", nominals);
            }

            private void WriteRow(CadTable cadTable, string prewHead, ElementPki pki, CadTableRow row, string posDenotation = "")
            {
                CadTableRow lastRow;
                if ((pki.Name.Contains(prewHead) || (prewHead == "Резистор" && (pki.Name.Contains("Терморезистор")))) && !string.IsNullOrEmpty(pki.TU))
                {
                    if (pki.Name.Contains("Конденсатор") || pki.Name.Contains("Резистор") || pki.Name.Contains("Терморезистор"))
                    {
                        string name;
                        if (pki.Name.Contains("Терморезистор"))
                        {
                            if (!string.IsNullOrEmpty(pki.TU.Trim()))
                            {
                                name = pki.Name.Replace(pki.TU.Trim(), "").Trim().Replace("Терморезистор", "").Trim();
                            }
                            else
                            {
                                name = pki.Name.Trim().Replace("Терморезистор", "").Trim();
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(pki.TU.Trim()))
                            {
                                name = pki.Name.Replace(pki.TU.Trim(), "").Trim().Replace(prewHead, "").Trim();
                            }
                            else
                            {
                                name = pki.Name.Trim().Replace(prewHead, "").Trim();
                            }
                        }

                        lastRow = WriteName(cadTable, row, name);
                    }
                    else
                    {
                        var name = pki.Name.Replace(prewHead, "").Trim();
                        lastRow = WriteName(cadTable, row, name);
                    }
                }
                else
                {
                    var name = pki.Name + " " + pki.Denotation;
                    lastRow = WriteName(cadTable, row, name);
                }
                var count = string.IsNullOrEmpty(posDenotation) ? pki.Count.ToString() : ПолучитьКоличествоИзПримечания(posDenotation).ToString();
                lastRow.AddText(count.ToString(), textStyle: TextStyle.Italic);
                var comment = pki.Comment.Trim();
                if (string.IsNullOrEmpty(comment))
                {
                    comment = pki.Remark;
                }
                else
                {
                    if (!string.IsNullOrEmpty(pki.Remark))
                    {
                        comment = comment + ", " + pki.Remark;
                    }
                }
                lastRow.AddText(comment, textStyle: TextStyle.Italic);
            }

            private int ПолучитьКоличествоИзПримечания(string примечание)
            {
                var remarkWithOutLetter = примечание.Replace("*", "");
                var regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
                var regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
                var regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
                var количество = 0;
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    var matchValue = match.Value.ToString();
                    var matches = regexDigital.Matches(matchValue);
                    количество += Convert.ToInt32(matches[1].Value) - Convert.ToInt32(matches[0].Value) + 1;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(matchValue, "");
                }
                количество += regexDigitalOne.Matches(remarkWithOutLetter).Count;
                return количество;
            }

            private int CountLineName(string prewHead, ElementPki pki)
            {
                CadTableRow lastRow;
                if (pki.Name.Contains(prewHead) || (prewHead == "Резистор" && (pki.Name.Contains("Терморезистор"))))
                {
                    if (pki.Name.Contains("Конденсатор") || pki.Name.Contains("Резистор") || pki.Name.Contains("Терморезистор"))
                    {
                        string name;
                        if (pki.Name.Contains("Терморезистор"))
                        {
                            if (!string.IsNullOrEmpty(pki.TU.Trim()))
                            {
                                name = pki.Name.Replace(pki.TU.Trim(), "").Trim().Replace("Терморезистор", "").Trim();
                            }
                            else
                            {
                                name = pki.Name.Trim().Replace("Терморезистор", "").Trim();
                            }
                            return ListName(name).Count();
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(pki.TU.Trim()))
                            {
                                name = pki.Name.Replace(pki.TU.Trim(), "").Trim().Replace(prewHead, "").Trim();
                            }
                            else
                            {
                                name = pki.Name.Trim().Replace(prewHead, "").Trim();
                            }
                            return ListName(name).Count();
                        }

                    }
                    else
                    {
                        var name = pki.Name.Replace(prewHead, "").Trim();
                        return ListName(name).Count();

                    }
                }
                else
                {
                    var name = pki.Name;
                    return ListName(name).Count();

                }           
            }

            private CadTableRow WriteName(CadTable cadTable, CadTableRow row, string name)
            {
                var listName = ListName(name);
                if (listName.Count == 1)
                {
                    row.AddText(name, textStyle: TextStyle.Italic);
                    return row;
                }
                else
                {
                    row.AddText(listName[0], textStyle: TextStyle.Italic);

                    for (int j = 1; j < listName.Count; j++)
                    {
                        var rowName = cadTable.CreateRow();
                        rowName.AddText("");
                        rowName.AddText(listName[j], textStyle: TextStyle.Italic);
                        if (listName[j] == listName.Last())
                            return rowName;
                    }
                }
                return row;
            }

            public  List<string> ListName(string name)
            {
                string patternTU = @"(?:((ГОСТ|ОСТ|ТУ)+[\s,\S]+))$";
                Regex regexTU = new Regex(patternTU);
                int indxTU = -1;
                int lengthTU = -1;
                var tu = regexTU.Match(name).Value;
                if (regexTU.Match(name).Success)
                {                      
                    indxTU = name.IndexOf(tu);
                    lengthTU = tu.Length;
                }
                List<string> listResult = new List<string>();
                string bufStr = string.Empty;
                listResult.Add(string.Empty);
                int lengthLine = 50;

                for (int i = 0; i < name.Length; i++)
                {
                    bufStr += name[i].ToString();
                    if (indxTU > 0)
                    {
                        if ((name[i].ToString() == "," || name[i].ToString() == " ") && (name.IndexOf("ТУ") != i + 1))//(i < indxTU || i> indxTU + lengthTU) )
                        {
                            if ((listResult[listResult.Count - 1].Length + bufStr.Length) <= lengthLine || listResult[listResult.Count - 1] == string.Empty)
                            {
                                listResult[listResult.Count - 1] += bufStr;
                                bufStr = string.Empty;
                            }
                            else
                            {
                                listResult.Add(bufStr);
                                bufStr = string.Empty;
                            }

                        }
                    }
                    else
                    {
                        if (name[i].ToString() == "," || name[i].ToString() == " " )
                        {
                            if ((listResult[listResult.Count - 1].Length + bufStr.Length) <= lengthLine || listResult[listResult.Count - 1] == string.Empty)
                            {
                                listResult[listResult.Count - 1] += bufStr;
                                bufStr = string.Empty;
                            }
                            else
                            {
                                listResult.Add(bufStr);
                                bufStr = string.Empty;
                            }

                        }
                    }
                }

                if ((listResult[listResult.Count - 1].Length + bufStr.Length) <= lengthLine)
                {
                    listResult[listResult.Count - 1] += bufStr;
                }
                else
                {
                    if (listResult.Count() == 1 && listResult[0].Trim() == string.Empty)
                    {
                        listResult[listResult.Count - 1] += bufStr;
                    }
                    else
                    {
                        listResult.Add(bufStr);
                    }
                }

                listResult.RemoveAll(t => string.IsNullOrWhiteSpace(t));
                return listResult.Count != 0 ? listResult.Select(t => t.Trim()).ToList() : new List<string> { name };
            }

            // Чтение данных для отчета
            public List<ElementPki> ReadData(ReportParameters reportParameters, LogDelegate logDelegate, bool writeAll, out List<ElementPki> objForSpec)
            {
                objForSpec = new List<ElementPki>();
                StringBuilder strB = new StringBuilder();   
                List<ElementPki> table = new List<ElementPki>();
                //Объекты для спецификации, невходящие в перечень элементов (KONTAKT и ТОЧКИ)       

                if (reportParameters.filePath.Contains(".atr") || reportParameters.filePath.Contains(".xls"))
                {
                    for (int i = 0; i < reportParameters.FileData.Count; i++)
                    {
                        try
                        {
                            var elementPki = new ElementPki();
                            elementPki.PosDenotation = reportParameters.FileData[i].Designator;
                            elementPki.Type = reportParameters.FileData[i].Type;
                            elementPki.Count = 1;
                            elementPki.Value = reportParameters.FileData[i].Value;
                            elementPki.Admission = reportParameters.FileData[i].Admission;
                            elementPki.Tks = reportParameters.FileData[i].Tks;
                            elementPki.DenotationDevice = reportParameters.FileData[i].Наименование;
                            elementPki.NameDevice = reportParameters.FileData[i].Обозначение;
                            elementPki.Remark = reportParameters.FileData[i].Примечание;

                            if (reportParameters.FileData[i].Type.Contains("KONTAKT") || reportParameters.FileData[i].Type.Contains("ТОЧК"))
                            {
                                objForSpec.Add(elementPki);
                            }
                            else
                            {
                                table.Add(elementPki);
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Обратитесь в отдел 911. Перечень элементов не может быть сформирован!\n"
                                  + "Некорректные данные. Ошибка в строке № " + i + "\n" + reportParameters.FileData[i],
                                  "Ошибка в исходном файле!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                logDelegate("Из файла считано " + table.Count + " объектов");

                return table;
            }

            //Добавление допуска
            public void AddAdmission(List<ElementPki> elements)
            {
                List<ElementPki> withOutAdmissionObjects = new List<ElementPki>();
                withOutAdmissionObjects.AddRange(elements.Where(t => t.Admission.ToLower().Contains("допуск") && t.dbElements[0].Admission == "").ToList());
                List<ElementPki> modifyAdmissionObject = new List<ElementPki>();
                if (withOutAdmissionObjects.Count() > 0)
                {
                    var addAdmissiontForm = new AddAdmissionForm();

                    foreach (var obj in withOutAdmissionObjects)
                    {
                        var itmx = new ListViewItem[1];
                        itmx[0] = new ListViewItem(string.Join("", obj.ModifyPosDenotation()));

                        itmx[0].SubItems.Add(obj.Name);
                        itmx[0].SubItems.Add(obj.Count.ToString());
                        itmx[0].Tag = obj;

                        addAdmissiontForm.listViewObjectsReport.Items.AddRange(itmx);       
                    }

                    if (addAdmissiontForm.ShowDialog() == DialogResult.OK)
                    {
                        modifyAdmissionObject.AddRange(addAdmissiontForm.elementPKIs);
                        foreach (var obj in elements)
                        {
                            foreach (var pki in modifyAdmissionObject)
                            {
                                if (pki.PosDenotation == obj.PosDenotation)
                                {
                                    obj.Name = pki.Name;
                                }
                            }
                        }
                    }
                }
            }

            //Добавление наименования
            public void AddDescription(List<ElementPki> elements, Dictionary<string, ReferenceObject> normalRefObjects, bool otherItems)
            {
                //Объекты с * в поле Description БД Access 
                List<ElementPki> withOutDescriptionObjects = new List<ElementPki>();                
                withOutDescriptionObjects.AddRange(elements.Where(t => t.dbElements[0].Description.Contains("*")).ToList());

                //Формирование списка объектов с ошибкой в единице измерения
                var errorValueElements = elements.Where(t => t.ErrorValueText != string.Empty).Where(t => !withOutDescriptionObjects.Contains(t)).ToList();
                if (errorValueElements.Count >0)
                {
                    MessageBox.Show("Некорректный номинал или единица измерения у объектов: \r\n" + 
                        string.Join("\r\n", errorValueElements.Select(t=>t.PosDenotation + " " + t.Type + " " + t.Comment + " " + t.TU)) + 
               "\nДопустимые значения единиц измерения:\nмк\nк\nм\nг\nмкГ\nнГн\nма", "Ошибка!!!",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Clipboard.SetText(string.Join("\r\n", errorValueElements.Select(t => t.PosDenotation + " " + t.Type + " " + t.Comment + " " + t.TU)));
                }

                //Похожие объекты справочника ПКИ T-Flex Docs
                Dictionary<string, ReferenceObject> similarObjects = new Dictionary<string, ReferenceObject>();

                //Измененные объекты, записываемые в ПЭ
                List<ElementPki> modifyDescriptionObject = new List<ElementPki>();
                
                TextNormalizer textNorm = new TextNormalizer();
                if (withOutDescriptionObjects.Count() > 0)
                {
                    var addDescriptionForm = new AddDescriptionForm();
                    foreach (var obj in withOutDescriptionObjects)
                    {
                        int i = 0;
                        //Формирование списка похожих объектов справочника 
                        var objNorm = textNorm.GetNormalForm(obj.dbElements[0].Comment + " " +  obj.dbElements[0].Description.Replace("*", ""));

                        foreach (var simObj in normalRefObjects.Where(t => t.Key.Contains(objNorm) 
                            && t.Key.Contains(textNorm.GetNormalForm(obj.dbElements[0].TU))).ToList())
                        {
                            try
                            {
                                i++;
                                similarObjects.Add(simObj.Key, simObj.Value);                            
                            }
                            catch
                            {

                            }
                        }

                        if (i != 0 || !otherItems)
                        {
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(string.Join("", obj.ModifyPosDenotation()));
                           
                            itmx[0].SubItems.Add(obj.Name);
                            itmx[0].SubItems.Add(obj.NameDevice);
                            itmx[0].SubItems.Add(obj.DenotationDevice);
                            itmx[0].SubItems.Add(obj.Count.ToString());

                            itmx[0].Tag = obj;

                            addDescriptionForm.listViewObjectsReport.Items.AddRange(itmx);
                        }
                    }
                    if (addDescriptionForm.listViewObjectsReport.Items.Count == 0)
                    {
                        addDescriptionForm.Close();
                        return;
                    }
                    addDescriptionForm.normalRefObjects = similarObjects;
                    addDescriptionForm.otherItems = otherItems;
                    
                    if (addDescriptionForm.ShowDialog() == DialogResult.OK)
                    {
                        modifyDescriptionObject.AddRange(addDescriptionForm.elementPKIs);
                        foreach (var obj in elements)
                        {
                            foreach (var pki in modifyDescriptionObject)
                            {
                                if (pki.PosDenotation == obj.PosDenotation && pki.DenotationDevice == obj.DenotationDevice && pki.NameDevice == obj.NameDevice)
                                {
                                    obj.Name = pki.Name;
                                    if (!otherItems)
                                    {
                                        var elNormalForm = textNorm.GetNormalForm(pki.Name);
                                        try
                                        {
                                            var findedPKI = normalRefObjects.Where(t => t.Key == elNormalForm).ToList();
                                            if (findedPKI.Count > 1)
                                            {
                                                var analogPKI = findedPKI.Where(t => (t.Value[NomenclatureNameGuid].Value.ToString() +
                                                 t.Value[NomenclatureDenotationGuid].Value.ToString()).Trim().ToLower() ==
                                                (obj.Denotation + obj.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                                                if (analogPKI != null)
                                                    obj.NomObj = analogPKI;
                                                else
                                                    obj.NomObj = findedPKI[0].Value;
                                            }
                                            else
                                            {
                                                obj.NomObj = findedPKI[0].Value;
                                            }
                                           // obj.NomObj = normalRefObjects.Where(t => t.Key == elNormalForm).ToList()[0].Value;
                                        }
                                        catch
                                        {

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //public List<ElementPKI> FindObjectForSpecificationInNomenclature(List<ElementPKI> objForSpec, LogDelegate m_writeToLog)
            //{
            //    m_writeToLog("Обращение к Номенклатуре");
            //    // получение описания справочника          
            //    ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(NomenclatureReference);
            //    // создание объекта для работы с данными
            //    Reference referenceNomenclature = referenceNomenclatureInfo.CreateReference();


            //    ClassObject classObjectDetail = referenceNomenclatureInfo.Classes.Find(TypeDetailNomReference);
            //    ClassObject classObjectAssembly = referenceNomenclatureInfo.Classes.Find(TypeAssemblyNomReference);
            //    ClassObject classObjectOtherItems = referenceNomenclatureInfo.Classes.Find(TypeOtherItemsNomReference);
            //    ClassObject classStandartItems = referenceNomenclatureInfo.Classes.Find(TypeStandartItemsNomReference);
            //    ClassObject classMaterials = referenceNomenclatureInfo.Classes.Find(TypeMaterialNomReference);

            //    var findedData = new List<ElementPKI>();
            //    findedData.AddRange(SearchObjInNomenclature(objForSpec, m_writeToLog, classObjectDetail, referenceNomenclature, referenceNomenclatureInfo, false));
            //    findedData.AddRange(SearchObjInNomenclature(objForSpec, m_writeToLog, classObjectAssembly, referenceNomenclature, referenceNomenclatureInfo, false));
            //    findedData.AddRange(SearchObjInNomenclature(objForSpec, m_writeToLog, classObjectOtherItems, referenceNomenclature, referenceNomenclatureInfo, false));
            //    findedData.AddRange(SearchObjInNomenclature(objForSpec, m_writeToLog, classStandartItems, referenceNomenclature, referenceNomenclatureInfo, false));
            //    findedData.AddRange(SearchObjInNomenclature(objForSpec, m_writeToLog, classMaterials, referenceNomenclature, referenceNomenclatureInfo, false));

            //    var notFindedObj = objForSpec.Where(t => !findedData.Select(f => f.PosDenotation + f.Name).Contains(t.PosDenotation + t.Name));
            //   if (notFindedObj.Count() > 0)
            //    MessageBox.Show("Объекты для спецификации не найдены в Номенклатуре и не будут добавлены в результирующий файл str\r\n" + string.Join("\r\n", notFindedObj.Select(t => t.PosDenotation + " " + t.Name)),
            //        "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //    return findedData;
            //}


            public void AddTKS(List<ElementPki> elements, Dictionary<string, ReferenceObject> normalRefObjects, bool otherItems)
            {
                //Объекты с * в поле Description БД Access 
                List<ElementPki> withOutDescriptionObjects = new List<ElementPki>();
                withOutDescriptionObjects.AddRange(elements.Where(t => t.dbElements[0].OptionalParameter.Contains("*")).ToList());

                //Похожие объекты справочника ПКИ T-Flex Docs
                Dictionary<string, ReferenceObject> similarObjects = new Dictionary<string, ReferenceObject>();

                //Измененные объекты, записываемые в ПЭ
                List<ElementPki> modifyDescriptionObject = new List<ElementPki>();

                TextNormalizer textNorm = new TextNormalizer();
                if (withOutDescriptionObjects.Count() > 0)
                {
                    var addDescriptionForm = new AddDescriptionForm();
                    addDescriptionForm.selectTKS = true;

                    foreach (var obj in withOutDescriptionObjects)
                    {
                       
                        int i = 0;
                        //Формирование списка похожих объектов справочника 
                        var objNorm = textNorm.GetNormalForm(obj.Name.Replace("*", "").Replace(obj.dbElements[0].TU,""));                      


                        foreach (var simObj in normalRefObjects.Where(t => t.Key.Contains(objNorm)
                            && t.Key.Contains(textNorm.GetNormalForm(obj.dbElements[0].TU))).ToList())
                        {
                            try
                            {
                                i++;
                                similarObjects.Add(simObj.Key, simObj.Value);
                            }
                            catch
                            {

                            }
                        }

                        if (i != 0 || !otherItems)
                        {
                            var itmx = new ListViewItem[1];
                            itmx[0] = new ListViewItem(string.Join("", obj.ModifyPosDenotation()));

                            itmx[0].SubItems.Add(obj.Name);
                            itmx[0].SubItems.Add(obj.Count.ToString());
                            itmx[0].Tag = obj;

                            addDescriptionForm.listViewObjectsReport.Items.AddRange(itmx);
                        }
                    }
                    if (addDescriptionForm.listViewObjectsReport.Items.Count == 0)
                    {
                        addDescriptionForm.Close();
                        return;
                    }
                    addDescriptionForm.normalRefObjects = similarObjects;
                    addDescriptionForm.otherItems = otherItems;

                    if (addDescriptionForm.ShowDialog() == DialogResult.OK)
                    {
                        modifyDescriptionObject.AddRange(addDescriptionForm.elementPKIs);
                        foreach (var obj in elements)
                        {
                            foreach (var pki in modifyDescriptionObject)
                            {
                                if (pki.PosDenotation == obj.PosDenotation)
                                {
                                    obj.Name = pki.Name;
                                    if (!otherItems)
                                    {
                                        var elNormalForm = textNorm.GetNormalForm(pki.Name);
                                        try
                                        {
                                            var findedPKI = normalRefObjects.Where(t => t.Key == elNormalForm).ToList();
                                            if (findedPKI.Count > 1)
                                            {
                                                var analogPKI = findedPKI.Where(t => (t.Value[NomenclatureNameGuid].Value.ToString() +
                                                 t.Value[NomenclatureDenotationGuid].Value.ToString()).Trim().ToLower() ==
                                                (obj.Denotation + obj.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                                                if (analogPKI != null)
                                                    obj.NomObj = analogPKI;
                                                else
                                                    obj.NomObj = findedPKI[0].Value;
                                            }
                                            else
                                            {
                                                obj.NomObj = findedPKI[0].Value;
                                            }
                                          //  obj.NomObj = normalRefObjects.Where(t => t.Key == elNormalForm).ToList()[0].Value;
                                        }
                                        catch
                                        {

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }


            public List<ElementPki> SearchObjectInSpecReference(List<ElementPki> elements, LogDelegate m_writeToLog, Guid specNomReferenceGuid, Guid nameGuid, Guid denotationGuid, Guid typeReferenceGuid)
            {
                var findedObject = new List<ElementPki>();
                List<ElementPki> data = new List<ElementPki>();
                // получение описания справочника          
                ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(specNomReferenceGuid);
                // создание объекта для работы с данными
                Reference reference = referenceInfo.CreateReference();
                reference.LoadSettings.AddParameters(nameGuid, denotationGuid);
                ClassObject classObject = referenceInfo.Classes.Find(typeReferenceGuid);
                //Создаем фильтр
                Filter filter = new Filter(referenceInfo);
                //Добавляем условие поиска – «Тип = Длина»
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                    ComparisonOperator.Equal, classObject);

                m_writeToLog("Загрузка объектов справочника " + reference);
                //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
                List<ReferenceObject> listObj = reference.Find(filter);

                // reference.LoadLinks(listObj, reference.LoadSettings);

                var lStr = new List<string>();

                TextNormalizer textNormal = new TextNormalizer();

                m_writeToLog("Считано " + listObj.Count + " объектов");
                //  listObj = listObj.Where(t => t[CommentItemsGuid].GetString().ToLower().Trim() != "не использовать").ToList();
                m_writeToLog("Преобразование к нормальной форме " + listObj.Count + " объектов");
                Dictionary<string, ReferenceObject> normalRefObjects = new Dictionary<string, ReferenceObject>();

                foreach (var obj in listObj)
                {
                    var textNormName = textNormal.GetNormalForm(obj[nameGuid].Value.ToString() + " " + obj[denotationGuid].Value.ToString());

                    try
                    {
                        normalRefObjects.Add(textNormName, obj);
                    }
                    catch
                    {

                    }
                }

                foreach (var el in elements)
                {
                    var findEl = normalRefObjects.Where(t => t.Key == textNormal.GetNormalForm(el.Name + " " + el.Denotation)).ToDictionary(t => t.Key, t => t.Value).Values.FirstOrDefault();
                    if (findEl != null)
                    {
                        ElementPki pki = new ElementPki();
                        el.Name = findEl[nameGuid].Value.ToString();
                        el.Denotation = findEl[denotationGuid].Value.ToString();
                        findedObject.Add(el);
                    }
                }

                return findedObject;
            }


            public List<ElementPki> selRezConds = new List<ElementPki>();
            //  private List<ElementPKI> detailsNomenclature = new List<ElementPKI>();

            private int GetDigitSort(string posDenotation)
            {
                var digitForSort = 0;
                var regexDigital = new Regex(@"\d{1,}");
                var matchDigit = regexDigital.Match(posDenotation.Trim());
                return matchDigit.Success ? Convert.ToInt32(matchDigit.Value) : digitForSort;
            }

            //Проверка элемента на существование в справочнике "Прочие изделия"
            public List<ElementPki> ProcessingData(ReportParameters reportParameters, List<ElementPki> elements, LogDelegate m_writeToLog, bool writeAll, out List<ElementPki> detailsNomenclature)
            {
                detailsNomenclature = new List<ElementPki>();
                List<ElementPki> data = new List<ElementPki>();
                List<ElementPki> resultSortData = new List<ElementPki>();
                // получение описания справочника          
                ReferenceInfo referenceInfo = ReferenceCatalog.FindReference(OtherItemsReference);
                // создание объекта для работы с данными
                Reference reference = referenceInfo.CreateReference();
                reference.LoadSettings.AddParameters(NameOtherItemsGuid, DenotationItemsGuid);
                ClassObject classObject = referenceInfo.Classes.Find(TypeOtherItems);
                //Создаем фильтр
                Filter filter = new Filter(referenceInfo);
                //Добавляем условие поиска – «Тип = Длина»
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                    ComparisonOperator.Equal, classObject);
                m_writeToLog("Загрузка объектов справочника прочих изделий");
                //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
                List<ReferenceObject> listObj = reference.Find(filter);

                // reference.LoadLinks(listObj, reference.LoadSettings);

                TextNormalizer textNormal = new TextNormalizer();

                m_writeToLog("Считано " + listObj.Count + " объектов");
                //  listObj = listObj.Where(t => t[CommentItemsGuid].GetString().ToLower().Trim() != "не использовать").ToList();
                m_writeToLog("Преобразование к нормальной форме " + listObj.Count + " объектов");
                Dictionary<string, ReferenceObject> normalRefObjects = new Dictionary<string, ReferenceObject>();
                foreach (var obj in listObj)
                {
                    var textNormName = textNormal.GetNormalForm(obj[NameOtherItemsGuid].Value.ToString() + " " + obj[DenotationItemsGuid].Value.ToString());

                    try
                    {
                        normalRefObjects.Add(textNormName, obj);
                    }
                    catch
                    {

                    }
                }


                if (!writeAll)
                {
                    AddDescription(elements, normalRefObjects, true);

                    //Если выбран флажок Выбор ТКС вызываем процедуру и форму 
                    if (reportParameters.SelectTKS)
                    {
                        AddTKS(elements, normalRefObjects, true);
                    }
                }

                List<ElementPki> notFoundElement = new List<ElementPki>();
                m_writeToLog("Поиск объектов");
                foreach (var el in elements)
                {
                    var elNormalForm = textNormal.GetNormalForm(el.Name);

                    var findedPKI = normalRefObjects.Where(t => t.Key == elNormalForm).ToList();
                    if (findedPKI.Count > 0)
                    {
                        if (findedPKI.Count > 1)
                        {
                            var analogPKI = findedPKI.Where(t => (t.Value[NomenclatureNameGuid].Value.ToString() +
                             t.Value[NomenclatureDenotationGuid].Value.ToString()).Trim().ToLower() ==
                            (el.Denotation + el.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                            if (analogPKI != null)
                                el.NomObj = analogPKI;
                            else
                                el.NomObj = findedPKI[0].Value;
                        }
                        else
                        {
                            el.NomObj = findedPKI[0].Value;
                        }

                        el.Type = "Прочее изделие";
                        data.Add(el);

                    }
                    else
                    {
                        if (el.RezWithOutNominal || writeAll)
                        {
                            data.Add(el);
                            continue;
                        }                   
                        else
                        {
                            if (reportParameters.IsKonstr && el.Type.ToLower().Contains("konstr"))
                            {
                                data.Add(el);
                            }
                            else
                            {
                                notFoundElement.Add(el);
                            }
                        }
                    }
                }
                var resultData = new List<ElementPki>();
                var resultDataWithAssembly = new List<ElementPki>();
                var newResultData = new List<ElementPki>();
                var prewObj = new ElementPki();
                var selectionObjects = new List<ElementPki>();

                resultData.AddRange(data);

                //foreach (var obj in data)
                //{

                //    try
                //    {

                //        if (prewObj.Name != obj.Name || (obj.dbElements.Count > 0
                //            && obj.dbElements[0].PartNumber.Trim().Last() == 'P' && (obj.Name.Contains("Конденсатор") || obj.Name.Contains("Резистор"))))
                //        {
                //            resultData.Add(obj);
                //        }
                //        else
                //        {
                //            if (resultData.Count > 0)
                //            {
                //                resultData.Last().PosDenotation += "," + obj.PosDenotation;
                //                resultData.Last().Count += obj.Count;
                //            }
                //        }
                //    }
                //    catch
                //    {

                //    }
                //    prewObj = new ElementPKI();
                //    prewObj = obj;

                //}

                foreach (var obj in resultData)
                {
                    try
                    {
                        if (obj.dbElements.Count > 0 && obj.dbElements[0].PartNumber.Trim().Last() == 'P' && (obj.Name.Contains("Конденсатор") || obj.Name.Contains("Резистор")))
                        {
                            obj.SelectionObject = true;
                            selectionObjects.Add(obj);
                        }
                    }
                    catch
                    {

                    }
                }
                if (selectionObjects.Count > 0)
                {
                    var addCommentForm = new AddCommentForm();
                    foreach (var obj in selectionObjects)
                    {
                        var itmx = new ListViewItem[1];
                        itmx[0] = new ListViewItem(string.Join("", obj.ModifyPosDenotation()));

                        itmx[0].SubItems.Add(obj.Name);
                        itmx[0].SubItems.Add(obj.Count.ToString());
                        itmx[0].SubItems.Add("");
                        itmx[0].Tag = obj;

                        addCommentForm.listViewObjectsReport.Items.AddRange(itmx);

                    }

                    if (addCommentForm.ShowDialog() == DialogResult.OK)
                    {
                        selRezConds.AddRange(addCommentForm.elementPKIs);
                        foreach (var obj in resultData)
                        {
                            foreach (var pki in selRezConds)
                            {
                                if (pki.Name == obj.Name && pki.PosDenotation == obj.PosDenotation && pki.Comment.Trim() != string.Empty)
                                {
                                    obj.Comment = pki.Comment;
                                }
                            }
                        }
                    }
                }
                m_writeToLog("Найдено " + resultData.Count + " объектов");
                if (notFoundElement.Count > 0)
                {
                    //MessageBox.Show(notFoundElement.Count + "");
                    //StringBuilder strB = new StringBuilder();

                    //foreach (var el in notFoundElement)
                    //{
                    //    strB.AppendLine(el.PosDenotation + " " +el.Name);
                    //}
                    //MessageBox.Show(strB.ToString());
                    //Clipboard.SetText(strB.ToString());

                    m_writeToLog("Обращение к Номенклатуре");
                    // получение описания справочника          
                    ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(NomenclatureReference);
                    // создание объекта для работы с данными
                    Reference referenceNomenclature = referenceNomenclatureInfo.CreateReference();


                    ClassObject classObjectDetail = referenceNomenclatureInfo.Classes.Find(TypeDetailNomReference);
                    ClassObject classObjectAssembly = referenceNomenclatureInfo.Classes.Find(TypeAssemblyNomReference);
                    //  ClassObject classObjectOtherItems = referenceNomenclatureInfo.Classes.Find(TypeOtherItemsNomReference);

                    var foundedDetails = new List<ElementPki>();
                    foundedDetails.AddRange(SearchObjInNomenclature(notFoundElement, m_writeToLog, classObjectDetail, referenceNomenclature, referenceNomenclatureInfo, false));
                    foreach (var el in foundedDetails)
                    {
                        notFoundElement.Remove(el);
                        el.Type = "Деталь";
                        detailsNomenclature.Add(el);
                    }
                    var foundAssembly = SearchObjInNomenclature(notFoundElement, m_writeToLog, classObjectAssembly, referenceNomenclature, referenceNomenclatureInfo, true).OrderBy(t => t.Name).ToList();

                    //  var foundedOtherItems = new List<ElementPKI>();

                    //if (reportParameters.SelectTKS)
                    //{
                    //   // MessageBox.Show("не надйенные: " + string.Join(";",notFoundElement.Select(t=>t.Name)));
                    //    foundedOtherItems = SearchObjInNomenclature(notFoundElement, m_writeToLog, classObjectOtherItems, referenceNomenclature, referenceNomenclatureInfo, true).OrderBy(t => t.Name).ToList();
                    //}
                    string prewName = string.Empty;
                    ElementPki prewEl = new ElementPki();
                    string positions = string.Empty;
                    bool write = false;
                    int count = 0;
                    int prewDigit = 0;
                    Regex regexLetter = new Regex(@"\D");
                    var assemblySort = new List<ElementPki>();
                    assemblySort.AddRange(foundAssembly.OrderBy(t => regexLetter.Match(t.PosDenotation.Trim()).Value).ThenBy(t => GetDigitSort(t.PosDenotation)).ToList());
                    for (int i = 0; i < assemblySort.Count; i++)
                    {

                        assemblySort[i].Type = "Сборочная единица";
                        notFoundElement.Remove(assemblySort[i]);

                        resultDataWithAssembly.Add(assemblySort[i]);

                        //if (assemblySort[i] != assemblySort.Last() && assemblySort[i].Name != assemblySort[i + 1].Name)
                        //{
                        //    resultDataWithAssembly.Add(assemblySort[i]);
                        //    continue;
                        //}

                        //if (assemblySort[i] != assemblySort.Last() && assemblySort[i].Name == assemblySort[i + 1].Name)
                        //{
                        //    assemblySort[i + 1].Count = assemblySort[i + 1].Count + assemblySort[i].Count;
                        //    assemblySort[i + 1].PosDenotation = assemblySort[i].PosDenotation + "," + assemblySort[i + 1].PosDenotation;
                        //    continue;
                        //}

                        //if (assemblySort[i] == assemblySort.Last())
                        //{
                        //    resultDataWithAssembly.Add(assemblySort[i]);
                        //}
                    }

                    m_writeToLog("Найдено " + assemblySort.Count + " объектов");
                    List<string> notFoundStr = new List<string>();

                    foreach (var obj in notFoundElement)
                    {
                       // MessageBox.Show(obj.PosDenotation + " " + obj.Comment + "???!!!");
                        if (reportParameters.IsKonstr)
                        {
                            if (!obj.Type.ToLower().Contains("konstr"))
                            {
                                notFoundStr.Add(obj.PosDenotation + " " + obj.Name);
                            }
                        }
                        else
                        {
                            notFoundStr.Add(obj.PosDenotation + " " + obj.Name);
                        }
                    }
                    if (notFoundStr.Count() > 0)
                    {
                        notFoundStr = notFoundStr.OrderBy(t => t).ToList();
                        Clipboard.Clear();
                        Clipboard.SetText(string.Join("\r\n", notFoundStr));
                        //MessageBox.Show("Не найдены объекты, в количестве " + notFoundStr.Count + ":\n" + string.Join("\r\n", notFoundStr) + "\nОни не будут записаны в Перечень элементов.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        globalLog.AppendLine("Не найдены объекты, в количестве " + notFoundStr.Count + ":\n" + string.Join("\r\n", notFoundStr) + "\nОни не будут записаны в Перечень элементов.");
                        Clipboard.SetText(globalLog.ToString());
                    }
                }

                //resultDataWithAssembly.AddRange(resultData);
                string patternPosDigit = @"\d{1,}";
                Regex regexPosDigit = new Regex(patternPosDigit);
                resultSortData.AddRange(resultDataWithAssembly);
                resultSortData.AddRange(resultData);

                var notKorObj = new List<ElementPki>();

                foreach (var obj in resultSortData)
                {
                    try
                    {
                        var res = (Convert.ToInt32(regexPosDigit.Match(obj.PosDenotation).Value));
                    }
                    catch
                    {
                        notKorObj.Add(obj);
                    }

                }
                if (notKorObj.Count > 0)
                {
                    foreach (var obj in notKorObj)
                        resultSortData.Remove(obj);

                    MessageBox.Show("Неправильное позиционное обозначение у объектов:\r\n" + string.Join("\r\n", notKorObj.Select(t => t.PosDenotation + "  " + t.Name + " " + t.Denotation)), "Внимание!!! Некорректный объект.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

              //  resultSortData = resultSortData.Where(t => t.PosDenotation != string.Empty).ToList().OrderBy(t => t.LetterPosDenotation).ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value)).ToList();
                //Суммирование резисторов, в том случае, если они располагаются рядом
                //Суммирование осуществляется в пределах устройства, если выбран флаг СВЧ
                var groupsData = resultSortData.GroupBy(t => t.Name + t.Denotation + t.dbElements[0].PartNumber.Trim().Last() + (reportParameters.IsManyDevice ? t.DenotationDevice + t.NameDevice : string.Empty))
                    .ToDictionary(t => t.Key, t => t.ToList());

                foreach (var group in groupsData)
                {
                    var groupsByNumber = group.Value.GroupBy(t => GetNumberGroup(t.PosDenotation, group.Value.Select(s => s.PosDenotation).ToList())).ToDictionary(t => t.Key, t => t.ToList());
                    foreach (var groupByNumber in groupsByNumber)
                    {
                        groupByNumber.Value.First().PosDenotation = string.Join(",", groupByNumber.Value.Select(t => t.PosDenotation));
                        groupByNumber.Value.First().Count = groupByNumber.Value.Sum(t => t.Count);
                        newResultData.Add(groupByNumber.Value.First());
                    }
                }
                resultSortData = newResultData.Where(t => t.PosDenotation != string.Empty).ToList().OrderBy(t => t.LetterPosDenotation).ThenBy(t => Convert.ToInt32(regexPosDigit.Match(t.PosDenotation).Value)).ToList();

                return resultSortData;
            }

            private int GetNumberGroup(string posDenotaion, List<string> posDenotations)
            {
                //resultDataWithAssembly.AddRange(resultData);
                var digitsOrder = posDenotations.OrderBy(x => GetDigitSort(x)).ToList();
                var groups = new List<List<string>>();
                var currentGroup = new List<string> { digitsOrder[0] };

                for (int i = 1; i < digitsOrder.Count; i++)
                {
                    if (GetDigitSort(digitsOrder[i]) - GetDigitSort(digitsOrder[i - 1]) <= 1)
                    {
                        currentGroup.Add(digitsOrder[i]);
                    }
                    else
                    {
                        groups.Add(currentGroup);
                        currentGroup = new List<string> { digitsOrder[i] };
                    }
                }
                groups.Add(currentGroup);
                return groups.IndexOf(groups.First(t => t.Contains(posDenotaion)));
            }

            public List<ElementPki> SearchObjInNomenclature(List<ElementPki> elements, LogDelegate m_writeToLog, ClassObject classObject, Reference reference, ReferenceInfo referenceInfo, bool assembly)
            {
                List<ElementPki> data = new List<ElementPki>();             
              
                Filter filter = new Filter(referenceInfo);
                filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                    ComparisonOperator.Equal, classObject);

                m_writeToLog("Загрузка объектов справочника " + reference + " для типа объектов " + classObject);
                //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
                List<ReferenceObject> listObj = reference.Find(filter);
                reference.LoadSettings.AddParameters(NomenclatureNameGuid, NomenclatureDenotationGuid);
               // reference.LoadLinks(listObj);

                var lStr = new List<string>();

                TextNormalizer textNormal = new TextNormalizer();
                Dictionary<string, ReferenceObject> normalNomRefObjects = new Dictionary<string, ReferenceObject>();
                m_writeToLog("Преобразование к нормальной форме " + listObj.Count + " объектов" );
                foreach (var obj in listObj)
                {
                    try
                    {
                        //if (obj[ParametersInfo.Denotation].Value.ToString() == "АЕЯР.431420.668ТУ")
                        //{
                        //    MessageBox.Show(elements[0].Name + "\n" + obj[ParametersInfo.Name].Value.ToString() + " " + obj[ParametersInfo.Denotation].Value.ToString());
                        //}
                       // normalNomRefObjects.Add(textNormal.GetNormalForm(obj[ParametersInfo.Name].Value.ToString() +" "+ obj[ParametersInfo.Denotation].Value.ToString()), obj);
                        normalNomRefObjects.Add(textNormal.GetNormalForm(obj[NomenclatureNameGuid].Value.ToString() + " " + obj[NomenclatureDenotationGuid].Value.ToString()), obj);
                    }
                    catch
                    {

                    }
                }

                //if (classObject.ToString() == "Прочее изделие")
                //{
                //    MessageBox.Show("Объектов для сравнения: " + listObj.Count );
                //}

                if (assembly)
                {
                    AddDescription(elements, normalNomRefObjects, false);
                }
                //if (classObject.ToString() == "Прочее изделие")
                //{
                //    MessageBox.Show(" " + elements.Count);
                //}

                m_writeToLog("Поиск объектов");
                foreach (var el in elements)
                {
                    var elNormalForm = textNormal.GetNormalForm(el.Name);

                    var findedPKI = normalNomRefObjects.Where(t => t.Key == elNormalForm).ToList();
                    if (findedPKI.Count > 0)
                    {
                        if (findedPKI.Count > 1)
                        {

                            var analogPKI = findedPKI.Where(t => (t.Value[NomenclatureNameGuid].Value.ToString() +
                             t.Value[NomenclatureDenotationGuid].Value.ToString()).Trim().ToLower() ==
                            (el.Denotation + el.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                            if (analogPKI != null)
                            {
                                el.Name = analogPKI[NomenclatureNameGuid].Value.ToString();
                                el.Denotation = analogPKI[NomenclatureDenotationGuid].Value.ToString();
                                el.NomObj = analogPKI;
                                data.Add(el);
                                //  obj.NomObj = analogPKI;
                            }
                            else
                            {
                                el.Name = findedPKI[0].Value[NomenclatureNameGuid].Value.ToString();
                                el.Denotation = findedPKI[0].Value[NomenclatureDenotationGuid].Value.ToString();
                                el.NomObj = findedPKI[0].Value;
                                data.Add(el);
                                // obj.NomObj = findedPKI[0].Value;
                            }

                        }
                        else
                        {
                            el.Name = findedPKI[0].Value[NomenclatureNameGuid].Value.ToString();
                            el.Denotation = findedPKI[0].Value[NomenclatureDenotationGuid].Value.ToString();
                            el.NomObj = findedPKI[0].Value;
                            data.Add(el);
                        }
                    }

                    
                }

                return data;
            }

            public List<ElementPki> SearchLinkedObjInNomenclature(List<ElementPki> elements, List<ElementPki> detailsNomenclature, LogDelegate m_writeToLog, List<ElementPki> objectsForSpecification)
            {
                if (elements.Count(t => t.dbElements.Count > 0 && t.dbElements[0].LinkedObject.Count > 0) == 0)
                {
                    return null;
                }

                ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(NomenclatureReference);
                List<ElementPki> pkiLinkedObjects = new List<ElementPki>();
                List<ElementPki> distinctLinkedObjects = new List<ElementPki>();

                // List<ElementPKI> notFindedLinkedObjects = new List<ElementPKI>();
                List<ElementPki> notFindedOjForSpec = new List<ElementPki>();

                // создание объекта для работы с данными
                Reference referenceNomenclature = referenceNomenclatureInfo.CreateReference();
                referenceNomenclature.LoadSettings.AddParameters(NomenclatureNameGuid, NomenclatureDenotationGuid);

                ClassObject classObjectDetail = referenceNomenclatureInfo.Classes.Find(TypeDetailNomReference);
                ClassObject classObjectAssembly = referenceNomenclatureInfo.Classes.Find(TypeAssemblyNomReference);
                ClassObject classObjectStandartItems = referenceNomenclatureInfo.Classes.Find(TypeStandartItemsNomReference);
                ClassObject classObjectMaterials = referenceNomenclatureInfo.Classes.Find(TypeMaterialNomReference);
                List<ClassObject> nomenclatureClasses = new List<ClassObject>();
                nomenclatureClasses.Add(classObjectAssembly);
                nomenclatureClasses.Add(classObjectDetail);
                nomenclatureClasses.Add(classObjectStandartItems);
                nomenclatureClasses.Add(classObjectMaterials);

                Filter filter = new Filter(referenceNomenclatureInfo);
                filter.Terms.AddTerm(referenceNomenclature.ParameterGroup[SystemParameterType.Class],
                    ComparisonOperator.IsOneOf, nomenclatureClasses);

                m_writeToLog("Загрузка объектов справочника Номенклатура для присоединенных объектов");

                //Получаем список объектов, в качестве условия поиска – сформированный фильтр            
                List<ReferenceObject> listObj = referenceNomenclature.Find(filter);

                var lStr = new List<string>();

                TextNormalizer textNormal = new TextNormalizer();
                Dictionary<string, ReferenceObject> normalNomRefObjects = new Dictionary<string, ReferenceObject>();
                m_writeToLog("Преобразование к нормальной форме " + listObj.Count + " объектов");
                foreach (var obj in listObj)
                {
                    try
                    {
                        // normalNomRefObjects.Add(textNormal.GetNormalForm(obj[ParametersInfo.Name].Value.ToString() + " " + obj[ParametersInfo.Denotation].Value.ToString()), obj);
                        normalNomRefObjects.Add(textNormal.GetNormalForm(obj[NomenclatureNameGuid].Value.ToString() + " " + obj[NomenclatureDenotationGuid].Value.ToString()), obj);
                    }
                    catch
                    {

                    }
                }

                m_writeToLog("Поиск объектов");

                //foreach (ElementPKI el in elements)
                //{
                //    if (el.dbElements.Count > 0)
                //    {
                //        foreach (var linkedObj in el.dbElements[0].LinkedObject)
                //        {
                //            var elNormalForm = textNormal.GetNormalForm(linkedObj.Name + " " + linkedObj.Denotation);

                //            var findedPKI = normalNomRefObjects.Where(t => t.Key == elNormalForm).ToList();
                //            if (findedPKI.Count > 0)
                //            {

                //                ElementPKI pki = new ElementPKI();

                //                if (findedPKI.Count > 1)
                //                {

                //                    var analogPKI = findedPKI.Where(t => (t.Value[ParametersInfo.Name].Value.ToString() +
                //                     t.Value[ParametersInfo.Denotation].Value.ToString()).Trim().ToLower() ==
                //                    (linkedObj.Denotation + linkedObj.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                //                    if (analogPKI != null)
                //                    {
                //                        pki.Name = analogPKI[ParametersInfo.Name].Value.ToString();
                //                        pki.Denotation = analogPKI[ParametersInfo.Denotation].Value.ToString();
                //                        pki.Type = analogPKI.Class.Name;
                //                        pki.NomObj = analogPKI;
                //                        pki.Parent = el.Name + " " + el.Denotation;
                //                    }
                //                    else
                //                    {
                //                        pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                //                        pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                //                        pki.Type = findedPKI[0].Value.Class.Name;
                //                        pki.NomObj = findedPKI[0].Value;
                //                        pki.Parent = el.Name + " " + el.Denotation;
                //                    }

                //                }
                //                else
                //                {
                //                    pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                //                    pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                //                    pki.Type = findedPKI[0].Value.Class.Name;
                //                    pki.NomObj = findedPKI[0].Value;
                //                    pki.Parent = el.Name + " " + el.Denotation;
                //                }


                //                pki.PosDenotation = " ";
                //                pki.Comment = linkedObj.Comment;
                //                pki.Preparation = linkedObj.Preparation;
                //                pki.Count = el.Count * linkedObj.Count;
                //                pki.Parent = el.Name + " " + el.Denotation;

                //                pkiLinkedObjects.Add(pki);
                //            }
                //            else
                //            {
                //                var notFoundPKI = new ElementPKI();
                //                notFoundPKI.Name = linkedObj.Name;
                //                notFoundPKI.Denotation = linkedObj.Denotation;
                //                notFoundPKI.Type = linkedObj.Type;

                //                notFoundPKI.PosDenotation = " ";
                //                notFoundPKI.Comment = linkedObj.Comment;
                //                notFoundPKI.Preparation = linkedObj.Preparation;
                //                notFoundPKI.Count = el.Count * linkedObj.Count;

                //                if (!notFindedOjForSpec.Contains(notFoundPKI))
                //                    notFindedOjForSpec.Add(notFoundPKI);
                //            }
                //        }
                //    }
                //}
                pkiLinkedObjects.AddRange(GetPKILinkedObjects(detailsNomenclature, notFindedOjForSpec, textNormal, normalNomRefObjects));
               
                pkiLinkedObjects.AddRange(GetPKILinkedObjects(elements, notFindedOjForSpec, textNormal, normalNomRefObjects));

                //Добавление связанных объектов для спецификации для поиска в номенклатуре
                pkiLinkedObjects.AddRange(GetPKILinkedObjects(objectsForSpecification, notFindedOjForSpec, textNormal, normalNomRefObjects));

                //Добавление объектов для спецификации для поиска в номенклатуре
                foreach (ElementPki el in objectsForSpecification)
                {
                    var elNormalForm = textNormal.GetNormalForm(el.Name + " " + el.Denotation);
                    // MessageBox.Show(linkedObj.Name + linkedObj.Denotation);

                    var findedPKI = normalNomRefObjects.Where(t => t.Key == elNormalForm).ToList();
                    if (findedPKI.Count > 0)
                    {
                        ElementPki pki = new ElementPki();

                        if (findedPKI.Count > 1)
                        {
                            var analogPKI = findedPKI.Where(t => (t.Value[ParametersInfo.Name].Value.ToString() +
                             t.Value[ParametersInfo.Denotation].Value.ToString()).Trim().ToLower() ==
                            (el.Denotation + el.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                            
                            if (analogPKI != null)
                            {
                                pki.Name = analogPKI[ParametersInfo.Name].Value.ToString();
                                pki.Denotation = analogPKI[ParametersInfo.Denotation].Value.ToString();
                                pki.Type = analogPKI.Class.Name;
                                pki.NomObj = analogPKI;
                            }
                            else
                            {
                                pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                                pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                                pki.Type = findedPKI[0].Value.Class.Name;
                                pki.NomObj = findedPKI[0].Value;
                            }
                        }
                        else
                        {
                            pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                            pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                            pki.Type = findedPKI[0].Value.Class.Name;
                            pki.NomObj = findedPKI[0].Value;
                        }

                        pki.PosDenotation = " ";
                        pki.Comment = el.Comment;
                        pki.Preparation = el.Preparation;
                        pki.Count = el.Count;
                        pki.AddedNotLinkedObj = true;
                        pkiLinkedObjects.Add(pki);
                    }
                    else
                    {
                        if (!notFindedOjForSpec.Contains(el))
                            notFindedOjForSpec.Add(el);
                    }
                }

                //Поиск объектов для спецификации в справочнике Стандартные изделия
                var findedStandartItems = SearchObjectInSpecReference(notFindedOjForSpec, m_writeToLog,
                    Guids.SpecReferences.StandartItems,
                    Guids.StandartItemsParameters.Name,
                    Guids.StandartItemsParameters.Denotation,
                    Guids.SpecReferencesTypes.StandartItems);

                if (findedStandartItems.Count > 0)
                {
                    foreach (var el in findedStandartItems)
                    {
                        notFindedOjForSpec.Remove(el);
                        pkiLinkedObjects.Add(el);
                    }
                }

                //Поиск объектов для спецификации в справочнике Прочие изделия
                var findedOtherItems = SearchObjectInSpecReference(notFindedOjForSpec, m_writeToLog,
                  Guids.SpecReferences.OtherItems,
                  Guids.OtherItemsParameters.Name,
                  Guids.OtherItemsParameters.Denotation,
                  Guids.SpecReferencesTypes.OtherItems);

                if (findedOtherItems.Count > 0)
                {
                    foreach (var el in findedOtherItems)
                    {
                        notFindedOjForSpec.Remove(el);
                        pkiLinkedObjects.Add(el);
                    }
                }
                //Поиск объектов для спецификации в справочнике Материалы
                var findedMaterials = SearchObjectInSpecReference(notFindedOjForSpec, m_writeToLog,
                  Guids.SpecReferences.Materials,
                  Guids.MaterialsParameters.Name,
                  Guids.MaterialsParameters.Denotation,
                  Guids.SpecReferencesTypes.Materials);
                if (findedMaterials.Count > 0)
                {
                    foreach (var el in findedMaterials)
                    {
                        notFindedOjForSpec.Remove(el);
                        pkiLinkedObjects.Add(el);
                    }
                }


                if (detailsNomenclature.Count > 0)
                {
                    ////Удаление объектов из списка PkiLinkedObject, если такие уже есть в списке detailsNomenclature
                    //foreach (var el in detailsNomenclature)
                    //{
                    //    var dublicateObj = pkiLinkedObjects.Where(t => t.NomObj != null && el.NomObj != null && t.NomObj.SystemFields.Id == el.NomObj.SystemFields.Id && t.Parent == el.Name + " " + el.Denotation).FirstOrDefault();

                    //    if (dublicateObj != null)
                    //    {
                    //        pkiLinkedObjects.Remove(dublicateObj);
                    //    }
                    //}

                    Regex regexDigitals = new Regex(@"\d{1,}");
                    Regex regexLetter = new Regex(@"\D");
                    List<ElementPki> resultDetails = new List<ElementPki>();
                    try
                    {
                        var detailsSort = detailsNomenclature.OrderBy(t => regexLetter.Match(t.PosDenotation.Trim()).Value).ThenBy(t => Convert.ToInt32(regexDigitals.Match(t.PosDenotation.Trim()).Value)).ToList();

                        for (int i = 0; i < detailsSort.Count; i++)
                        {
                            if (detailsSort[i] != detailsSort.Last() && detailsSort[i].Name != detailsSort[i + 1].Name)
                            {
                                resultDetails.Add(detailsSort[i]);
                                continue;
                            }

                            if (detailsSort[i] != detailsSort.Last() && detailsSort[i].Name == detailsSort[i + 1].Name)
                            {
                                detailsSort[i + 1].Count = detailsSort[i + 1].Count + detailsSort[i].Count;
                                detailsSort[i + 1].PosDenotation = detailsSort[i].PosDenotation + "," + detailsSort[i + 1].PosDenotation;
                                continue;
                            }

                            if (detailsSort[i] == detailsSort.Last())
                            {
                                resultDetails.Add(detailsSort[i]);
                            }
                        }
                    }
                    catch
                    {
                        Clipboard.SetText(string.Join("\r\n", detailsNomenclature.Select(t => t.Name + " " + t.PosDenotation)));
                        MessageBox.Show("Некорректные позиции у объектов \r\n" + string.Join("\r\n", detailsNomenclature.Select(t => t.Name + " " + t.PosDenotation)) + "\r\nОни не будут добавлены в отчет",
                            "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    distinctLinkedObjects.AddRange(resultDetails);
                }



                pkiLinkedObjects = pkiLinkedObjects.OrderBy(t => t.Name).ThenBy(t => t.Denotation).ThenBy(t => t.SizeMaterial).ToList();
                string prewName = string.Empty;
                string prewType = string.Empty;
                string prewComment = string.Empty;


                foreach (var pki in pkiLinkedObjects)
                {
                    if (prewName == pki.Name + pki.Denotation && prewType == pki.Type && prewComment == pki.Comment)
                    {
                        distinctLinkedObjects.Last().Count += pki.Count;
                    }
                    else
                    {
                        distinctLinkedObjects.Add(pki);
                    }
                    prewName = pki.Name + pki.Denotation;
                    prewType = pki.Type;
                    prewComment = pki.Comment;
                }

                m_writeToLog("Найдено " + distinctLinkedObjects.Count + " объектов");


                if (notFindedOjForSpec.Count > 0)
                {
                    var notFindedObjString = string.Join("\r\n", notFindedOjForSpec.Select(t => t.PosDenotation + " " + t.Name + " " + t.Denotation));

                    Clipboard.SetText(notFindedObjString);
                    MessageBox.Show("Объекты ненайденные в справочнике Номенклатура и изделия для формирования спецификации:\r\n" + notFindedObjString, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //if (notFindedLinkedObjects.Count > 0)
                //{ 
                //    var notFindedObjString = string.Join("\r\n", notFindedLinkedObjects.Select(t => t.Name + " " + t.Denotation + " " + t.Type));

                //    Clipboard.SetText(notFindedObjString);
                //    MessageBox.Show("Связанные объекты ненайденные в справочнике Номенклатура и изделия для формирования спецификации:\r\n" + notFindedObjString, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}


                return distinctLinkedObjects;

            }

            private List<ElementPki> GetPKILinkedObjects(List<ElementPki> objectsForSpecification, List<ElementPki> notFindedOjForSpec, TextNormalizer textNormal, Dictionary<string, ReferenceObject> normalNomRefObjects)
            {
                var pkiLinkedObjects = new List<ElementPki>();

                foreach (ElementPki el in objectsForSpecification)
                {
                    if (el.dbElements.Count > 0)
                    {
                        foreach (var linkedObj in el.dbElements[0].LinkedObject)
                        {
                            var elNormalForm = textNormal.GetNormalForm(linkedObj.Name + " " + linkedObj.Denotation);

                            var findedPKI = normalNomRefObjects.Where(t => t.Key == elNormalForm).ToList();
                            if (findedPKI.Count > 0)
                            {
                                ElementPki pki = new ElementPki();

                                if (findedPKI.Count > 1)
                                {
                                    var analogPKI = findedPKI.Where(t => (t.Value[ParametersInfo.Name].Value.ToString() +
                                     t.Value[ParametersInfo.Denotation].Value.ToString()).Trim().ToLower() ==
                                    (linkedObj.Denotation + linkedObj.Name).Trim().ToLower()).Select(t => t.Value).FirstOrDefault();
                                    if (analogPKI != null)
                                    {
                                        pki.Name = analogPKI[ParametersInfo.Name].Value.ToString();
                                        pki.Denotation = analogPKI[ParametersInfo.Denotation].Value.ToString();
                                        pki.Type = analogPKI.Class.Name;
                                        pki.NomObj = analogPKI;
                                        pki.Parent = el.Name + " " + el.Denotation;
                                    }
                                    else
                                    {
                                        pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                                        pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                                        pki.Type = findedPKI[0].Value.Class.Name;
                                        pki.NomObj = findedPKI[0].Value;
                                        pki.Parent = el.Name + " " + el.Denotation;
                                    }
                                }
                                else
                                {
                                    pki.Name = findedPKI[0].Value[ParametersInfo.Name].Value.ToString();
                                    pki.Denotation = findedPKI[0].Value[ParametersInfo.Denotation].Value.ToString();
                                    pki.Type = findedPKI[0].Value.Class.Name;
                                    pki.NomObj = findedPKI[0].Value;
                                    pki.Parent = el.Name + " " + el.Denotation;
                                }

                                pki.PosDenotation = " ";
                                pki.Comment = linkedObj.Comment;
                                pki.Preparation = linkedObj.Preparation;
                                pki.Count = el.Count * linkedObj.Count;
                                pkiLinkedObjects.Add(pki);

                            }
                            else
                            {
                                var notFoundPKI = new ElementPki();
                                notFoundPKI.Name = linkedObj.Name;
                                notFoundPKI.Denotation = linkedObj.Denotation;
                                notFoundPKI.Type = linkedObj.Type;
                                notFoundPKI.PosDenotation = " ";
                                notFoundPKI.Comment = linkedObj.Comment;
                                notFoundPKI.Preparation = linkedObj.Preparation;
                                notFoundPKI.Count = el.Count * linkedObj.Count;

                                if (!notFindedOjForSpec.Contains(notFoundPKI))
                                    notFindedOjForSpec.Add(notFoundPKI);
                            }
                        }
                    }
                }
                return pkiLinkedObjects;
            }
        }
    }
}
