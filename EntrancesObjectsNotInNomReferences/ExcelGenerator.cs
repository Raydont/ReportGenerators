using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace EntrancesObjectsNotInNomReferences
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
                    EntrancesReport.MakeReport(context, selectionForm.reportParameters);
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
    public class EntrancesReport
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            EntrancesReport report = new EntrancesReport();
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

            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

            if (reportData == null)
            {
                m_form.Close();
                return;
            }
                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

              
                if (reportParameters.MissingObjects)
                    MakeMissingObjects(reportData.Select(t => t.NomObject).ToList(), reportParameters, m_writeToLog, m_form.progressBar);

                if (reportParameters.MissingObjectsWithEntrances)
                    MakeMissingObjectsWithEntrances(reportData.Select(t=>t.NomObject).ToList(), reportParameters, m_writeToLog, m_form.progressBar);

                if (reportParameters.ObjectsEqualRefNomWithRemarks)
                    MakeObjectsEqualRefNomWithRemarks(reportData, reportParameters, m_writeToLog, m_form.progressBar);

                if (reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarks)
                    MakeObjectsWithEntrancesEqualRefNomWithRemarks(reportData, reportParameters, m_writeToLog, m_form.progressBar);

                if (reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarksWithOrder)
                    MakeObjectsWithEntrancesEqualRefNomWithRemarksWithOrder(reportData, reportParameters, m_writeToLog, m_form.progressBar);


                if (reportParameters.OrderWithRemarksStdItems)
                {
                    m_writeToLog("=== Обработка данных ===");

                    var ordersWithRemarkSTdItems = ProcessingOrderWithRemarksStdItems(reportData, reportParameters, m_writeToLog);
                    m_writeToLog("=== Формирование отчета2 ===");

                    MakeOrderWithRemarksStdItems(ordersWithRemarkSTdItems.OrderBy(t => t.Order[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()).ToList(), reportParameters, m_writeToLog, m_form.progressBar);
                }

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");
            


            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row, bool order)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№", col, row, xls);
            col = InsertHeader("Наименование", col, row, xls);
            col = InsertHeader("Обозначение", col, row, xls); 
            col = InsertHeader("Заказы", col, row, xls);

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

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row, NomenclatureObject nomObject)
        {
            var name = nomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = nomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();

            if (denotation != string.Empty)
                name += " " + denotation;

            xls[1, row].SetValue(name);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 3, 1].MergeCells = true;

            row = row + 1;

            return row;
        }


        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row, string name, int mergeCol)
        {
            xls[1, row].SetValue(name);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, mergeCol, 1].MergeCells = true;

            row = row + 1;

            return row;
        }


        private void MakeMissingObjectsWithEntrances(List<NomenclatureObject> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Count);

            if (progressBar.Maximum < reportData.Count)
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
            row = PasteHeadName(xls, reportParameters, col, row, "Список объектов типа " + reportParameters.ObjectTypes + ", отсутствующие в Номенклатурном справочнике " + reportParameters.ObjectTypes,3);

            int iObj = 1;


            row++;

            foreach (var obj in reportData)
            {
                int iParentObj = 1;
                WriteRow(xls, logDelegate, row, iObj, obj, true);
                iObj++;
                row++;
                if (obj.Parents.Count() > 0)
                {
                    step += progressStep;
                    if (step >= 1)
                    {
                        progressBar.PerformStep();
                        step = step - progressBar.Step;
                    }

                    HeaderTable(xls, reportParameters, col, row, false);
                    row++;               
                    foreach(NomenclatureObject parent in obj.Parents.OrderBy(t=>t[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                    {
                        WriteRow(xls, logDelegate, row, iParentObj, parent, false);
                        iParentObj++;
                        row++;
                    }
                }
                else
                {
                    xls[2, row].SetValue("Вхождений нет");
                    row++;
                }
                row++;
            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }

        private List<OrderStructure> ProcessingOrderWithRemarksStdItems(List<NomObjectsWithRemarks> reportData, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            List<OrderStructure> orders = new List<OrderStructure>();
            var findedOrders = GetAllOrders();

           // MessageBox.Show( findedOrders.Count + "");

            foreach (NomenclatureObject obj in findedOrders)
            {
                logDelegate("Раскрытие заказа " + obj);
                orders.Add(new OrderStructure(obj, GetChildren( obj, null, reportData), true));
            }

            return orders;
        }

        public List<NomenclatureObject> GetAllOrders()
        {
            ReferenceInfo nomenclatureInfo = ReferenceCatalog.FindReference(Guids.NomenclatureReferenceGuid);
            Reference nomenclatureReference = nomenclatureInfo.CreateReference();
            ClassObject orderClassObject = nomenclatureInfo.Classes.Find(Guids.NomenclatureType.Complex);
            //Создаем фильтр
            TFlex.DOCs.Model.Search.Filter nomFilter = new TFlex.DOCs.Model.Search.Filter(nomenclatureInfo);

            nomFilter.Terms.AddTerm(nomenclatureReference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, orderClassObject);
      //  nomFilter.Terms.AddTerm("ID",
         //      ComparisonOperator.Equal,430857);
          //
            
            nomenclatureReference.LoadSettings.AddParameters(Guids.NomenclatureParameters.NameGuid, Guids.NomenclatureParameters.DenotationGuid);
            var orders = nomenclatureReference.Find(nomFilter).Select(t => (NomenclatureObject)t).ToList();

            nomenclatureReference.RecursiveLoad(orders, RelationLoadSettings.RecursiveLoadDirection.Children, nomenclatureReference.LoadSettings);

            return orders;
        }


        public List<DSEWithStdItems> GetChildren(NomenclatureObject childObject, NomenclatureObject parentObject, List<NomObjectsWithRemarks> reportData)
        {
            var listObjects = new List<DSEWithStdItems>();
            var stdItem = reportData.Where(t => t.NomObject == childObject).FirstOrDefault();

            if (parentObject != null && stdItem != null)
            {
                var findedParent = listObjects.Where(t => t.DSE == parentObject).FirstOrDefault();
                if (findedParent != null)
                {
                  //  findedParent.StandartItems.Add(stdItem);
                }
                else
                {
                    var dse = new DSEWithStdItems(parentObject, stdItem);
                    listObjects.Add(dse);
                }
            }
            else
            {
                foreach (NomenclatureObject child in childObject.Children)
                {
                    listObjects.AddRange(GetChildren(child, childObject, reportData));
                }
            }
            return listObjects;
        }

        public static List<NomenclatureObject> SearchParentOrders(NomenclatureObject nomObj)
        {
            var orders = new List<NomenclatureObject>();

            foreach (NomenclatureObject parent in nomObj.Parents)
            {
                if (parent.Class.IsInherit(Guids.NomenclatureType.Complex))
                {
                    orders.Add(parent);
                }
                else
                {
                    orders.AddRange(SearchParentOrders(parent));
                }
            }

            return orders;
        }

        private void MakeObjectsWithEntrancesEqualRefNomWithRemarks(List<NomObjectsWithRemarks> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Count);

            if (progressBar.Maximum < reportData.Count)
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
            row = PasteHeadName(xls, reportParameters, col, row, "Список объектов типа " + reportParameters.ObjectTypes + ", отсутствующие в Номенклатурном справочнике " + reportParameters.ObjectTypes, 4);

            int iObj = 1;


            row++;

            foreach (var obj in reportData)
            {
                int iParentObj = 1;
                WriteRow(xls, logDelegate, row, iObj, obj, true);
                iObj++;
                row++;
                if (obj.NomObject.Parents.Count() > 0)
                {
                    step += progressStep;
                    if (step >= 1)
                    {
                        progressBar.PerformStep();
                        step = step - progressBar.Step;
                    }

                    HeaderTable(xls, reportParameters, col, row, false);
                    row++;
                    foreach (NomenclatureObject parent in obj.NomObject.Parents.OrderBy(t => t[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                    {
                        WriteRow(xls, logDelegate, row, iParentObj, parent, false);
                        iParentObj++;
                        row++;
                    }
                }
                else
                {
                    xls[2, row].SetValue("Вхождений нет");
                    row++;
                }
                row++;
            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }


        private void MakeObjectsWithEntrancesEqualRefNomWithRemarksWithOrder(List<NomObjectsWithRemarks> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Count);

            if (progressBar.Maximum < reportData.Count)
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
            row = PasteHeadName(xls, reportParameters, col, row, "Список объектов типа " + reportParameters.ObjectTypes + ", отсутствующие в Номенклатурном справочнике " + reportParameters.ObjectTypes, 4);

            int iObj = 1;


            row++;

            foreach (var obj in reportData.OrderBy(t=>t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()))
            {
                int iParentObj = 1;
                WriteRow(xls, logDelegate, row, iObj, obj, true);
                iObj++;
                row++;
                if (obj.NomObject.Parents.Count() > 0)
                {
                    step += progressStep;
                    if (step >= 1)
                    {
                        progressBar.PerformStep();
                        step = step - progressBar.Step;
                    }

                    HeaderTable(xls, reportParameters, col, row, true);
                    row++;
                    foreach (NomenclatureObject parent in obj.NomObject.Parents.OrderBy(t => t[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                    {
                        WriteRowWithOrder(xls, logDelegate, row, iParentObj, parent);
                        iParentObj++;
                        row++;
                    }
                }
                else
                {
                    xls[2, row].SetValue("Вхождений нет");
                    row++;
                }
                row++;
            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }

        private void MakeOrderWithRemarksStdItems(List<OrderStructure> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
          
            //  MessageBox.Show(reportData.Sum(t => t.ListSTD.Sum(r => r.DSE.Count)) + "");
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Sum(t=>t.ListSTD.Sum(r=>r.DSE.Count)));

            if (progressBar.Maximum < reportData.Sum(t => t.ListSTD.Sum(r => r.DSE.Count)))
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            double step = 0;



            StringBuilder strOrdersError = new StringBuilder();
            foreach (var obj in reportData)
            {
                StringBuilder orderListStrB = new StringBuilder();
                var numberRowStdItem = new List<int>();

                try
                {
                    Xls xls = new Xls();
                    xls.Application.DisplayAlerts = false;

                    //   xls.AddWorksheet(obj.Order[Guids.NomenclatureParameters.DenotationGuid].Value.ToString(), true);
                    int row = 1;
                    int col = 1;
                    row = PasteHeadName(xls, reportParameters, col, row, "Список Стандартных изделий со списком прямых вхождений для заказа " + obj.Order, 8);

                    int iObj = 1;

                    row++;

                    //   WriteOrder(xls, logDelegate, row, iObj, obj.Order);
                    // iObj++;
                    // row++;
                    int iParentObj = 1;
                    if (obj.ListSTD.Count() > 0)
                    {
                        foreach (var std in obj.ListSTD.OrderBy(t => t.StandartItem.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString()))
                        {
                            numberRowStdItem.Add(row);
                            orderListStrB.AppendLine(iParentObj + "\t" +
                                std.StandartItem.NomObject[Guids.NomenclatureParameters.NameGuid].Value + "\t" +
                                std.StandartItem.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value + "\t" +
                                std.StandartItem.Remark);
                            // WriteRow(xls, logDelegate, row, iParentObj, std.StandartItem, true);
                            iParentObj++;
                            row++;
                            int istd = 1;
                            foreach (NomenclatureObject stdItem in std.DSE.OrderBy(t => t[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                            {
                                step += progressStep;
                                if (step >= 1)
                                {
                                    progressBar.PerformStep();
                                    step = step - progressBar.Step;
                                }
                                orderListStrB.AppendLine(istd + "\t" +
                              stdItem[Guids.NomenclatureParameters.NameGuid].Value + "\t" +
                              stdItem[Guids.NomenclatureParameters.DenotationGuid].Value);
                                // WriteRow(xls, logDelegate, row, istd, stdItem, false);
                                istd++;
                                row++;
                            }
                        }
                    }
                    //else
                    //{
                    //    xls[2, row].SetValue("Вхождений нет");
                    //    row++;
                    //}

                    Clipboard.SetText(orderListStrB.ToString());
                    xls.Worksheet.Select(xls[1, 3, 1, 1]);
                    xls.Worksheet.Paste(xls[1, 3, 1, 1]);
                    foreach (var indx in numberRowStdItem)
                    {
                        xls[1, indx, 4, 1].Interior.Color = Color.LightGray;
                    }

                    xls[1, 1, 4, row - 1].BorderTable(weight: XlBorderWeight.xlThin);
                    xls.AutoWidth();
                    xls.Application.DisplayAlerts = true;
                    // xls.SelectWorksheet("Лист1");
                    // xls.Worksheet.Delete();
                    var nameFile = obj.Order[Guids.NomenclatureParameters.DenotationGuid].Value + " " + obj.Order[Guids.NomenclatureParameters.CodeGuid].Value;
                    nameFile = nameFile.Replace("#", "").Replace("%", "").Replace("&", "").Replace("*", "").Replace("{", "").Replace("}", "").Replace(@"\", "").Replace(@"/", "").
                        Replace(":", "").Replace("<", "").Replace(">", "").Replace("?", "").Replace("+", "").Replace("|", "").Replace("\"", "").Replace("~", "");
                    xls.SaveAs(reportParameters.PathForOrdersFiles + @"\" + nameFile + ".xlsx");

                    xls.Close();
                }
                catch
                {
                    strOrdersError.AppendLine(obj.Order.ToString());

                }

            }
            System.Diagnostics.Process.Start(reportParameters.PathForOrdersFiles);
            if (strOrdersError.ToString() != string.Empty)
            {
                Clipboard.SetText(strOrdersError.ToString());
                MessageBox.Show("Файлы выгружены. Не сформированы файлы для заказов:\r\n" + strOrdersError);
            }
         
        }

        private void MakeMissingObjects(List<NomenclatureObject> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Count);

            if (progressBar.Maximum < reportData.Count)
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
            row = PasteHeadName(xls, reportParameters, col, row, "Список объектов типа " + reportParameters.ObjectTypes + ", имеющих примечание в Номенклатурном справочнике " + reportParameters.ObjectTypes,4);

            int iObj = 1;


            row++;

            foreach (var obj in reportData)
            {
                int iParentObj = 1;
                WriteRow(xls, logDelegate, row, iObj, obj, false);
                step += progressStep;
                if (step >= 1)
                {
                    progressBar.PerformStep();
                    step = step - progressBar.Step;
                }
                iObj++;
                row++;
             
            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }


        private void MakeObjectsEqualRefNomWithRemarks(List<NomObjectsWithRemarks> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            double progressStep = Convert.ToDouble(progressBar.Maximum) / Convert.ToDouble(reportData.Count);

            if (progressBar.Maximum < reportData.Count)
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
            row = PasteHeadName(xls, reportParameters, col, row, "Список заказов с объектами типа " + reportParameters.ObjectTypes + ", имеющих примечание в Номенклатурном справочнике " + reportParameters.ObjectTypes,4);

            int iObj = 1;


            row++;

            foreach (var obj in reportData)
            {
                int iParentObj = 1;
                WriteRow(xls, logDelegate, row, iObj, obj, false);
                step += progressStep;
                if (step >= 1)
                {
                    progressBar.PerformStep();
                    step = step - progressBar.Step;
                }
                iObj++;
                row++;

            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }
        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, NomenclatureObject nomObject, bool mainObject)
        {
            int col = 1;
            xls[col, row].SetValue(i.ToString());
            col++;

            var name = nomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = nomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            logDelegate("Запись объекта: " + name + " " + denotation);

            xls[col, row].SetValue(name);
            col++;
            if (denotation.Trim() != string.Empty)
            {
                xls[col, row].SetValue(denotation);
            }

            if (mainObject)
            {
                for (int j = 1; j < col+1; j++)
                {
                    xls[j, row].Font.Bold = true;
                }

                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                                         XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
            else
                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                                        XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteOrder(Xls xls, LogDelegate logDelegate, int row, int i, NomenclatureObject order)
        {
            int col = 1;
            xls[col, row].SetValue(i.ToString());
            col++;

            var name = order[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = order[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            logDelegate("Запись объекта: " + name + " " + denotation);

            xls[col, row].SetValue(name);
            col++;
            if (denotation.Trim() != string.Empty)
            {
                xls[col, row].SetValue(denotation);
            }


            for (int j = 1; j < col + 1; j++)
            {
                xls[j, row].Font.Bold = true;
                xls[j, row].Interior.Color = Color.LightGray;
            }
            

            xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                                     XlBordersIndex.xlEdgeBottom,
                                                                                                     XlBordersIndex.xlEdgeLeft,
                                                                                                     XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);

        }

        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, NomObjectsWithRemarks nomObject, bool mainObject)
        {
            int col = 1;
            xls[col, row].SetValue(i.ToString());
            col++;

            var name = nomObject.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = nomObject.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            var remark = nomObject.Remark;
            logDelegate("Запись объекта: " + name + " " + denotation);

            xls[col, row].SetValue(name);
            col++;
            if (denotation.Trim() != string.Empty)
            {
                xls[col, row].SetValue(denotation);
            }
            col++;

            if (remark.Trim() != string.Empty)
            {
                xls[col, row].SetValue(remark);
            }

            if (mainObject)
            {
                for (int j = 1; j < col + 1; j++)
                {
                 //   xls[j, row].Font.Bold = true;
                    xls[j, row].Interior.Color = Color.LightBlue;
                }

                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlBordersIndex.xlEdgeTop,
                                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                                         XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
            }
            else
                xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                                        XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteRowWithOrder(Xls xls, LogDelegate logDelegate, int row, int i, NomenclatureObject nomObject)
        {
            int col = 1;
            xls[col, row].SetValue(i.ToString());
            col++;

            var name = nomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            var denotation = nomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            var order = SearchParentOrders(nomObject);

            logDelegate("Запись объекта: " + name + " " + denotation);

            xls[col, row].SetValue(name);
            col++;
            if (denotation.Trim() != string.Empty)
            {
                xls[col, row].SetValue(denotation);
            }
            col++;
  
            xls[col, row].SetValue(string.Join("\r\n", order.Select(t=>t[Guids.NomenclatureParameters.CodeGuid].Value.ToString()).Where(t=>t.Trim() != string.Empty).OrderBy(t=>t).ToList()));
            
            xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                                        XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }



        private List<NomObjectsWithRemarks> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");

            Guid selectedNomTypeGuid;
            Guid selectedSpecNomTypeGuid;
            Guid nomSpecReferenceGuid;
            Guid nameSpecNomRefGuid;
            Guid denotationSpecNomRefGuid;
            Guid remarkSpecNomRefGuid;

            if (reportParameters.StandartItemsReference)
            {
                selectedNomTypeGuid = Guids.NomenclatureType.StandartItem;
                selectedSpecNomTypeGuid = Guids.SpecNomRefType.StandartItemsGuid;
                nomSpecReferenceGuid = Guids.StandartItemsReferenceGuid;
                nameSpecNomRefGuid = Guids.SpecNomParameters.NameSIGuid;
                denotationSpecNomRefGuid = Guids.SpecNomParameters.DenotationSIGuid;
                remarkSpecNomRefGuid = Guids.SpecNomParameters.RemarkSIGuid;
                reportParameters.ObjectTypes = "Стандартные изделия";
            }
            else
            {
                if (reportParameters.OtherItemsReference)
                {
                    selectedNomTypeGuid = Guids.NomenclatureType.OtherItem;
                    selectedSpecNomTypeGuid = Guids.SpecNomRefType.OtherItemsGuid;
                    nomSpecReferenceGuid = Guids.OtherItemsReferenceGuid;
                    nameSpecNomRefGuid = Guids.SpecNomParameters.NameOIGuid;
                    denotationSpecNomRefGuid = Guids.SpecNomParameters.DenotationOIGuid;
                    remarkSpecNomRefGuid = Guids.SpecNomParameters.RemarkOIGuid;
                    reportParameters.ObjectTypes = "Прочие(покупные) изделия";
                }
                else
                {
                    selectedNomTypeGuid = Guids.NomenclatureType.Material;
                    selectedSpecNomTypeGuid = Guids.SpecNomRefType.MaterialGuid;
                    nomSpecReferenceGuid = Guids.MaterialsReferenceGuid;
                    nameSpecNomRefGuid = Guids.SpecNomParameters.NameMaterialsGuid;
                    denotationSpecNomRefGuid = Guids.SpecNomParameters.DenotationMaterialsGuid;
                    remarkSpecNomRefGuid = Guids.SpecNomParameters.RemarkMaterialGuid;
                    reportParameters.ObjectTypes = "Материалы";
                }
            }

            logDelegate("Загрузка объектов справочника " + reportParameters.ObjectTypes);
            //Создаем ссылку на справочник
            ReferenceInfo nomSpecInfo = ReferenceCatalog.FindReference(nomSpecReferenceGuid);
            Reference nomeSpecReference = nomSpecInfo.CreateReference();

            ClassObject nomSpecClassObject = nomSpecInfo.Classes.Find(selectedSpecNomTypeGuid);
            //Создаем фильтр
            TFlex.DOCs.Model.Search.Filter nomSpecFilter = new TFlex.DOCs.Model.Search.Filter(nomSpecInfo);

            nomSpecFilter.Terms.AddTerm(nomeSpecReference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, nomSpecClassObject);

            if (reportParameters.MissingObjects || reportParameters.MissingObjectsWithEntrances)
            {
                nomeSpecReference.LoadSettings.AddParameters(nameSpecNomRefGuid, denotationSpecNomRefGuid);
            }

            if (reportParameters.ObjectsEqualRefNomWithRemarks || 
                reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarks ||
                reportParameters.OrderWithRemarksStdItems ||
                reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarksWithOrder)
            {
                nomSpecFilter.Terms.AddTerm(nomeSpecReference.ParameterGroup[remarkSpecNomRefGuid], ComparisonOperator.IsNotEmptyString, null);
                nomeSpecReference.LoadSettings.AddParameters(nameSpecNomRefGuid, denotationSpecNomRefGuid, remarkSpecNomRefGuid);                
            }

            //Получаем список объектов, в качестве условия поиска – сформированный фильтр 
            var nomSpecRefObjects = nomeSpecReference.Find(nomSpecFilter).ToList();
            var nomSpecObjectsString = nomSpecRefObjects.Select(t=>t[nameSpecNomRefGuid].Value.ToString() + t[denotationSpecNomRefGuid].Value.ToString()).ToList();

            logDelegate("Загрузка объектов справочника Номенклатура");
            //Создаем ссылку на справочник
            ReferenceInfo nomenclatureInfo = ReferenceCatalog.FindReference(Guids.NomenclatureReferenceGuid);
            Reference nomenclatureReference = nomenclatureInfo.CreateReference();

            ClassObject nomClassObject = nomenclatureInfo.Classes.Find(selectedNomTypeGuid);
            //Создаем фильтр
            TFlex.DOCs.Model.Search.Filter nomFilter = new TFlex.DOCs.Model.Search.Filter(nomenclatureInfo);

            nomFilter.Terms.AddTerm(nomenclatureReference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, nomClassObject);
            nomenclatureReference.LoadSettings.AddParameters(Guids.NomenclatureParameters.NameGuid, Guids.NomenclatureParameters.DenotationGuid);

            var nomObjects = nomenclatureReference.Find(nomFilter);

            List<NomObjectsWithRemarks> reportData = new List<NomObjectsWithRemarks>();

            if (reportParameters.MissingObjects || reportParameters.MissingObjectsWithEntrances)
            {
                var listNomObj = nomObjects.Where(t => !nomSpecObjectsString.Contains(t[Guids.NomenclatureParameters.NameGuid].Value.ToString() + t[Guids.NomenclatureParameters.DenotationGuid].Value.ToString()))
                                 .Select(t => (NomenclatureObject)t).OrderBy(t => t.ToString()).ToList();
                foreach (var obj in listNomObj)
                {
                    reportData.Add(new NomObjectsWithRemarks(obj, string.Empty));
                }
            }

            if (reportParameters.ObjectsEqualRefNomWithRemarks ||
                reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarks || 
                reportParameters.OrderWithRemarksStdItems ||
                reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarksWithOrder)
            {
                foreach (NomenclatureObject nomObj in nomObjects)
                {
                    foreach (var specObj in nomSpecRefObjects)
                    {
                        if (specObj[nameSpecNomRefGuid].Value.ToString() + specObj[denotationSpecNomRefGuid].Value.ToString() == 
                            nomObj[Guids.NomenclatureParameters.NameGuid].Value.ToString() + nomObj[Guids.NomenclatureParameters.DenotationGuid].Value.ToString())
                        reportData.Add(new NomObjectsWithRemarks(nomObj, specObj[remarkSpecNomRefGuid].Value.ToString()));
                    }
                }
            }

            logDelegate("Считано объектов: " + reportData.Count);
            return reportData;
        }
    }
}
