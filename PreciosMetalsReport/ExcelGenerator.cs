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
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using Filter = TFlex.DOCs.Model.Search.Filter;

namespace PreciosMetalsReport
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
                        EntrancesReport.MakeReport(context, selectionForm.reportParameters);
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

          //  m_form.setStage(IndicationForm.Stages.Initialization);
          //  m_writeToLog("=== Инициализация ===");

           // m_form.setStage(IndicationForm.Stages.DataAcquisition);

           // var devicesForPI = GetDocumentPI(m_writeToLog);
            //m_writeToLog("Получено " + devicesForPI.Count + " устройств ПИ");

           // m_writeToLog("=== Получение списка заказов ===");

           // var deviseUsageOrders = ProcessingOrders(devicesForPI, reportParameters, m_writeToLog, m_form.progressBar);

            m_writeToLog("=== Формирование отчета ===");
            m_form.setStage(IndicationForm.Stages.ReportGenerating);
            MakeUsageDeviceInOrders(reportParameters, m_writeToLog, m_form.progressBar, context);

            m_form.setStage(IndicationForm.Stages.Done);
            m_writeToLog("=== Завершение работы ===");

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, int col, int row)
        {
            var beginCol = col;
            var beginRow = row;
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("Наименование", col, row, xls);
            var colName = col;         
            col = InsertHeader("Содержание д/м", col, row, xls);
            var colCont = col;
            col = InsertHeader("Остаток 01.01." + (DateTime.Now.Year - 1), col, row, xls);
            xls[col - 1, row, 2, 1].Merge();
            col++;
            col = InsertHeader("Приход", col, row, xls);
            xls[col-1, row, 2, 1].Merge();
            col++;    
            col = InsertHeader("Расход", col, row, xls);
            xls[col - 1, row, 2, 1].Merge();
            col++;
            var colEx = col;
            col = InsertHeader("Остаток 01.01." + DateTime.Now.Year, col, row, xls);
            xls[col - 1, row, 2, 1].Merge();
            col++;
            col = InsertHeader("Проверка расхода", col, row, xls);
            xls[colName-1, row, 1, 2].Merge();
            xls[colCont-1, row, 1, 2].Merge();
            xls[col - 1, row, 1, 2].Merge();
            xls[1, row, 11, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                               XlBordersIndex.xlEdgeBottom,
                                                                                               XlBordersIndex.xlEdgeLeft,
                                                                                               XlBordersIndex.xlEdgeRight,
                                                                                               XlBordersIndex.xlInsideVertical);
            /*SetBorders(
              DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
               XlsxBorder.Top,
               XlsxBorder.Bottom,
               XlsxBorder.Left,
               XlsxBorder.Right,
               XlsxBorder.InsideVertical);*/

            row++;
            col = 3;

            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Сод-е", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Сод-е", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Сод-е", col, row, xls);
            col = InsertHeader("Кол-во", col, row, xls);
            col = InsertHeader("Сод-е", col, row, xls);

            xls[1, row, 11, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                 XlBordersIndex.xlEdgeBottom,
                                                                                                 XlBordersIndex.xlEdgeLeft,
                                                                                                 XlBordersIndex.xlEdgeRight,
                                                                                                 XlBordersIndex.xlInsideVertical);

            /* xls[1, row, 8, 1].SetBorders(
                 DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
                  XlsxBorder.Top,
                  XlsxBorder.Bottom,
                  XlsxBorder.Left,
                  XlsxBorder.Right,
                  XlsxBorder.InsideVertical);*/
            xls[1, row-1, 11, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row-1, 11, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[1, row-1, 11, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row-1, 11, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;


           

          //  xls.Application.ActiveWindow.FreezePanes = true;

            return col;
        }


        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            //  xls[col, row].VerticalAlignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center;
            //   xls[col, row].HorizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
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
            //  xls[1, row].HorizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, row, 3, 1].Merge();

            row = row + 1;

            return row;
        }


        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row, string name, int mergeCol)
        {
            xls[1, row].SetValue(name);
            xls[1, row].Font.Bold = true;
            xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //   xls[1, row].HorizontalAlignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
            xls[1, row, mergeCol, 1].Merge();

            row = row + 1;

            return row;
        }


        private void MakeUsageDeviceInOrders(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar,IReportGenerationContext context)
        {
            // XlsxDocument xls = new XlsxDocument();
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            xls.Open(context.ReportFilePath);
            //Задание начальных условий для progressBar
            var countObj = Convert.ToDouble(reportParameters.PKIData.Count);
            if (countObj == 0)
                return;
            double progressStep = Convert.ToDouble(progressBar.Maximum) / countObj;

            if (progressBar.Maximum < countObj)
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            double step = 0;
            xls.AddWorksheet("Содержание золота", true);
            int row = 1;
            int col = 1;

            row = PasteHeadName(xls, reportParameters, col, row, "Отчет по драгметаллам за " + (DateTime.Now.Date.Year - 1) + " год", 6);
            row++;
            HeaderTable(xls, col, row);

          //  xls[1, 1].Select();
          //  xls.Application.ActiveWindow.FreezePanes = true;
            row = row + 2;



            var beginRow = row;
            foreach (var obj in reportParameters.PKIData.OrderBy(t => t.Name).Where(t=>Convert.ToDouble(t.StoreObject[Guids.NomStore.GoldContains].Value) > 0))
            {
                var containsMetalls = Convert.ToDouble(obj.StoreObject[Guids.NomStore.GoldContains].Value);
                WriteRow(xls, logDelegate, row,obj, containsMetalls);
                progressBar.PerformStep();
                row++;
            }
            WriteSummary(xls, logDelegate, beginRow, row);

            xls.AutoWidth();
            xls.AddWorksheet("Содержание серебра", true);

            row = 1;
            col = 1;

            row = PasteHeadName(xls, reportParameters, col, row, "Отчет по драгметаллам за " + (DateTime.Now.Date.Year - 1) + " год", 6);
            row++;
            HeaderTable(xls, col, row);
            row = row + 2;


            beginRow = row;
            foreach (var obj in reportParameters.PKIData.OrderBy(t => t.Name).Where(t => Convert.ToDouble(t.StoreObject[Guids.NomStore.SilverContains].Value) > 0))
            {
                var containsMetalls = Convert.ToDouble(obj.StoreObject[Guids.NomStore.SilverContains].Value);
                WriteRow(xls, logDelegate, row, obj, containsMetalls);
                progressBar.PerformStep();
                row++;
            }
            WriteSummary(xls, logDelegate, beginRow, row);
            xls.AutoWidth();
            xls.AddWorksheet("Содержание платины", true);

            row = 1;
            col = 1;

            row = PasteHeadName(xls, reportParameters, col, row, "Отчет по драгметаллам за " + (DateTime.Now.Date.Year - 1) + " год", 6);
            row++;
            HeaderTable(xls, col, row);
            row = row + 2;

            beginRow = row;
            foreach (var obj in reportParameters.PKIData.OrderBy(t => t.Name).Where(t => Convert.ToDouble(t.StoreObject[Guids.NomStore.PlatinaContains].Value) > 0))
            {
                var containsMetalls = Convert.ToDouble(obj.StoreObject[Guids.NomStore.PlatinaContains].Value);
                WriteRow(xls, logDelegate, row, obj, containsMetalls);
                progressBar.PerformStep();
     
                row++;
            }
            WriteSummary(xls, logDelegate, beginRow, row);
            xls.AutoWidth();
            //   xls.sheets.Remove(xls.sheets.[0]);
            xls.SelectWorksheet("Лист1");
            xls.Worksheet.Delete();
            xls.SelectWorksheet("Содержание золота");
            xls.AutoWidth();
            //  xls.SaveAs(context.ReportFilePath);
            xls.Save();
            xls.Quit(true);
            xls.Application.DisplayAlerts = true;
            System.Diagnostics.Process.Start(context.ReportFilePath);
            //   progressForm.progressBarMakeReport.PerformStep();
            //   progressForm.Close();

           
            // Открытие файла
            //   System.Diagnostics.Process.Start(context.ReportFilePath);
        }

        private Dictionary<ReferenceObject, List<NomenclatureObject>> ProcessingOrders(List<ReferenceObject> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            var ordersUsage = new Dictionary<ReferenceObject, List<NomenclatureObject>>();

            var countObj = Convert.ToDouble(reportData.Count);
            if (countObj == 0)
                return null;
            double progressStep = Convert.ToDouble(progressBar.Maximum) / countObj;

            if (progressBar.Maximum < countObj)
            {
                progressBar.Step = 1;
            }
            else
            {
                progressBar.Step = Convert.ToInt32(Math.Round(progressStep));
            }

            foreach (ReferenceObject obj in reportData.OrderBy(t => t.ToString()))
            {
                var findedDevices = new List<NomenclatureObject>();
                progressBar.PerformStep();
                var nomObj = (NomenclatureObject)obj.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature);
                findedDevices = GetParents(nomObj, reportParameters);
                logDelegate("Поиск заказов для " + nomObj);
                ordersUsage.Add(obj, findedDevices);
            }

            return ordersUsage;
        }

        private void RemoveDublicates(List<NomenclatureObject> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].SystemFields.Id == data[slaveID].SystemFields.Id)
                    {
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
        }

        public List<NomenclatureObject> GetAllOrders()
        {
            ReferenceInfo nomenclatureInfo = ReferenceCatalog.FindReference(Guids.NomenclatureReferenceGuid);
            Reference nomenclatureReference = nomenclatureInfo.CreateReference();
            ClassObject orderClassObject = nomenclatureInfo.Classes.Find(Guids.NomenclatureType.Complex);
            //Создаем фильтр
            Filter nomFilter = new Filter(nomenclatureInfo);

            nomFilter.Terms.AddTerm(nomenclatureReference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, orderClassObject);

            nomenclatureReference.LoadSettings.AddParameters(Guids.NomenclatureParameters.NameGuid, Guids.NomenclatureParameters.DenotationGuid);
            var orders = nomenclatureReference.Find(nomFilter).Select(t => (NomenclatureObject)t).ToList();

            nomenclatureReference.RecursiveLoad(orders, RelationLoadSettings.RecursiveLoadDirection.Children, nomenclatureReference.LoadSettings);

            return orders;
        }


        public List<NomenclatureObject> GetParents(NomenclatureObject childObject, ReportParameters reportParameters)
        {
            var listObjects = new List<NomenclatureObject>();
            var devices = new List<NomenclatureObject>();

                devices.AddRange(childObject.Parents.Where(t => t.Class.IsInherit(Guids.NomenclatureType.Complex)).Select(t => (NomenclatureObject)t).ToList());
         
            if (devices != null)
            {
                listObjects.AddRange(devices);
            }
            else
            {
                foreach (NomenclatureObject parent in childObject.Parents)
                {
                    listObjects.AddRange(GetParents(parent, reportParameters));
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

        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, StoreData storeData, double containsMetall)
        {
            int col = 1;

            logDelegate("Запись объекта: " + storeData.Name);

            xls[col, row].SetValue(storeData.Article + " " + storeData.Name);

            col++;
            xls[col, row].NumberFormat = "@";
            xls[col, row].SetValue(containsMetall);
            var addressContains = xls[col, row].Address;
            col++;
            xls[col, row].NumberFormat = "@";
            xls[col, row].SetValue(storeData.RemainsBeginCount);
            var addressRemainsBegin = xls[col, row].Address;
            col++;
            xls[col, row].SetFormula(@"=ПРОИЗВЕД(" + addressRemainsBegin + ";" + addressContains + ")");
            var addressRemainsBeginContains = xls[col, row].Address;
            col++;
            xls[col, row].NumberFormat = "@";
            xls[col, row].SetValue(storeData.ComingCount);
            var addressComing = xls[col, row].Address;
            col++;
            xls[col, row].SetFormula(@"=ПРОИЗВЕД(" + addressContains + ";" + addressComing + ")");
            var addressComingContains = xls[col, row].Address;
            col++;
            //xls[col, row].NumberFormat = "@";
            var colEx = col;

            xls[col, row].NumberFormat = "@";
            xls[col, row].SetValue(storeData.ExpenditureCount);
            var addressExpenditure = xls[col, row].Address;
            col++;
            xls[col, row].SetFormula(@"=ПРОИЗВЕД(" + addressContains + ";" + addressExpenditure + ")");
      

            col++;
            xls[col, row].SetValue(storeData.RemainsEndCount);
            var addressRemainsEndContains = xls[col, row].Address;
          
            col++;
            xls[col, row].SetFormula(@"=ПРОИЗВЕД(" + addressContains + ";" + addressRemainsEndContains + ")");
            var addressSummaryContains = xls[col, row].Address;
            col++;
            xls[col, row].SetFormula(@"=" + addressRemainsBeginContains + "+" + addressComingContains + "-" + addressSummaryContains);
            xls[1, row, col,1].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            /* xls[1, row, col, 1].SetBorders(DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin,
                  XlsxBorder.Top,
                  XlsxBorder.Bottom,
                  XlsxBorder.Left,
                  XlsxBorder.Right,
                  XlsxBorder.InsideVertical);*/
            xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                         XlBordersIndex.xlEdgeBottom,
                                                                                         XlBordersIndex.xlEdgeLeft,
                                                                                         XlBordersIndex.xlEdgeRight,
                                                                                         XlBordersIndex.xlInsideVertical);
        }

        private static void WriteSummary(Xls xls, LogDelegate logDelegate, int beginRow, int row)
        {
            int col = 1;

            xls[col, row].SetValue("Итого:");

            col = col + 3;
            xls[col, row].SetFormula(@"=СУММ(" + xls[col,beginRow].Address + ":" + xls[col, row - 1].Address + ")");
            col = col + 2;
            xls[col, row].SetFormula(@"=СУММ(" + xls[col, beginRow].Address + ":" + xls[col, row - 1].Address + ")");
            col = col + 2;
            xls[col, row].SetFormula(@"=СУММ(" + xls[col, beginRow].Address + ":" + xls[col, row - 1].Address + ")");
            col = col + 2;
            xls[col, row].SetFormula(@"=СУММ(" + xls[col, beginRow].Address + ":" + xls[col, row - 1].Address + ")");
            col++;
            xls[col, row].SetFormula(@"=СУММ(" + xls[col, beginRow].Address + ":" + xls[col, row - 1].Address + ")");
        }


        private List<ReferenceObject> GetDocumentPI(LogDelegate logDelegate)
        {
            logDelegate("Загрузка данных справочника документы ПИ");

            //Создаем ссылку на справочник
            ReferenceInfo refPIInfo = ReferenceCatalog.FindReference(Guids.NomenclatureStoreGuid);
            Reference referencePI = refPIInfo.CreateReference();

            ClassObject classObject = referencePI.Classes.Find(Guids.PeriodicTestsParameters.TypePI);

            //Создаем фильтр
            Filter filter = new Filter(refPIInfo);
            filter.Terms.AddTerm(referencePI.ParameterGroup[SystemParameterType.Class],
                   ComparisonOperator.Equal, classObject);
            List<ReferenceObject> listObj = referencePI.Find(filter)
                .Where(t => t.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature) != null).ToList();
            // .Select(t=>(NomenclatureObject)t.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature)).ToList();

            return listObj;
        }
    }
}
