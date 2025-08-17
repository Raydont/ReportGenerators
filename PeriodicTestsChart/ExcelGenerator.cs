using Microsoft.Office.Interop.Excel;
using PeriodicTestsChart;
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

namespace PeriodicTestsChartReport
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

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            m_form.setStage(IndicationForm.Stages.DataAcquisition);

            var devicesForPI = GetDocumentPI(m_writeToLog);
            m_writeToLog("Получено " + devicesForPI.Count + " устройств ПИ");        

            m_writeToLog("=== Получение списка заказов ===");

            var deviseUsageOrders = ProcessingOrders(devicesForPI, reportParameters, m_writeToLog, m_form.progressBar);

            m_writeToLog("=== Формирование отчета ===");
            m_form.setStage(IndicationForm.Stages.ReportGenerating);
            MakeUsageDeviceInOrders(deviseUsageOrders, reportParameters, m_writeToLog, m_form.progressBar);

            m_form.setStage(IndicationForm.Stages.Done);
            m_writeToLog("=== Завершение работы ===");

            m_form.Dispose();
        }


        public static int HeaderTable(Xls xls, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№", col, row, xls);
            col = InsertHeader("Наименование устройств и их шифр", col, row, xls);
            col = InsertHeader("Акт ПИ", col, row, xls);
            col = InsertHeader("Дата окончания действия акта ПИ", col, row, xls);
            col = InsertHeader("Корневые устройства и аппаратура", col, row, xls);
            col = InsertHeader("Исполнители", col, row, xls);

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


        private void MakeUsageDeviceInOrders(Dictionary<ReferenceObject, List<NomenclatureObject>> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            //Задание начальных условий для progressBar
            var countObj = Convert.ToDouble(reportData.Count) * 2;
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

            int row = 1;
            int col = 1;
          
            row = PasteHeadName(xls, reportParameters, col, row, "График ПИ от " + DateTime.Now.Date.ToShortDateString(), 6);
            int iObj = 1;
            HeaderTable(xls, col, row);
            row++;
           

            NomenclatureObject prewNomObj = null;

            foreach (var obj in reportData.OrderBy(t=>t.Key.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature)[Guids.NomenclatureParameters.NameGuid].Value.ToString())
                .ThenBy(t => t.Key[Guids.PeriodicTestsParameters.DatePIDocument].Value))
            {       
                WriteRow(xls, logDelegate, row, iObj, obj.Key, obj.Value, prewNomObj);
                progressBar.PerformStep();
                prewNomObj = (NomenclatureObject)obj.Key.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature);
                iObj++;
                row++;  
            }

            xls.AutoWidth();
            xls.Application.DisplayAlerts = true;
            xls.Visible = true;
        }

        private Dictionary<ReferenceObject, List<NomenclatureObject>> ProcessingOrders(List<ReferenceObject> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            var ordersUsage = new Dictionary<ReferenceObject, List<NomenclatureObject>>();

            var countObj = Convert.ToDouble(reportData.Count)*2;
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

            foreach (ReferenceObject obj in reportData.OrderBy(t=>t.ToString()))
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
                    if (data[mainID].SystemFields.Id ==  data[slaveID].SystemFields.Id)
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
            if (reportParameters.AllOrder)
                devices.AddRange(childObject.Parents.Where(t => t.Class.IsInherit(Guids.NomenclatureType.Complex)).Select(t=>(NomenclatureObject)t).ToList());
            else
                devices.AddRange(childObject.Parents.Where(t => reportParameters.OrdersList.Select(k=>k.SystemFields.Id).Contains(t.SystemFields.Id)).Select(t => (NomenclatureObject)t).ToList());

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

        private static void WriteRow(Xls xls, LogDelegate logDelegate, int row, int i, ReferenceObject piObject, List<NomenclatureObject> orders,  NomenclatureObject prewNomObj)
        {
            int col = 1;
            xls[col, row].SetValue(i.ToString());
            col++;

            var nomObject = piObject.GetObject(Guids.PeriodicTestsParameters.LinkPItoNomenclature);

           
            logDelegate("Запись объекта: " + nomObject.ToString());

            xls[col, row].SetValue(nomObject.ToString());

            col++;
            xls[col, row].NumberFormat = "@";
            xls[col, row].SetValue(piObject[Guids.PeriodicTestsParameters.NumberPIDocument].Value.ToString());
            col++;
            var date = Convert.ToDateTime(piObject[Guids.PeriodicTestsParameters.DatePIDocument].Value).Date;
            if (date != DateTime.MinValue)
            {
                xls[col, row].NumberFormat = "@";
                xls[col, row].SetValue(date.ToShortDateString());    
            }
            col++;
            string orderName = string.Empty;
            var orderNames = new List<string>();
            foreach (var order in orders)
            {
                string codeName = order[Guids.NomenclatureParameters.CodeGuid].Value.ToString();
                
                if (codeName.Trim() == string.Empty)
                {
                    codeName = order[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
                }
                orderNames.Add(codeName);
            }
            if (orderNames != null)
                xls[col, row].SetValue(string.Join("\r\n", orderNames.Select(t=>t)));
            col++;

            var makers = piObject.GetObjects(Guids.PeriodicTestsParameters.ListObjectMaker);

            if (makers.Count > 0)
            {
                xls[col, row].SetValue(string.Join("\r\n", makers.Select(t => t[Guids.PeriodicTestsParameters.NameMaker].Value.ToString())));               
            }

            if (prewNomObj != null && nomObject.SystemFields.Id == prewNomObj.SystemFields.Id)
            {
                xls[1,row, col, 1].Interior.Color = Color.PaleVioletRed;
            }

            xls[1, row, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }


        private List<ReferenceObject> GetDocumentPI(LogDelegate logDelegate)
        {
            logDelegate("Загрузка данных справочника документы ПИ");

            //Создаем ссылку на справочник
            ReferenceInfo refPIInfo = ReferenceCatalog.FindReference(Guids.PeriodicTestsReferenceGuid);
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
