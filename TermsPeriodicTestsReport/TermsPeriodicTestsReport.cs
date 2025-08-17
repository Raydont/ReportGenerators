using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using Microsoft.Office.Interop.Excel;
using TermsPeriodicTestsReport;
using System.Drawing;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Search;
using System.Text.RegularExpressions;
using TFlex.DOCs.Model.References.Nomenclature;

namespace TermsPeriodicTestsReport
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
            using (new WaitCursorHelper(false))
            {
                selectionForm.Init(context);
                if (selectionForm.ShowDialog() == DialogResult.OK)
                {                   
                    var reportParameters = selectionForm.reportParameters;

                    TermsPeriodicTestsReport termPeriodic = new TermsPeriodicTestsReport();
                    termPeriodic.MakeReport(context, reportParameters);
                }
                
            }
        }
    }

    public delegate void LogDelegate(string line);

    public class TermsPeriodicTestsReport
    {

        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            TermsPeriodicTestsReport report = new TermsPeriodicTestsReport();

            LogDelegate m_writeToLog;
            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
        {
            var m_writeToLog = new LogDelegate(m_form.writeToLog);
            context.CopyTemplateFile();    // Создаем копию шаблона
            m_form.progressBar.PerformStep();
            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);

            m_writeToLog("=== Инициализация ===");

            Xls xls = new Xls();

            TermsPeriodicTestsReport report = new TermsPeriodicTestsReport();

            m_form.setStage(IndicationForm.Stages.DataProcessing);
            m_writeToLog("=== Обработка данных ===");

            var reportData = GetData(reportParameters, m_writeToLog, m_form.progressBar);

            RemoveDublicatePIWithLessDate(reportData);

            m_form.progressBar.PerformStep();
            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeSelectionByTypeReport(xls, reportData, reportParameters, m_writeToLog, m_form.progressBar);

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

        public void RemoveDublicatePIWithLessDate(List<PTDocument> data)
        {
            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].NomenclatureObject.SystemFields.Id == data[slaveID].NomenclatureObject.SystemFields.Id && !data[mainID].Name.ToLower().Contains("кабель"))
                    {
                        if (data[mainID].Date.Date >= data[slaveID].Date.Date)
                        {
                            data.RemoveAt(slaveID);
                            slaveID--;
                        }
                        else
                        {
                            data.RemoveAt(mainID);
                            slaveID--;
                        }
                    }
                }
            }
        }
        public static int HeaderTable(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("№", col, row, xls);
            col = InsertHeader("Обозначение", col, row, xls);
            col = InsertHeader("Наименование", col, row, xls);
            col = InsertHeader("Литера", col, row, xls);
            col = InsertHeader("Документа ПИ", col, row, xls);
            col = InsertHeader("Дата ПИ", col, row, xls);
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

        public int PasteHeadName(Xls xls, ReportParameters reportParameters, int col, int row)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(reportParameters.OrderName.ToString());
            xls[col, row].Font.Bold = true;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row, 7, 1].MergeCells = true;
            row = row + 2;

            return row;
        }

        private void MakeSelectionByTypeReport(Xls xls, List<PTDocument> reportData, ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = reportData.Count;
            int i = 0;
            int step = 0;

            int row = 1;
            int col = 1;
            string department = string.Empty;
            row = PasteHeadName(xls, reportParameters, col, row);
            HeaderTable(xls, reportParameters, col, row);

            xls[1, row + 1, 1, 11].Select();
            xls.Application.ActiveWindow.FreezePanes = true;
            row++;

            var curValProgressBar = progressBar.Value;
            var restValProgressBar = progressBar.Maximum - curValProgressBar;

            if (restValProgressBar > reportData.Count)
                progressBar.Step = (int)Math.Round((double)restValProgressBar / (double)reportData.Count);


            double partProgressBar = (double)restValProgressBar / (double)reportData.Count;

            double sumPartProgressBar = 0;
            //Запись строки

            foreach (var obj in reportData.OrderBy(t=>t.Date).ThenBy(t=>t.NomenclatureObject[Guids.Nomenclature.Denotation].Value.ToString()).ToList())
            {
                i++;
                WriteRow(xls,  logDelegate, reportParameters, row, i, obj);
                row++;
                sumPartProgressBar = sumPartProgressBar + partProgressBar;

                if (restValProgressBar < reportData.Count && sumPartProgressBar >= 1)
                {
                    progressBar.PerformStep();
                    sumPartProgressBar = sumPartProgressBar - 1;
                }

                if (restValProgressBar > reportData.Count)
                    progressBar.PerformStep();
            }

            xls[1, 1, 1, 1].Select();
            xls.AutoWidth();
        }


        //Получение всех сборок, входящих в состав выбранного заказа/сборки и поиск этих сборок в справочнике "Документы для периодических испытаний" 
        private List<PTDocument> ProcessData(ReportParameters reportParameters, LogDelegate logDelegate)
        {
            logDelegate("Загрузка дерева состава для {0}" + reportParameters.OrderName);
            List<PTDocument> structuredData = new List<PTDocument>();
            ReferenceInfo referenceCategoriesInfo = ReferenceCatalog.FindReference(Guids.References.Nomenclature);
            Reference categoriesOtherItemsReference = referenceCategoriesInfo.CreateReference();

            logDelegate("Загрузка объектов из справочника Категории прочих изделий");
            categoriesOtherItemsReference.Objects.Load();
            logDelegate("Загружено " + categoriesOtherItemsReference.Objects.ToList().Count + " объектов");

            return structuredData;
        }

        //Вставка рамки
        private static void InsertFrame(Xls xls, int row, int width)
        {
            xls[1, row, width, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                    XlBordersIndex.xlEdgeBottom,
                                                                                                    XlBordersIndex.xlEdgeLeft,
                                                                                                    XlBordersIndex.xlEdgeRight, XlBordersIndex.xlInsideVertical);
        }

        private static void WriteRow(Xls xls, LogDelegate logDelegate, ReportParameters reportParameters, int row, int i, PTDocument ptDocument)
        {
            logDelegate(String.Format("Запись объекта: {0}", ptDocument.NomenclatureObject[Guids.Nomenclature.Name].Value + " " + ptDocument.NomenclatureObject[Guids.Nomenclature.Denotation].Value));

            xls[1, row].SetValue(i.ToString());
            xls[2, row].SetValue(ptDocument.NomenclatureObject[Guids.Nomenclature.Denotation].Value.ToString());
            xls[3, row].SetValue(ptDocument.NomenclatureObject[Guids.Nomenclature.Name].Value.ToString());
            xls[4, row].SetValue(ptDocument.NomenclatureObject[Guids.Nomenclature.Litera].Value.ToString());

            if (ptDocument.ObjectNotFound)
            {
                xls[5, row].SetValue("Объект не найден в БД ПИ");
                xls[1, row, 7, 1].Interior.Color = Color.LightYellow;
            }
            else
            {
                xls[5, row].SetValue(ptDocument.Name);
                xls[6, row].SetValue(ptDocument.Date.ToShortDateString());
                xls[7, row].SetValue(string.Join("\r\n", ptDocument.Executers));
                if (reportParameters.TermDeliveryOrder > ptDocument.Date.AddYears(3))
                {
                    xls[1, row, 7, 1].Interior.Color = Color.LightPink;
                }
            }

           

            InsertFrame(xls, row, 7);
        }


     


        //Рекурсивное получение дерева состава  
        public List<ReferenceObject> GetChildren(ReferenceObject rootObject, LogDelegate logDelegate)
        {
            var listObjects = new List<ReferenceObject>();
            if (rootObject.Class.IsInherit(Guids.Nomenclature.AssemblyType) || rootObject.Class.IsInherit(Guids.Nomenclature.ComplexType))
            {
                 listObjects.Add(rootObject);
            }
            var children = rootObject.Children;

            foreach (var child in children)
            {
                listObjects.AddRange(GetChildren(child, logDelegate));
            }
            return listObjects;
        }

        //Получение списка документов для периодических испытаний
        private List<PTDocument> GetData(ReportParameters reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            List<PTDocument> reportData = new List<PTDocument>();
            logDelegate("=== Загрузка данных ===");
            ReferenceInfo referenceNomenclatureInfo = ReferenceCatalog.FindReference(Guids.References.Nomenclature);
            Reference nomenclatureReference = referenceNomenclatureInfo.CreateReference();
            progressBar.PerformStep();
            nomenclatureReference.LoadSettings.AddParameters(Guids.Nomenclature.Name, Guids.Nomenclature.Denotation);

            //Дерево заказа
            List<ReferenceObject> orderTree = reportParameters.OrderName.Children.RecursiveLoad().Where(t => t.Class.IsInherit(Guids.Nomenclature.AssemblyType) || 
            t.Class.IsInherit(Guids.Nomenclature.ComplexType) || 
            t.Class.IsInherit(Guids.Nomenclature.KomplektType)).ToList();

            foreach (var obj in orderTree)
            {
                var linkedObjects = obj.GetObjects(Guids.Nomenclature.NomToPTDocumentLink);
                if (linkedObjects != null)
                {
                    foreach (var lo in linkedObjects)
                    {
                        PTDocument ptDocument = new PTDocument(lo, obj);
                        reportData.Add(ptDocument);
                    }

                    if (linkedObjects.Count == 0 && obj[Guids.Nomenclature.Code].Value.ToString().Trim() != string.Empty)
                    {
                        PTDocument ptDocument = new PTDocument(obj);
                        reportData.Add(ptDocument);
                    }
                }
                else
                {
                    if (obj[Guids.Nomenclature.Code].Value.ToString().Trim() != string.Empty)
                    {
                        PTDocument ptDocument = new PTDocument(obj);
                        reportData.Add(ptDocument);
                    }
                }
            }
            return reportData;
        }
    }
}
