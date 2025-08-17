using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using TFlex.Reporting;
using TFlex.DOCs.Model;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.SqlClient;
using TFlex.DOCs.Model.References.Nomenclature;

using System.Diagnostics;
using System.Activities.Statements;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Runtime.InteropServices;
using ReportHelpers;
using TFlex.DOCs.Model.References;

namespace RequirementWaybill
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
            selectionForm.reportContext = context;
            var nomenclatureReference = ServerGateway.Connection.ReferenceCatalog.Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83")).CreateReference();
            var rootObject = (NomenclatureObject)nomenclatureReference.Find(context.ObjectsInfo[0].ObjectId);
            selectionForm.Init(rootObject);

            if (selectionForm.ShowDialog(Form.ActiveForm) == DialogResult.OK)
            {
                RequirementWaybill.MakeReport(context, selectionForm.reportParameters);
            }
        }
    }

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);

    public class ReportParameters
    {
        public bool OtherItems { get; set; }
        public bool StandardItems { get; set; }
        public bool SeparateAssemblies { get; set; }
        public int ItemsCount { get; set; }
        public string ActionType { get; internal set; }
        public string Receiver { get; internal set; }
        public string Schet { get; internal set; }
        public string Department { get; internal set; }
    }

    public class TFDDocument
    {
        // Параметры объекта

        public int ObjectID;
        public int ParentID;
        public int Class;
        public double Amount;
        public string Zone;
        public int Position;
        public string Format;
        public string Naimenovanie;
        public string Denotation;
        public string Remarks;
        public int Level;
        public string Letter;

        public bool bold { get; internal set; }
    }

    public class RequirementWaybill : IDisposable
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            var report = new RequirementWaybill();
            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона

            Xls xls = new Xls();
            LogDelegate m_writeToLog;
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            var reportData = new List<TFDDocument>();
            List<TFDDocument> sortedReportData;

            RequirementWaybill report = new RequirementWaybill();

            int docID = report.GetDocsDocumentID(context);
            if (docID == -1)
            {
                MessageBox.Show("Выберите объект для формирования отчета");
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);

                m_form.setStage(IndicationForm.Stages.DataAcquisition);
                m_writeToLog("=== Получение данных ===");

                //Чтение данных для отчета                    
                reportData = report.ReadData(docID, context, reportParameters, m_writeToLog);
                m_form.setStage(IndicationForm.Stages.DataProcessing);

                m_writeToLog("=== Обработка данных ===");
                sortedReportData = SumDublicates(reportData);

                m_form.setStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeWaybillRequirement(xls, sortedReportData, reportParameters, m_writeToLog);

                m_form.progressBarParam(2);

                m_form.setStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");

                xls.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                try
                {
                    xls.Save();
                }
                catch { }
                xls.Visible = true;
                xls.ClearApplicationLink();

            }

            m_form.Dispose();
        }



        int GetDocsDocumentID(IReportGenerationContext context)
        {
            int documentID;

            // Выборка Id выделенного документа в T-FlexDocs 2010
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            return documentID;
        }

        public List<TFDDocument> ReadData(int documentID, IReportGenerationContext context, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            List<TFDDocument> objects = new List<TFDDocument>();

            var nomenclatureReference = ServerGateway.Connection.ReferenceCatalog.Find(new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83")).CreateReference();
            var rootObject = (NomenclatureObject)nomenclatureReference.Find(documentID);

            var hierarchyLinks = rootObject.Children.RecursiveLoadHierarchyLinks();
            var hierarchy = new Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>>();
            var objectsDictionary = new Dictionary<int, NomenclatureObject>();

            objectsDictionary[rootObject.SystemFields.Id] = rootObject;

            foreach (var hierarchyLink in hierarchyLinks)
            {
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
            CalculateCount(rootObject.SystemFields.Id, 1, hierarchy, countObjects);

            var objectTypes = new List<Guid>();
            if (reportParameters.OtherItems)
            {
                objectTypes.Add(OtehrItemGuid);
            }

            if (reportParameters.StandardItems)
            {
                objectTypes.Add(StandardItemGuid);
            }

            if (reportParameters.SeparateAssemblies)
            {

                var rootHierarchy = hierarchy[rootObject.SystemFields.Id];
                foreach (var nomObj in rootHierarchy.Values.Select(t => (NomenclatureObject)t.ChildObject)
                        .Where(t => objectTypes.Any(t1 => t.Class.IsInherit(t1)))
                        .OrderBy(t => t.Denotation.GetString()).ThenBy(t => t.Name.GetString()))
                {
                    var doc = new TFDDocument();
                    doc.ObjectID = nomObj.SystemFields.Id;
                    doc.Denotation = nomObj.Denotation;
                    doc.Naimenovanie = nomObj.Name;
                    // Округление результата до целого в большую сторону 
                    doc.Amount = Math.Ceiling(countObjects[nomObj.SystemFields.Id] * reportParameters.ItemsCount);
                    objects.Add(doc);
                }

                foreach (var subassembly in rootHierarchy.Values
                        .Where(t => t.ChildObject.Class.IsInherit(AssemblyGuid))
                        .OrderBy(t => ((NomenclatureObject)t.ChildObject).Denotation.GetString())
                        .ThenBy(t => ((NomenclatureObject)t.ChildObject).Name.GetString()))
                {
                    AddAssembly(subassembly, hierarchy, objects, objectTypes, reportParameters.ItemsCount);
                }

            }
            else
            {
                foreach (var nomObj in objectsDictionary.Values.Where(t => objectTypes.Any(t1 => t.Class.IsInherit(t1)))
                    .OrderBy(t => t.Denotation.GetString()).ThenBy(t => t.Name.GetString()))
                {
                    var doc = new TFDDocument();
                    doc.ObjectID = nomObj.SystemFields.Id;
                    doc.Denotation = nomObj.Denotation;
                    doc.Naimenovanie = nomObj.Name;
                    // Округление результата до целого в большую сторону 
                    doc.Amount = Math.Ceiling(countObjects[nomObj.SystemFields.Id] * reportParameters.ItemsCount);
                    objects.Add(doc);
                }
            }

            return objects;
        }

        static Guid OtehrItemGuid = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
        static Guid StandardItemGuid = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
        static Guid AssemblyGuid = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
        static Guid NomenclatureHierarchyCount = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");

        private void AddAssembly(ComplexHierarchyLink subassembly, Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>> hierarchy,
            List<TFDDocument> objects, List<Guid> objectTypes, double itemsCount)
        {
            if (!HasSubItems(subassembly, hierarchy, objectTypes, true)) return;

            var subAssemblyHierarchy = hierarchy[subassembly.ChildObjectId];

            if (HasSubItems(subassembly, hierarchy, objectTypes, false))
            {
                {
                    var doc = new TFDDocument();
                    doc.bold = true;
                    doc.ObjectID = subassembly.ChildObjectId;
                    doc.Denotation = ((NomenclatureObject)subassembly.ChildObject).Denotation;
                    doc.Naimenovanie = ((NomenclatureObject)subassembly.ChildObject).Name;
                    // Округление результата до целого в большую сторону 
                    doc.Amount = Math.Ceiling(subassembly[NomenclatureHierarchyCount].GetDouble() * itemsCount);
                    objects.Add(doc);
                }

                foreach (var nomObj in subAssemblyHierarchy.Values
                        .Where(t => objectTypes.Any(t1 => t.ChildObject.Class.IsInherit(t1)))
                        .OrderBy(t => ((NomenclatureObject)t.ChildObject).Denotation.GetString())
                        .ThenBy(t => ((NomenclatureObject)t.ChildObject).Name.GetString()))
                {
                    var doc = new TFDDocument();
                    doc.ObjectID = nomObj.ChildObjectId;
                    doc.Denotation = ((NomenclatureObject)nomObj.ChildObject).Denotation;
                    doc.Naimenovanie = ((NomenclatureObject)nomObj.ChildObject).Name;
                    // Округление результата до целого в большую сторону 
                    doc.Amount = Math.Ceiling(nomObj[NomenclatureHierarchyCount].GetDouble() * itemsCount *
                        subassembly[NomenclatureHierarchyCount].GetDouble());
                    objects.Add(doc);
                }
            }

            foreach (var subassembly1 in subAssemblyHierarchy.Values
                .Where(t => t.ChildObject.Class.IsInherit(AssemblyGuid))
                .OrderBy(t => ((NomenclatureObject)t.ChildObject).Denotation.GetString())
                .ThenBy(t => ((NomenclatureObject)t.ChildObject).Name.GetString()))
            {
                AddAssembly(subassembly1, hierarchy, objects, objectTypes, itemsCount * subassembly[NomenclatureHierarchyCount].GetDouble());
            }
        }

        private bool HasSubItems(ComplexHierarchyLink subassembly, Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>> hierarchy,
            List<Guid> objectTypes, bool recursive)
        {
            var subAssemblyHierarchy = hierarchy[subassembly.ChildObjectId];

            if (subAssemblyHierarchy.Values.Any(t => objectTypes.Any(t1 => t.ChildObject.Class.IsInherit(t1))))
            {
                return true;
            }

            if (recursive)
            {
                foreach (var subassembly1 in subAssemblyHierarchy.Values.Where(t => t.ChildObject.Class.IsInherit(AssemblyGuid)))
                {
                    var result = HasSubItems(subassembly1, hierarchy, objectTypes, recursive);
                    if (result) return true;
                }
            }


            return false;
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
                    var linkCount = link[NomenclatureHierarchyCount].GetDouble() * count;
                    CalculateCount(link.ChildObjectId, linkCount, hierarchy, countObjects);
                }
            }
        }

        /*
        // Чтение данных для отчета
        public List<TFDDocument> ReadData(int documentID, IReportGenerationContext context, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            // соединяемся с БД T-FLEX DOCs 2012
            SqlConnection conn;
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            sqlConStringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";
            // string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=S2"; //SRV1";
            conn = new SqlConnection(sqlConStringBuilder.ToString());
            conn.Open();

            List<TFDDocument> objects = new List<TFDDocument>();


            #region запрос для отчета с разузловкой
            //запрос для отчета с разузловкой
            string getDocTreeCommandText = String.Format(@"
                                  declare @docid INT
                                DECLARE @level int
                                declare @insertcount int
                                set @docid={0}
                                set @insertcount = 0
                                SET @level=0

                                IF OBJECT_ID('tempdb..#TmpR')is NULL
                                CREATE TABLE #TmpR ( 
                                                        s_ObjectID INT,
                                                        [level] INT,
                                                        Denotation NVARCHAR(255),
                                                        Name NVARCHAR(255),
                                                        Amount FLOAT,
                                                        s_ClassID INT,
                                                        TotalCount FLOAT)
                                ELSE DELETE FROM #TmpR

                                INSERT INTO #TmpR
                                SELECT n.s_ObjectID,0,n.Denotation,n.Name,0,0,1
                                FROM Nomenclature n
                                WHERE n.s_ObjectID = @docid
                                      AND n.s_Deleted = 0 
                                      AND n.s_ActualVersion = 1 

                                WHILE 1=1
                                BEGIN

                                  INSERT INTO #TmpR
                                  SELECT
                                         nh.s_ObjectID,
                                         @level+1,      
                                         n.Denotation,
                                         n.Name,
                                         nh.Amount,
                                         n.s_ClassID,
                                         #TmpR.TotalCount*nh.Amount
                                  FROM NomenclatureHierarchy nh INNER JOIN #TmpR ON nh.s_ParentID=#TmpR.s_ObjectID 

                                       INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                  WHERE
                                      #TmpR.[level]=@level
                                      AND nh.s_ActualVersion = 1
                                      AND nh.s_Deleted = 0
                                      AND n.s_ActualVersion = 1
                                      AND n.s_Deleted = 0

                                  SET @insertcount = @@ROWCOUNT 
                                  SET @level = @level + 1 

                                  IF @insertcount = 0 
                                  GOTO end1

                                END
                                end1:


                                SELECT s_ObjectID,
                                       Denotation, 
                                       Name, 
                                       TotalCount                          
                                FROM #TmpR
                                Where s_ClassID = 566
                                GROUP BY  Name,Denotation,s_ObjectID, TotalCount", documentID);

            SqlCommand docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
            docTreeCommand.CommandTimeout = 0;
            SqlDataReader reader = docTreeCommand.ExecuteReader();

            TFDDocument doc;

            while (reader.Read())
            {
                doc = new TFDDocument();
                doc.ObjectID = GetSqlInt32(reader, 0);
                doc.Denotation = GetSqlString(reader, 1).Trim();
                doc.Naimenovanie = GetSqlString(reader, 2).Trim();
                // Округление результата до целого в большую сторону 
                doc.Amount = Math.Ceiling(GetSqlDouble(reader, 3));


                objects.Add(doc);
                logDelegate(String.Format("Добавлен объект: {0} {1}", doc.Denotation, doc.Naimenovanie));
            }
            m_form.progressBarParam(2);
            reader.Close();
            #endregion

            return objects;
        }
        */
        private string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }

        private int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        private double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }




        // Запись в ячейку параметра объекта
        public int InsertCell(int col, int row, string value, Xls xls)
        {
            xls[col, row].SetValue(value);


            //for (int i = 2; i < 13; i++)
            //{
            //    xls[i, row].Interior.Color = Color.White.ToArgb();
            //    xls[i, row].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            //}
            col++;

            return col;
        }

        // Запись в ячейку параметра объекта
        public int InsertCell(int col, int row, double value, Xls xls)
        {

            if (value != 0)
                xls[col, row].SetValue(value);
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            return col;
        }

        public int countStr = 0;

        public int InsertObject(ReportParameters reportParameters, TFDDocument doc, int row, Xls xls, LogDelegate logDelegate, TextFormatter _textFormatter)
        {
            logDelegate(String.Format("Запись объекта: {0} {1}", doc.Denotation, doc.Naimenovanie));
            int col = 4;
            string strWrap = string.Empty;

            string[] strNameDen = _textFormatter.Wrap((doc.Naimenovanie + " " + doc.Denotation), 42f);

            for (int i = 0; i < strNameDen.Count() - 1; i++)
            {
                strWrap += strNameDen[i] + "\n";
                countStr++;
            }
            strWrap += strNameDen[strNameDen.Count() - 1];
            countStr++;

            col = InsertCell(col, row, strWrap, xls);
            col = 6;
            col = InsertCell(col, row, "шт", xls);
            col = 8;
            col = InsertCell(col, row, doc.Amount, xls);
            xls[2, row, 11, 1].Interior.Color = Color.White.ToArgb();
            xls[2, row, 11, 1].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium);
            xls[2, row, 11, 1].Font.Size = 10;
            xls[2, row, 11, 1].EntireRow.AutoFit();
            xls[2, row, 6, 1].HorizontalAlignment = XlHAlign.xlHAlignLeft;

            if (doc.bold)
            {
                xls[2, row, 11, 1].Font.Bold = true;
                xls[2, row, 3, 1].Font.Underline = true;
            }
            row++;

            return row;
        }

        public void MakeWaybillRequirement(Xls xls, List<TFDDocument> reportData, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            System.Drawing.Font font = new System.Drawing.Font("Arial Cyr", 2.3f, GraphicsUnit.Millimeter);
            TextFormatter _textFormatter = new TextFormatter(font);
            xls[5, 12, 6, 1].NumberFormat = "@";
            xls[5, 12].SetValue(reportParameters.Department);
            xls[7, 12].SetValue(reportParameters.ActionType);
            xls[8, 12].SetValue(reportParameters.Receiver);
            xls[10, 12].SetValue(reportParameters.Schet);

            int row = 19;
            string elementHeader = string.Empty;
            int countWrap = 0;

            var pageEnds = new List<int>();

            foreach (TFDDocument doc in reportData)
            {
                countWrap = _textFormatter.Wrap((doc.Naimenovanie + " " + doc.Denotation), 42f).Count();
                row = InsertObject(reportParameters, doc, row, xls, logDelegate, _textFormatter);

                if ((countStr > 34) && doc != reportData[reportData.Count - 1])
                {
                    xls[1, row].Cells.RowHeight = 15;
                    row++;

                    xls[2, row].SetValue("Отпустил\n");
                    xls[2, row].Font.Size = 7;
                    xls[3, row].SetValue("____________\n  должность");
                    xls[4, row].SetValue("____________\n  подпись");
                    xls[5, row, 2, 1].Merge();
                    xls[5, row].SetValue("_________________\nрасшифровка подписи");

                    xls[2, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[3, row, 3, 1].Font.Size = 5;

                    xls[8, row].SetValue("Получил\n");
                    xls[8, row].Font.Size = 7;
                    xls[9, row].SetValue("____________\n  должность");
                    xls[10, row].SetValue("____________\n  подпись");
                    xls[11, row, 2, 1].Merge();
                    xls[11, row].SetValue("_________________\nрасшифровка подписи");

                    xls[8, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    xls[9, row, 3, 1].Font.Size = 5;

                    pageEnds.Add(row);

                    xls[2, 2, 11, 17].Copy(Type.Missing);
                    row += 1;
                    xls[2, row, 11, 17].PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    xls[12, row + 8].Cells.RowHeight = 33;
                    xls[6, row + 14].Cells.RowHeight = 26.25;
                    row = row + 17;
                    countStr = 0;
                }

            }

            xls[10, row].SetValue("Итого:");
            xls[10, row].Font.Size = 8;
            xls[10, row].HorizontalAlignment = XlHAlign.xlHAlignRight;

            row++;
            row++;

            xls[2, row].SetValue("Отпустил\n");
            xls[2, row].Font.Size = 7;
            xls[3, row].SetValue("____________\n  должность");
            xls[4, row].SetValue("____________\n  подпись");
            xls[5, row, 2, 1].Merge();
            xls[5, row].SetValue("_________________\nрасшифровка подписи");

            xls[2, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[3, row, 3, 1].Font.Size = 5;

            xls[8, row].SetValue("Получил\n");
            xls[8, row].Font.Size = 7;
            xls[9, row].SetValue("____________\n  должность");
            xls[10, row].SetValue("____________\n  подпись");
            xls[11, row, 2, 1].Merge();
            xls[11, row].SetValue("_________________\nрасшифровка подписи");

            xls[8, row, 4, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[9, row, 3, 1].Font.Size = 5;

            pageEnds.Add(row);

            //MessageBox.Show(string.Join("\r\n", listEnds));

            xls[12, row + 8].Cells.RowHeight = 33;
            xls[6, row + 14].Cells.RowHeight = 26.25;

            PageSetup pageSetup = xls.Worksheet.PageSetup;
            pageSetup.PrintArea = Xls.Address(2, 2, 11, row - 1);
            //pageSetup.Zoom = false;
            pageSetup.FitToPagesWide = 1;
            pageSetup.FitToPagesTall = 10000;

            xls.Application.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
            xls.Application.ReferenceStyle = XlReferenceStyle.xlA1;

            xls.Worksheet.ResetAllPageBreaks();

            int pageIndex = 0;
            foreach (var page in pageEnds)
            {
                pageIndex++;

                if (xls.Worksheet.HPageBreaks.Count < pageIndex)
                {
                    try
                    {
                        xls.Worksheet.HPageBreaks.Add(xls[2, page + 1]);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Add: " + ex);
                    }
                }
                else
                {
                    try
                    {
                        xls.Worksheet.HPageBreaks[pageIndex].Location = xls[2, page + 1];
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Change: " + ex + "\r\n\r\n" + pageIndex + " " + Xls.Address(2, page + 1));
                    }
                }
            }
            xls.Application.ActiveWindow.View = XlWindowView.xlNormalView;

            //try
            //{
            //    if (xls.Worksheet.VPageBreaks.Count > 0)
            //        xls.Worksheet.VPageBreaks[1].DragOff(XlDirection.xlToRight, 1);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(
            //        "Не удалось установить вертикальную границу страниц!\r\n" +
            //        "xls.Worksheet.VPageBreaks: " + xls.Worksheet.VPageBreaks.Count + "\r\n" +
            //         ex,
            //        "Ошибка",
            //        MessageBoxButtons.OK,
            //        MessageBoxIcon.Error);
            //}
        }


        // Сложение количеств повторяющихся элементов отсортированной коллекции
        static private List<TFDDocument> SumDublicates(List<TFDDocument> data)
        {

            //for (int mainID = 0; mainID < data.Count; mainID++)
            //{
            //    for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
            //    {
            //        //если документы имеют одинаковые обозначения и наименования - удаляем повторку
            //        if (data[mainID].ObjectID == data[slaveID].ObjectID)
            //        {

            //            data[mainID].Amount += data[slaveID].Amount;
            //            data.RemoveAt(slaveID);
            //            slaveID--;
            //        }
            //    }
            //}
            return data;
        }

    }

    public class TextFormatter
    {
        System.Drawing.Font _font;

        int _displayResolutionX;

        public TextFormatter(System.Drawing.Font font)
        {
            _font = font;
            _displayResolutionX = GetDisplayResolutionX();
        }

        [DllImport("User32.dll")]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        public enum DeviceCap
        {
            #region
            /// <summary>
            /// Device driver version
            /// </summary>
            DRIVERVERSION = 0,
            /// <summary>
            /// Device classification
            /// </summary>
            TECHNOLOGY = 2,
            /// <summary>
            /// Horizontal size in millimeters
            /// </summary>
            HORZSIZE = 4,
            /// <summary>
            /// Vertical size in millimeters
            /// </summary>
            VERTSIZE = 6,
            /// <summary>
            /// Horizontal width in pixels
            /// </summary>
            HORZRES = 8,
            /// <summary>
            /// Vertical height in pixels
            /// </summary>
            VERTRES = 10,
            /// <summary>
            /// Number of bits per pixel
            /// </summary>
            BITSPIXEL = 12,
            /// <summary>
            /// Number of planes
            /// </summary>
            PLANES = 14,
            /// <summary>
            /// Number of brushes the device has
            /// </summary>
            NUMBRUSHES = 16,
            /// <summary>
            /// Number of pens the device has
            /// </summary>
            NUMPENS = 18,
            /// <summary>
            /// Number of markers the device has
            /// </summary>
            NUMMARKERS = 20,
            /// <summary>
            /// Number of fonts the device has
            /// </summary>
            NUMFONTS = 22,
            /// <summary>
            /// Number of colors the device supports
            /// </summary>
            NUMCOLORS = 24,
            /// <summary>
            /// Size required for device descriptor
            /// </summary>
            PDEVICESIZE = 26,
            /// <summary>
            /// Curve capabilities
            /// </summary>
            CURVECAPS = 28,
            /// <summary>
            /// Line capabilities
            /// </summary>
            LINECAPS = 30,
            /// <summary>
            /// Polygonal capabilities
            /// </summary>
            POLYGONALCAPS = 32,
            /// <summary>
            /// Text capabilities
            /// </summary>
            TEXTCAPS = 34,
            /// <summary>
            /// Clipping capabilities
            /// </summary>
            CLIPCAPS = 36,
            /// <summary>
            /// Bitblt capabilities
            /// </summary>
            RASTERCAPS = 38,
            /// <summary>
            /// Length of the X leg
            /// </summary>
            ASPECTX = 40,
            /// <summary>
            /// Length of the Y leg
            /// </summary>
            ASPECTY = 42,
            /// <summary>
            /// Length of the hypotenuse
            /// </summary>
            ASPECTXY = 44,
            /// <summary>
            /// Shading and Blending caps
            /// </summary>
            SHADEBLENDCAPS = 45,

            /// <summary>
            /// Logical pixels inch in X
            /// </summary>
            LOGPIXELSX = 88,
            /// <summary>
            /// Logical pixels inch in Y
            /// </summary>
            LOGPIXELSY = 90,

            /// <summary>
            /// Number of entries in physical palette
            /// </summary>
            SIZEPALETTE = 104,
            /// <summary>
            /// Number of reserved entries in palette
            /// </summary>
            NUMRESERVED = 106,
            /// <summary>
            /// Actual color resolution
            /// </summary>
            COLORRES = 108,

            // Printing related DeviceCaps. These replace the appropriate Escapes
            /// <summary>
            /// Physical Width in device units
            /// </summary>
            PHYSICALWIDTH = 110,
            /// <summary>
            /// Physical Height in device units
            /// </summary>
            PHYSICALHEIGHT = 111,
            /// <summary>
            /// Physical Printable Area x margin
            /// </summary>
            PHYSICALOFFSETX = 112,
            /// <summary>
            /// Physical Printable Area y margin
            /// </summary>
            PHYSICALOFFSETY = 113,
            /// <summary>
            /// Scaling factor x
            /// </summary>
            SCALINGFACTORX = 114,
            /// <summary>
            /// Scaling factor y
            /// </summary>
            SCALINGFACTORY = 115,

            /// <summary>
            /// Current vertical refresh rate of the display device (for displays only) in Hz
            /// </summary>
            VREFRESH = 116,
            /// <summary>
            /// Horizontal width of entire desktop in pixels
            /// </summary>
            DESKTOPVERTRES = 117,
            /// <summary>
            /// Vertical height of entire desktop in pixels
            /// </summary>
            DESKTOPHORZRES = 118,
            /// <summary>
            /// Preferred blt alignment
            /// </summary>
            BLTALIGNMENT = 119
            #endregion
        }

        int GetDisplayResolutionX()
        {
            IntPtr p = GetDC(new IntPtr(0));
            int res = GetDeviceCaps(p, (int)DeviceCap.LOGPIXELSX);
            return res;
        }

        /// Коэффициент, учитывающий различие между шириной строки,
        /// определяемой в этой программе и шириной строки, получаемой в CAD
        static readonly float widthFactor = 1.5f;

        public double GetTextWidth(string text, System.Drawing.Font font)
        {
            Size size = TextRenderer.MeasureText(text, font);
            double width = 25.4f * widthFactor * (double)size.Width / (double)_displayResolutionX;
            return width;
        }


        public string[] Wrap(string text, double maxWidth)
        {
            return Wrap(text, _font, maxWidth);
        }


        /// Разбивка на строки по символам перевода строки \n
        string[] Wrap(string text, System.Drawing.Font font, double maxWidth)
        {
            string[] lines = text.Split(new string[] { @"\n" }, StringSplitOptions.None);


            List<string> wrappedLines = new List<string>();
            foreach (string line in lines)
            {
                string[] wrappedBySpaces = WrapBySyllables(line, font, maxWidth);
                wrappedLines.AddRange(wrappedBySpaces);
            }
            return wrappedLines.ToArray();
        }

        /// Выделение слога - последовательности символов с завершающим пробелом
        /// или символом мягкого переноса. Символ мягкого переноса (если присутствует)
        /// выносится в отдельную группу
        static readonly Regex _syllableRegex = new Regex(
            @"(?<syllable>((ГОСТ|ОСТ)[\x20\u00AD]+)?[^\x20\u00AD]+)(?<softHyphen>\u00AD+)?",
            RegexOptions.Compiled);

        /// Разбивка на строки по слогам (разделители - пробелы и знаки мягкого переноса)
        string[] WrapBySyllables(string text, System.Drawing.Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            MatchCollection mc = _syllableRegex.Matches(text);
            string currentLine = "";
            bool currentLineEndsWithSoftHyphen = false;
            foreach (Match match in mc)
            {
                string syllable = match.Groups["syllable"].Value;
                bool endsBySoftHyphen = match.Groups["softHyphen"].Success;

                string candidateLine;
                if (currentLine.Length > 0)
                {
                    if (currentLineEndsWithSoftHyphen)
                        candidateLine = currentLine + syllable;
                    else
                        candidateLine = currentLine + " " + syllable;
                }
                else
                    candidateLine = syllable;
                double candidateLineWidth = GetTextWidth(candidateLine, font);

                if (candidateLineWidth > maxWidth)
                {
                    // Перенос очередной последовательности символов на новую строку

                    if (currentLine.Length > 0)
                    {
                        if (currentLineEndsWithSoftHyphen)
                            lines.Add(currentLine + "-");
                        else
                            lines.Add(currentLine);
                    }
                    currentLine = "";
                    currentLineEndsWithSoftHyphen = false;

                    candidateLine = syllable;
                    candidateLineWidth = GetTextWidth(candidateLine, font);
                }

                if (candidateLineWidth > maxWidth)
                {
                    // Разбивка строки-кандидата по разделителям-не буквам

                    if (currentLine.Length > 0)
                    {
                        if (currentLineEndsWithSoftHyphen)
                            lines.Add(currentLine + "-");
                        else
                            lines.Add(currentLine);
                    }
                    currentLine = "";
                    currentLineEndsWithSoftHyphen = false;

                    string[] candidateLines = WrapByAlphaAndDelimiter(candidateLine, font, maxWidth);
                    foreach (string s in candidateLines)
                        lines.Add(s);
                }
                else
                {
                    // строка-кандидат не превышает максимально допустимую ширину
                    currentLine = candidateLine;
                    currentLineEndsWithSoftHyphen = endsBySoftHyphen;
                }
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }

        static readonly Regex _alphaAndDelimiterRegex = new Regex(
            @"[\d()*\p{Ll}\p{Lu}]+[^\p{Ll}\p{Lu}]*", RegexOptions.Compiled);

        /// Разбивка на строки по разделителям-не буквам
        string[] WrapByAlphaAndDelimiter(string text, System.Drawing.Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            MatchCollection mc = _alphaAndDelimiterRegex.Matches(text);
            string currentLine = "";
            foreach (Match match in mc)
            {
                string candidateLine = currentLine + match.Value;
                double candidateLineWidth = GetTextWidth(candidateLine, font);

                if (candidateLineWidth > maxWidth)
                {
                    // Перенос очередной последовательности символов на новую строку

                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = "";

                    candidateLine = match.Value;
                    candidateLineWidth = GetTextWidth(candidateLine, font);
                }

                if (candidateLineWidth > maxWidth)
                {
                    // Жесткая разбивка новой последовательности символов

                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = "";

                    string[] candidateLines = WrapByCharacters(candidateLine, font, maxWidth);
                    foreach (string s in candidateLines)
                        lines.Add(s);
                }
                else
                    // строка-кандидат не превышает максимально допустимую ширину
                    currentLine = candidateLine;
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }

        /// Разбивка на строки по символам
        string[] WrapByCharacters(string text, System.Drawing.Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            string currentLine = "";
            for (int i = 0; i < text.Length; ++i)
            {
                string candidateLine = currentLine + text.Substring(i, 1);
                double candidateLineWidth = GetTextWidth(candidateLine, font);
                if (candidateLineWidth > maxWidth)
                {
                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = text.Substring(i, 1);
                }
                else
                    currentLine = candidateLine;
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }
    }

    internal class NaimenAndOboznachComparer : IComparer<TFDDocument>
    {
        public int Compare(TFDDocument ob1, TFDDocument ob2)
        {
            string designation1 = ob1.Denotation.Replace(" ", "");
            string designation2 = ob2.Denotation.Replace(" ", "");
            int ob = String.Compare(designation1, designation2);
            if (ob != 0)
                return ob;
            else
                return String.Compare(ob1.Naimenovanie, ob2.Naimenovanie);
        }
    }
}

