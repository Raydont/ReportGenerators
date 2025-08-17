using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSO.NetTDocs.Excel;
using Microsoft.Office.Interop.Excel;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model;


namespace SearchObjectsNotLinkFilesReport
{
    class MakeReport
    {
        public static Guid UserReferenceGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        public static Guid Department05Guid = new Guid("b9752a5d-fde8-4aec-ae0b-f51978960f3c");
        public static Guid UserTypeGuid = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");

            public IndicationForm m_form = new IndicationForm();

            public void Dispose()
            {
            }

            //Создание экземпляра формы выбора параметров формирования отчета
            public SelectionForm selectionForm = new SelectionForm();

            public void Make(IReportGenerationContext context)
            {
                TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    context.CopyTemplateFile();    // Создаем копию шаблона
                    m_form.Visible = true;
                    LogDelegate m_writeToLog;
                    m_writeToLog = new LogDelegate(m_form.writeToLog);

                    m_form.setStage(IndicationForm.Stages.Initialization);
                    m_writeToLog("=== Инициализация ===");

                    var reportData = new List<ReportParameters>();
                    var sortedReportData = new List<ReportParameters>();
                    var listObjectId = new List<int>();

                    
                    try
                    {
                        m_form.setStage(IndicationForm.Stages.DataAcquisition);
                        m_writeToLog("=== Получение данных ===");

                        listObjectId.AddRange(selectionForm.attributeParameters.ListOrderId);

                        var attributeParameters = new AttributeParametrsClass();
                        attributeParameters = selectionForm.attributeParameters;

                        // получение описания справочника          
                        ReferenceInfo referenceInfoUser = ReferenceCatalog.FindReference(UserReferenceGuid);
                        // создание объекта для работы с данными
                        Reference referenceUser = referenceInfoUser.CreateReference();


                        var dept05 = referenceUser.Find(Department05Guid);

                        GetAllUser(dept05);

                        var reportItems = GetData(context, listObjectId, reportData, m_writeToLog, selectionForm.attributeParameters);

                        attributeParameters.typeReport = !selectionForm.attributeParameters.typeReport;
                        var reportItemsOther = GetData(context, listObjectId, reportData, m_writeToLog, attributeParameters);

                        // MessageBox.Show("После сортировки " + reportItems.Count.ToString());

                        //   return;
                        var excelApp = new ExcelReader();
                        excelApp.NewDocument(context.ReportFilePath, true);
                        m_form.setStage(IndicationForm.Stages.DataProcessing);
                        m_writeToLog("=== Обработка данных ===");

                        m_form.setStage(IndicationForm.Stages.ReportGenerating);
                        m_writeToLog("=== Формирование отчета ===");

                        //Формирование отчета
                        MakeSelectionByTypeReport(excelApp, reportItems,reportItemsOther, m_writeToLog, attributeParameters);

                        m_form.progressBarParam(2);

                        m_form.setStage(IndicationForm.Stages.Done);
                        m_writeToLog("=== Завершение работы ===");
                        System.Diagnostics.Process.Start(context.ReportFilePath);

                    }
                    catch (Exception e)
                    {
                        string message = String.Format("ОШИБКА: {0}", e.ToString());
                        System.Windows.Forms.MessageBox.Show(message, "Ошибка",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    m_form.Dispose();
                }
               
                
            }

            private List<ReportItem> GetData(IReportGenerationContext context, List<int> listObjectId, List<ReportParameters> reportData, LogDelegate m_writeToLog, AttributeParametrsClass attributeParameters)
            {
                foreach (var id in listObjectId)
                {
                    reportData.AddRange(ReadData(id, context, m_writeToLog, attributeParameters));
                }
                // MessageBox.Show("До сортировки " + reportData.Count.ToString());
                FindDublicates(reportData);
                reportData = reportData.OrderBy(t => t.Type).ThenBy(t => t.Denotation).ToList();

                var reportItems = new List<ReportItem>();

                ReportItem lastItem = null;

                // формирование списка на вывод с учетом группировки по исполнениям
                foreach (var item in reportData)
                {
                    if (lastItem == null) // первый элемент отчета
                    {
                        lastItem = new ReportItem(item);
                        reportItems.Add(lastItem);
                    }
                    else
                    {
                        // если последний элемент отчета соответствует текущему, то добавляем исполнение
                        if (lastItem.IsValid(item))
                        {
                            lastItem.Add(item);
                        }
                        else
                        {
                            lastItem = new ReportItem(item);
                            reportItems.Add(lastItem);
                        }
                    }
                }


                reportItems.OrderBy(t => t.Type).ThenBy(t => t.denotation).ToList();
                return reportItems;
            }

        public List<int> listUserId = new List<int>();

            public void GetAllUser(ReferenceObject dept)
            {
                foreach (var userObj in dept.Children)
                {
                    if (userObj.Class.IsInherit(UserTypeGuid))
                    {
                        listUserId.Add(userObj.SystemFields.Id);
                    }
                    else
                    {
                        GetAllUser(userObj);
                    }

                }
            }

        // Чтение данных для отчета
            public List<ReportParameters> ReadData(int documentID, IReportGenerationContext context, LogDelegate logDelegate, AttributeParametrsClass attributeParameters)
            {
                var typeObjects = string.Empty;
                bool addType = false;
                if (attributeParameters.DocumentObject)
                {
                    typeObjects = "946";
                    addType = true;
                }
               

                if (selectionForm.attributeParameters.DetailObject)
                {
                    if (addType)
                    {
                        typeObjects += ",";
                        addType = false;
                    }
                    typeObjects += "499";
                    addType = true;
                }



                if (selectionForm.attributeParameters.ComplementObject)
                {
                    if (addType)
                    {
                        typeObjects += ",";
                    }
                    typeObjects += "501";
                }
                var stringUserId = string.Empty;

                foreach (var id in listUserId)
                {
                    stringUserId += id.ToString();
                    if (id != listUserId[listUserId.Count-1])
                    {
                        stringUserId += ",";
                    }
                }

               

                // соединяемся с БД T-FLEX DOCs 2012
                var sqlConStringBuilder = new SqlConnectionStringBuilder();
                sqlConStringBuilder.DataSource = "S2";
                sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
                sqlConStringBuilder.Password = "reportUser";
                sqlConStringBuilder.UserID = "reportUser";
                var conn = new SqlConnection(sqlConStringBuilder.ToString());
                conn.Open();
             
                var objects = new List<ReportParameters>();
                string getAllDetailAndDoc = String.Format(@"  SELECT DISTINCT tv.s_ObjectID,
                                                                                  tv.Denotation,
                                                                                  tv.Name,
                                                                                  c.Name, 
                                                                                 
                                                                                  u.FullName,
                                                                                  tv.Format
                                                                  FROM Nomenclature tv INNER JOIN Classes c ON (c.PK = tv.s_ClassID)
                                                                        INNER JOIN [Users] u ON (tv.s_AuthorID = u.s_ObjectID)
                                                                  WHERE tv.s_ObjectID not in (SELECT DISTINCT n.s_ObjectID
                                                                                              FROM Nomenclature n
                                                                                              INNER JOIN DocumentFiles df ON (n.s_LinkedObjectID = df.MasterID))
                                                                                              AND  tv.s_ClassID in (499,946) AND c.GroupID = 399 AND tv.Denotation not Like N'%ЛУ%' 
                                                                                              AND tv.Denotation not Like N'%УД%' AND tv.s_AuthorID not in(49)
                                                                                              AND tv.s_ActualVersion = 1 AND tv.s_Deleted = 0
                                                                   
                                                                   GROUP BY  c.Name,
                                                                             tv.Denotation, 
                                                                             tv.Name, 
                                                                             u.FullName,
                                                                             tv.s_ObjectID,         
                                                                             tv.Format
                                                                             
                                                                  ");

                string getDocTreeNotLinkWithDoc = String.Format(@"declare @docid INT
                                                                    DECLARE @level int
                                                                    declare @insertcount int
                                                                    set @docid = {0}
                                                                    set @insertcount = 0
                                                                    SET @level=0

                                                                    IF OBJECT_ID('tempdb..#TmpVspec')is NULL
                                                                    CREATE TABLE #TmpVspec (s_ObjectID INT,
                                                                                            [level] INT,
                                                                                            s_ParentID INT,
                                                                                            s_ClassID INT,
                                                                                            Denotation NVARCHAR(255),
                                                                                            Name NVARCHAR(255),
                                                                                            s_AuthorID INT,
                                                                                            Format NVARCHAR(255))
                                                                    ELSE DELETE
                                                                    FROM #TmpVspec

                                                                    INSERT INTO #TmpVspec
                                                                    SELECT n.s_ObjectID,0,0,n.s_ClassID,n.Denotation,n.Name, n.s_AuthorID,''
                                                                    FROM Nomenclature n
                                                                    WHERE n.s_ObjectID = @docid
                                                                          AND n.s_Deleted = 0 
                                                                          AND n.s_ActualVersion = 1 

                                                                    WHILE 1=1
                                                                    BEGIN

                                                                      INSERT INTO #TmpVspec 
                                                                      SELECT nh.s_ObjectID,
                                                                             @level+1,
                                                                             nh.s_ParentID,
                                                                             n.s_ClassID,
                                                                             n.Denotation,
                                                                             n.Name,
                                                                             n.s_AuthorID,
                                                                             n.Format
                                        
                                                                      FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                           
                                                                           INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                                                      WHERE
                                                                          #TmpVspec.[level]=@level
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

                                                                                 
                                                                  SELECT DISTINCT tv.s_ObjectID,
                                                                                  tv.Denotation,
                                                                                  tv.Name,
                                                                                  c.Name, 
                                                                                  tv.[level], 
                                                                                  u.FullName,
                                                                                  tv.Format
                                                                 FROM #TmpVspec tv INNER JOIN Classes c ON (c.PK = tv.s_ClassID)
                                                                        INNER JOIN [Users] u ON (tv.s_AuthorID = u.s_ObjectID)
                                                                 WHERE tv.s_ObjectID not in (SELECT DISTINCT temp.s_ObjectID
                                                                                              FROM #TmpVspec temp INNER JOIN Nomenclature n 
                                                                                              ON (n.s_ObjectID = temp.s_ObjectID) 
                                                                                              INNER JOIN DocumentFiles df ON (s_LinkedObjectID = df.MasterID))
                                                                   AND  tv.s_ClassID in ({1}) AND c.GroupID = 399 AND tv.Denotation not Like N'%ЛУ%' AND tv.Denotation not Like N'%УД%' 
                                                                   AND tv.s_AuthorID not in({2})
                                                                   
                                                                   AND  tv.Name not Like N'%КАСАК%' AND  tv.Name not Like N'%ВИУ%' AND  tv.Name not Like  N'%СМК%'
                                                                   AND  tv.Name not Like N'%БАЗОВ%'AND  tv.Name not Like N'%МОНИТОР%' 
                                                                   AND  tv.Name not Like N'%Описание применения%'
                                                                    GROUP BY u.FullName,
                                                                              c.Name,
                                                                             tv.Denotation, 
                                                                             tv.Name, 
                                                                             
                                                                             tv.s_ObjectID,
                                                                             tv.[level],
                                                                             tv.Format
                                                                   UNION ALL
                                                                   SELECT DISTINCT tv.s_ObjectID,
                                                                                   tv.Denotation,
                                                                                   tv.Name,
                                                                                   N'Комплекс',
                                                                                   tv.[level],
                                                                                   '',
                                                                                    tv.Format
                                                                  FROM #TmpVspec tv 
                                                                  WHERE tv.[level] = 0
                                                                
                                                                
                                                                  ", documentID, typeObjects, stringUserId);

                // Объекты с присоединенными файлами

                string getDocTreeLinkWithDoc = String.Format(@"declare @docid INT
                                                                    DECLARE @level int
                                                                    declare @insertcount int
                                                                    set @docid = {0}
                                                                    set @insertcount = 0
                                                                    SET @level=0

                                                                    IF OBJECT_ID('tempdb..#TmpVspec')is NULL
                                                                    CREATE TABLE #TmpVspec (s_ObjectID INT,
                                                                                            [level] INT,
                                                                                            s_ParentID INT,
                                                                                            s_ClassID INT,
                                                                                            Denotation NVARCHAR(255),
                                                                                            Name NVARCHAR(255),
                                                                                             s_AuthorID INT,
                                                                                              Format NVARCHAR(255))
                                                                    ELSE DELETE
                                                                    FROM #TmpVspec

                                                                    INSERT INTO #TmpVspec
                                                                    SELECT n.s_ObjectID,0,0,n.s_ClassID,n.Denotation,n.Name, n.s_AuthorID,''
                                                                    FROM Nomenclature n
                                                                    WHERE n.s_ObjectID = @docid
                                                                          AND n.s_Deleted = 0 
                                                                          AND n.s_ActualVersion = 1 

                                                                    WHILE 1=1
                                                                    BEGIN

                                                                      INSERT INTO #TmpVspec 
                                                                      SELECT nh.s_ObjectID,
                                                                             @level+1,
                                                                             nh.s_ParentID,
                                                                             n.s_ClassID,
                                                                             n.Denotation,
                                                                             n.Name,
                                                                             n.s_AuthorID,
                                                                             n.Format
                                                                      FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                           
                                                                           INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                                                      WHERE
                                                                          #TmpVspec.[level]=@level
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
                                                                                   SELECT DISTINCT tv.s_ObjectID,
                                                                                                              tv.Denotation,
                                                                                                              tv.Name,
                                                                                                              c.Name, 
                                                                                                              tv.[level],
                                                                                                              u.FullName,
                                                                                                              tv.Format
                                                                                              FROM #TmpVspec tv INNER JOIN Nomenclature n 
                                                                                              ON (n.s_ObjectID = tv.s_ObjectID)
                                                                                              INNER JOIN DocumentFiles df ON (s_LinkedObjectID = df.MasterID)
                                                                                              INNER JOIN Classes c ON (c.PK = tv.s_ClassID)
                                                                                              INNER JOIN [Users] u ON (tv.s_AuthorID = u.s_ObjectID)
                                                                   WHERE  tv.s_ClassID in ({1}) AND c.GroupID = 399   AND tv.s_AuthorID not in({2})
                                                                     AND  tv.Name not Like N'%КАСАК%' AND  tv.Name not Like N'%ВИУ%' AND  tv.Name not Like  N'%СМК%'
                                                                   AND  tv.Name not Like N'%БАЗОВ%'AND  tv.Name not Like N'%МОНИТОР%' 
                                                                   AND  tv.Name not Like N'%Описание применения%'
                                                                    GROUP BY c.Name,
                                                                             tv.Denotation, 
                                                                             tv.Name, 
                                                                             u.FullName,
                                                                              tv.s_ObjectID,
                                                                              tv.[level],
                                                                              tv.Format
                                                                   UNION ALL
                                                                    SELECT DISTINCT tv.s_ObjectID,
                                                                                   tv.Denotation,
                                                                                   tv.Name,
                                                                                   
                                                                                   N'Комплекс',
                                                                                   tv.[level],
                                                                                   '',
                                                                                    ''
                                                                  FROM #TmpVspec tv 
                                                                  WHERE tv.[level] = 0
                                                                
                                                                  ", documentID, typeObjects,stringUserId);
                SqlCommand docTreeCommand = null;
                if(attributeParameters.typeReport)
                    docTreeCommand = new SqlCommand(getDocTreeNotLinkWithDoc, conn);
                else
                    docTreeCommand = new SqlCommand(getDocTreeLinkWithDoc, conn);

                //docTreeCommand = new SqlCommand(getAllDetailAndDoc, conn);

                docTreeCommand.CommandTimeout = 0;
                SqlDataReader reader = docTreeCommand.ExecuteReader();
                ReportParameters doc;

                while (reader.Read())
                {
                    doc = new ReportParameters();
                    doc.Id = GetSqlInt32(reader, 0);
                    doc.Denotation = GetSqlString(reader, 1);
                    doc.Name = GetSqlString(reader, 2);
                    doc.Type = GetSqlString(reader, 3);
                    doc.Level = GetSqlInt32(reader, 4);
                    doc.Author = GetSqlString(reader, 5);
                    doc.Format = GetSqlString(reader, 6);
                    objects.Add(doc);
                    logDelegate(String.Format("Добавлен объект: {0} {1}", doc.Denotation, doc.Name));

                }
                m_form.progressBarParam(2);
                reader.Close();

                return objects;
            }



            public static string GetSqlString(SqlDataReader reader, int field)
            {
                if (reader.IsDBNull(field))
                    return String.Empty;

                return reader.GetString(field);
            }

            public static int GetSqlInt32(SqlDataReader reader, int field)
            {
                if (reader.IsDBNull(field))
                    return 0;

                return reader.GetInt32(field);
            }

            public static double GetSqlDouble(SqlDataReader reader, int field)
            {
                if (reader.IsDBNull(field))
                    return 0d;

                return reader.GetDouble(field);
            }

            // Запись в ячейку элемента заголовка таблицы
            public static int InsertHeader( string header, int col, ExcelReader excelApp)
            {
                
                    excelApp.ApplyStyle(2, col);
                    excelApp.setValueCell(2, col, header);
                    excelApp.Bold(2, col, true);
                    col++;
             
                return col;
            }

            // Запись заголовка раздела спецификации
            public static void InsertBomSectionHeader(int row, int col, string bomSectionHeader, ExcelReader excelApp)
            {
                string firstCell = excelApp.cellNumToAlphabet(row, 1);
                string lastCell = excelApp.cellNumToAlphabet(row, col);
                excelApp.ApplyStyle(firstCell, lastCell);
              
                
                excelApp.setValueCell(row, 1, bomSectionHeader);
                excelApp.Bold(row, 1, true);
                excelApp.Underline(row, 1, true);

                excelApp.MergeCells(firstCell, lastCell);

            }



            // Запись в ячейку параметра объекта
            public static int InsertCell( int col, int row, string value, ExcelReader excelApp)
            {
               

                    excelApp.ApplyStyle(row, col);
                    excelApp.setValueCell(row, col, value);
                    col++;
               
                return col;
            }

          

            // Запись в ячейку параметра объекта
            public static int InsertCell(bool isCheck, int col, int row, int value, ExcelReader excelApp)
            {
                if (isCheck)
                {
                    if (value != 0)
                    {
                        excelApp.ApplyStyle(row, col);
                        excelApp.setValueCell(row, col, value);
                    }
                    col++;
                }
                return col;
            }

            // Запись в ячейку параметра объекта
            public static int InsertCell(bool isCheck, int col, int row, double value, ExcelReader excelApp)
            {
                if (isCheck)
                {
                    if (value != 0)
                    {
                        excelApp.ApplyStyle(row, col);
                        excelApp.setValueCell(row, col, value);
                    }
                    col++;
                }
                return col;
            }

            private void FindDublicates(List<ReportParameters> data)
            {

                for (int mainID = 0; mainID < data.Count; mainID++)
                {
                    for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                    {
                        //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                        if (String.Equals(data[mainID].Name, data[slaveID].Name) &&
                            String.Equals(data[mainID].Denotation, data[slaveID].Denotation) &&
                            String.Equals(data[mainID].Type, data[slaveID].Type))
                        {
                            data.RemoveAt(slaveID);
                            slaveID--;
                        }
                    }
                }
            }

            public static string text(string[] textArray)
            {
                string str = string.Empty;
                if (textArray.Count() > 1)
                {
                    for (int i = 0; i < textArray.Count() - 1; i++)
                    {
                        if (textArray[i].Trim() != string.Empty)
                            str += textArray[i] + '\n';
                    }
                    str += textArray[textArray.Count() - 1];
                }
                else str = textArray[textArray.Count() - 1];

                return str;
            }

            public static int countStr = 0;
            public static int countWrap = 0;
            public static int countPage = 0;

            public static int InsertObject( ReportItem doc, int row, string typeName, ExcelReader excelApp, LogDelegate logDelegate)
            {


                int col = 1;
                if (doc.Type == typeName)
                {
                    logDelegate(String.Format("Запись объекта: {0} {1}", doc.Denotation, doc.Naimenovanie));

                    string denotation = doc.Denotation;
                    string name = doc.Naimenovanie;

                    col = InsertCell(col, row, doc.Format, excelApp);
                    col = InsertCell( col, row, denotation.Trim(), excelApp);
                    col = InsertCell( col, row, name.Trim(), excelApp);
                    col = InsertCell( col, row, doc.Author, excelApp);
                   



                    row++;
                }
                return row;
            }

            // Группировка элементов по типам с сортировкой внутри групп
            public static List<ReportParameters> SortData(List<ReportParameters> reportData)
            {
                var sortedReportData = new List<ReportParameters>();

                sortedReportData.AddRange(SortOneTypeObjects(reportData, "Документ"));
                sortedReportData.AddRange(SortOneTypeObjects(reportData, "Комплект"));
                sortedReportData.AddRange(SortOneTypeObjects(reportData, "Деталь"));
              
                return sortedReportData;
            }

            //Сортировка внутри группы объектов одного и того же типа
            public static List<ReportParameters> SortOneTypeObjects(List<ReportParameters> reportData, string typeName)
            {
                var oneTypeObjects = new List<ReportParameters>();
                foreach (ReportParameters doc in reportData)
                    if (doc.Type == typeName)
                        oneTypeObjects.Add(doc);

                oneTypeObjects.Sort(new NaimenAndOboznachComparer()); //сортируем записи по обозначению и наименованию
                return oneTypeObjects;
            }

            public static void MakeSelectionByTypeReport(ExcelReader excelApp, List<ReportItem> reportData, List<ReportItem> reportDataOther, LogDelegate logDelegate , AttributeParametrsClass attributeParametrs)
            {
                int col = 1;


                col = InsertHeader("Формат", col, excelApp);
                col = InsertHeader( "Обозначение", col, excelApp);
                col = InsertHeader("Наименование", col, excelApp);
                col = InsertHeader( "Автор", col, excelApp);
              
                excelApp.BorderLineCells(1, 1, 2, col - 1,
                         XlBordersIndex.xlEdgeTop,
                         XlBordersIndex.xlEdgeBottom,
                         XlBordersIndex.xlEdgeLeft,
                         XlBordersIndex.xlEdgeRight,
                         XlBordersIndex.xlInsideVertical);

                int colCount = col;

                var index = -1;
                var rootObjects = new List<ReportItem>();
                string headerString = string.Empty;
                if(attributeParametrs.typeReport)
                    headerString ="Поиск объектов не связанных с файлами в заказах: ";
                else
                    headerString = "Поиск объектов связанных с файлами в заказах: ";

                foreach (ReportItem doc in reportData)
                {
                    index++;
                    if (doc.Level == 0)
                    {
                        headerString += " " + doc.Denotation + ",";
                        rootObjects.Add(doc);
                    }
                }
                countStr = 3;

                excelApp.setValueCell(1, 1, headerString);
                string firstCell = excelApp.cellNumToAlphabet(1, 1);
                string lastCell = excelApp.cellNumToAlphabet(1, col);
                excelApp.ApplyStyle(firstCell, lastCell);
                excelApp.Bold(1, 1, true);
                excelApp.Underline(1, 1, true);

                excelApp.MergeCells(firstCell, lastCell);

                foreach (var obj in rootObjects)
                {
                    reportData.Remove(obj);
                }
                int row = 3;
                int changeClass = 0;
                //int nextSection = 0;
                string elementHeader = string.Empty;
                bool stopWrite = true;

                int row1;
                bool newPage = false;
                List<int> listRowForPage = new List<int>();
                List<int> listColForPage = new List<int>();
                List<int> listCountPage = new List<int>();


                InsertBomSectionHeader(row, col, TFDClass.InsertSection(reportData[0].Type), excelApp);

                row++;


                for (int i = 0; i < reportData.Count; i++)
                {
                    row1 = row;
                    excelApp.BorderLineCells(row1, 1, row - 1, colCount - 1,
                    XlBordersIndex.xlEdgeTop,
                    XlBordersIndex.xlEdgeBottom,
                    XlBordersIndex.xlEdgeLeft,
                    XlBordersIndex.xlEdgeRight,
                    XlBordersIndex.xlInsideVertical);
                    if ((i > 0) && (reportData[i].Type != reportData[i - 1].Type))
                    {
                        string orderDenotation = string.Empty;

                        InsertBomSectionHeader(row, col, TFDClass.InsertSection(reportData[i].Type), excelApp);
                        countStr++;
                        if (!newPage)
                            excelApp.BorderLineCells(row1, 1, row - 1, colCount - 1,
                                                     XlBordersIndex.xlEdgeTop,
                                                     XlBordersIndex.xlEdgeBottom,
                                                     XlBordersIndex.xlEdgeLeft,
                                                     XlBordersIndex.xlEdgeRight,
                                                     XlBordersIndex.xlInsideVertical);
                        else
                            excelApp.BorderLineCells(row1, 1, row - 3, colCount - 1,
                                                    XlBordersIndex.xlEdgeTop,
                                                    XlBordersIndex.xlEdgeBottom,
                                                    XlBordersIndex.xlEdgeLeft,
                                                    XlBordersIndex.xlEdgeRight,
                                                    XlBordersIndex.xlInsideVertical);

                        row++;
                    }

                    row = InsertObject(reportData[i], row, "Документ", excelApp, logDelegate);
                    row = InsertObject(reportData[i], row, "Деталь", excelApp, logDelegate);
                    row = InsertObject(reportData[i], row, "Комплект", excelApp, logDelegate);




                }

                excelApp.BorderLineCells(2, 1, row - 1, colCount - 1,
                      XlBordersIndex.xlEdgeTop,
                      XlBordersIndex.xlEdgeBottom,
                      XlBordersIndex.xlEdgeLeft,
                      XlBordersIndex.xlEdgeRight,
                      XlBordersIndex.xlInsideVertical);


                excelApp.setValueCell(row + 1, 2, "Всего элементов - " + reportData.Count);

                var proc = Math.Round((Convert.ToDouble(reportData.Count) / Convert.ToDouble(reportDataOther.Count)) * 100, 2);

                excelApp.setValueCell(row + 2, 2, "Процент от общего количества - " + (reportDataOther.Count).ToString()
                    + "(" + proc.ToString() + "%)");

                excelApp.AutoWidth();
                excelApp.SaveDocument();


                // MessageBox.Show(process.ProcessName);
            }

        

            internal class NaimenAndOboznachComparer : IComparer<ReportParameters>
            {
                public int Compare(ReportParameters ob1, ReportParameters ob2)
                {
                    string designation1 = ob1.Denotation.Replace(" ", "");
                    string designation2 = ob2.Denotation.Replace(" ", "");
                    int ob = String.Compare(designation1, designation2);
                    if (ob != 0)
                        return ob;
                    else
                        return String.Compare(ob1.Name, ob2.Name);
                }
            }

        
    }
}
