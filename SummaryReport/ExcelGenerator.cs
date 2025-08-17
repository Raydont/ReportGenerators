using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Nomenclature;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Diagnostics;
using Font = System.Drawing.Font;
using ReportHelpers;


namespace SummaryReport
{
    public class ExcelGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
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
                selectionForm.MakeReport();
            }
           
        }    
    }

    // ТЕКСТ МАКРОСА EXCEL ===================================
 
    public delegate void LogDelegate(string line);

    public class ReportParameters
    {
        public Dictionary<int, int> ListObjectCountId = new Dictionary<int, int>();

        // типы объектов
        public bool Documentation; //Документация
        public bool Complex; //Комплексы
        public bool Detail; //Детали
        public bool Assembly; //Сборочные единицы
        public bool StandartItem; //Стандартные изделия
        public bool OtherItem; //Прочие изделия
        public bool Material; //Материалы
        public bool Complement; //Комплекты
        public bool ComplexProgram; //Комплексы (программы)
        public bool ComponentProgram; //Компоненты (программы)  
    }

    public delegate void SetProgressDelegate(int value);

    public class SLBuilderException : ApplicationException
    {
        public SLBuilderException(string message)
            : base(message)
        {
        }
    }

    public class SLBuilderFatalException : SLBuilderException
    {
        public SLBuilderFatalException(string message)
            : base(message)
        {
        }
    }

    //
    // Класс для запуска макроса
    //

    public class Report
    {
        //======================================================================
        //
        // Точка входа
        //
        //======================================================================

        public void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона
            bool isCatch = false;

            Xls xls = new Xls();

            MainForm m_form = new MainForm();
            m_form.Visible = true;
            LogDelegate m_writeToLog;
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);
                MakeExcelReport(context, xls, m_form, reportParameters, m_writeToLog);
                /*string str = @" taskkill /IM TFlex.DOCs.FilePreview.Provider.exe /F";
                  ExecuteCommand(str);*/
                xls[1, 1].Select();
                xls.AutoWidth();
                xls.Save();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    System.Diagnostics.Process.Start(context.ReportFilePath);
                }
            }
        }

        public void MakeExcelReport(IReportGenerationContext context, Xls xls, MainForm m_form, ReportParameters reportParameters, LogDelegate m_writeToLog)
        {
            try
            {
                // Получение ID объекта, на который получаем ВС
                int baseDocumentID = Initialize(context, m_form, m_writeToLog);
                if (baseDocumentID == -1) return;

                m_form.setStage(MainForm.Stages.DataAcquisition);
                m_writeToLog("=== Получение данных ===");
                m_form.progressBarParam(2);

                // Заполнить отчет данными
                FillData(m_writeToLog, reportParameters, xls);

                // Создание отчета на основе сформированных данных
                m_form.progressBarParam(2);
                m_form.setStage(MainForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                m_form.progressBarParam(2);

                m_form.setStage(MainForm.Stages.Done);
                m_writeToLog("=== Работа завершена ===");
                m_form.progressBarParam(2);
                m_writeToLog("Сводная ведомость сформирована");
                m_form.progressBarParam(2);
                m_form.Close();
                //                 System.Diagnostics.Process.Start(context.ReportFilePath);
            }
            catch (SLBuilderFatalException e)
            {
                string message = String.Format("ФАТАЛЬНАЯ ОШИБКА: {0}\r\nДля разрешения ситуации обратитесь к системному администратору", e.Message);
                if (m_writeToLog != null)
                    m_writeToLog(message);
                System.Windows.Forms.MessageBox.Show(message, "Фатальная ошибка",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (SLBuilderException e)
            {
                string message = String.Format("ОШИБКА ПРИЛОЖЕНИЯ: {0}", e.Message);
                if (m_writeToLog != null)
                    m_writeToLog(message);
                System.Windows.Forms.MessageBox.Show(message, "Ошибка приложения",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception e)
            {
                string message = String.Format("ОШИБКА: {0}", e.Message);
                if (m_writeToLog != null)
                    m_writeToLog(message);
                System.Windows.Forms.MessageBox.Show(message, "Ошибка",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        // Нахождение ID объекта для отчета ВС
        static int Initialize(IReportGenerationContext context, MainForm m_form, LogDelegate m_writeToLog)
        {
            // --------------------------------------------------------------------------
            // Инициализация
            // --------------------------------------------------------------------------

            int documentID;

            m_form.setStage(MainForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");
            m_form.progressBarParam(2);
            // Получение ID обрабатываемого документа

            m_writeToLog("Получение идентификатора корневого документа");
            m_form.progressBarParam(2);

            // Получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            m_writeToLog("Идентификатор корневого документа: " + documentID.ToString());
            m_form.progressBarParam(2);

            m_writeToLog("Подготовка к загрузке данных из T-FLEX DOCs");
            m_form.progressBarParam(2);

            return documentID;
        }


        public static List<SpecListParams> _allElements;              // коллекция всех элементов сводной ведомости
        public static List<SpecListParams> _elementsWithStructure;    // коллекция элементов с составом сводной ведомости

        public SpecListParams FillData(LogDelegate logDelegate, ReportParameters reportParameters, Xls xls)
        {

            SpecListParams baseDocument = FillData(
                                            out _allElements, out _elementsWithStructure, logDelegate, reportParameters, xls);
            return baseDocument;
        }

        public SpecListParams FillData(

                            out List<SpecListParams> allElements,
                            out List<SpecListParams> elementsWithStructure,
                            LogDelegate logDelegate, ReportParameters reportParameters, Xls xls)
        {

            // Инициализация выходных переменных

            // все элементы
            allElements = new List<SpecListParams>();
            // элементы с составом
            elementsWithStructure = new List<SpecListParams>();

            // Получение параметров документов
            for (int i = 0; i < reportParameters.ListObjectCountId.Count; i++)
            {
                ReadReportDocuments(reportParameters.ListObjectCountId.Keys.ToList()[i], reportParameters.ListObjectCountId.Values.ToList()[i], reportParameters, logDelegate);
                MakeSummaryReport(xls, logDelegate, reportParameters.ListObjectCountId.Values.ToList()[i]);
                _allElements = new List<SpecListParams>();
                _elementsWithStructure= new List<SpecListParams>();
            }

            return BaseDocument;
        }

        private static SpecListParams FillDocumentMainParameters(TFDDocument doc, bool BaseDocument)
        {
            SpecListParams pars = new SpecListParams(doc.Id);

            pars.ParentDocID = doc.ParentId;

            //класс объекта
            pars.ClassObject = doc.Class;

            // Параметр: Наименование  
            if (doc.Naimenovanie == null)
                throw new SLBuilderFatalException("FillDocumentMainParameters: Ошибка получения параметра Наименования");
            pars.Naimenovanie = doc.Naimenovanie;

            // Параметр: Обозначение
            if (doc.Oboznachenie == null)
                throw new SLBuilderFatalException("FillDocumentMainParameters: Ошибка получения параметра Обозначения");
            pars.Oboznachenie = doc.Oboznachenie;
            pars.Format = doc.Format;

            if (BaseDocument == false)
            {
                // Параметр: Количество
                pars.Count = doc.Amount;
            }
            else
            {
                pars.ParentDocID = 0;
                pars.Count = 1;
            }

            if (doc.MesuareUnit != null)
                pars.MeasureUnit = doc.MesuareUnit;

            return pars;
        }

        public class SpecListParams
        {

            private int _s_ObjectID; //ID документа
            private int _parentDocID; //ID родителя
            private double _count; //число изделий
            private string _oboznachenie; //обозначение документа
            private string _naimenovanie; //наименование документа
            private string _format; //формата документа
            private int _classObject; //класс документа
            private bool _readChildren; //сигнализатор прочтения потомков(для алгоритма обхода дерева)
            public string MeasureUnit; // единица измерения

            public SpecListParams(int DocumentID)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 1;
                _oboznachenie = "";
                _naimenovanie = "";
                _format = "";
            }
            public SpecListParams(int DocumentID, string obozna4enie, string naimenovanie, int classObject)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 1;
                _oboznachenie = obozna4enie;
                _naimenovanie = naimenovanie;
                _classObject = classObject;
            }

            public SpecListParams(int DocumentID, string obozna4enie, string naimenovanie)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 1;
                _oboznachenie = obozna4enie;
                _naimenovanie = naimenovanie;
            }

            // ID документа
            public int DocumentID
            {
                get { return _s_ObjectID; }
            }

            // ID родителя
            public int ParentDocID
            {
                get { return _parentDocID; }
                set { _parentDocID = value; }
            }

            // кол-во изделий
            public double Count
            {
                get { return _count; }
                set { _count = value; }
            }


            // Обозначение документа
            public string Oboznachenie
            {
                get { return _oboznachenie; }
                set { _oboznachenie = value; }
            }

            // Наименование документа
            public string Naimenovanie
            {
                get { return _naimenovanie; }
                set { _naimenovanie = value; }
            }

            // Формат документа
            public string Format
            {
                get { return _format; }
                set { _format = value; }
            }

            // Класс документа
            public int ClassObject
            {
                get { return _classObject; }
                set { _classObject = value; }
            }

            //сигнализатор прочтения потомков(для алгоритма обхода дерева)
            public bool ReadChildren
            {
                get { return _readChildren; }
                set { _readChildren = value; }
            }

            // посещены ли потомки объекта в коллекции
            public bool IsVisitedChildren(List<SpecListParams> familyList)
            {
                bool visited = true;
                for (int i = 0; i < familyList.Count; i++)
                    if (familyList[i].DocumentID == this.DocumentID) { visited = familyList[i].ReadChildren; break; }
                return visited;
            }

            // установка признака посещения потомков объенкта в коллекции
            public List<SpecListParams> SetVisitedChildren(List<SpecListParams> familyList)
            {
                for (int i = 0; i < familyList.Count; i++)
                    if (familyList[i].DocumentID == this.DocumentID) familyList[i].ReadChildren = false;

                return familyList;
            }

            // получение старшего сына
            public SpecListParams GetBigChild(List<SpecListParams> familyList)
            {
                List<SpecListParams> children = new List<SpecListParams>();

                for (int i = 0; i < familyList.Count; i++)
                    if (familyList[i].ParentDocID == this.DocumentID) children.Add(familyList[i]);

                if (children.Count > 0)
                {
                    children = Sort(children);
                    return children[0];
                }
                else return null;
            }

            //Проверка - имеет ли братьев объект в коллекции
            public SpecListParams GetNextLitleBrother(List<SpecListParams> familyList)
            {
                List<SpecListParams> brothers = new List<SpecListParams>();
                int indx = 0;

                for (int i = 0; i < familyList.Count; i++)
                    if (familyList[i].ParentDocID == this.ParentDocID)
                    {
                        brothers.Add(familyList[i]);
                        if (this.DocumentID == familyList[i].DocumentID)
                            indx = i;
                    }

                if (brothers.Count > 1)
                {
                    brothers = Sort(brothers);

                    for (int i = 0; i < brothers.Count; i++)
                        if (this.DocumentID == brothers[i].DocumentID)
                            indx = i;
                    if (brothers.Count != (indx + 1))
                        return brothers[indx + 1];
                    else return null;
                }
                else return null;
            }
        }

        public SpecListParams BaseDocument;

        //------------------------------------------------------
        public void ReadReportDocuments(
            int documentId,
            int documentCount, 
            ReportParameters reportParameters,
            LogDelegate logDelegate)
        {
            // соединяемся с БД T-FLEX DOCs 2012
            SqlConnection conn;

            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            sqlConStringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";
            conn = new SqlConnection(sqlConStringBuilder.ToString());
            conn.Open();
            // находим обозначение
            SqlCommand getBaseDocumentCodeCommand = new SqlCommand(String.Format(@"SELECT Denotation
                                                                                    FROM Nomenclature
                                                                                    WHERE Nomenclature.s_ObjectID = {0}
                                                                                        AND Nomenclature.s_ActualVersion = 1
                                                                                        AND Nomenclature.s_Deleted = 0", documentId), conn);
            getBaseDocumentCodeCommand.CommandTimeout = 0;
            string Oboznachenie = Convert.ToString(getBaseDocumentCodeCommand.ExecuteScalar());

            SqlCommand getBaseDocumentNameCommand = new SqlCommand(String.Format(@"SELECT [Name]
                        FROM Nomenclature
                        WHERE Nomenclature.s_ObjectID={0}
                            AND Nomenclature.s_ActualVersion = 1
                            AND Nomenclature.s_Deleted = 0", documentId), conn);
            getBaseDocumentNameCommand.CommandTimeout = 0;
            string Naimenovanie = Convert.ToString(getBaseDocumentNameCommand.ExecuteScalar());

            BaseDocument = new SpecListParams(documentId, Oboznachenie, Naimenovanie);

            #region Чтение дерева иерархии изделия
            //подготавливаем дерево документов хранилища
            //формирование временной таблицы Vspec c параметрами элементов сводной ведомости 
            string prepareCollectionCommandText = String.Format(
                              @"declare @docid INT
                                DECLARE @level int
                                declare @insertcount int
                                set @docid={0}
                                set @insertcount = 0
                                SET @level=0

                                IF OBJECT_ID('tempdb..#TmpVspec')is NULL
                                CREATE TABLE #TmpVspec (s_ObjectID INT,
                                                        [level] INT,
                                                        s_ParentID INT,
                                                        s_ClassID INT,
                                                        Amount FLOAT,
                                                        Denotation NVARCHAR(255),
                                                        Name NVARCHAR(255),
                                                        Format NVARCHAR(20),
                                                        MeasureUnit NVARCHAR(255),
                                                        TotalCount FLOAT,
                                                        Position INT,
                                                        Remarks NVARCHAR(255))
                                ELSE DELETE FROM #TmpVspec

                                INSERT INTO #TmpVspec
                                SELECT n.s_ObjectID,0,0,n.s_ClassID,1,n.Denotation,n.Name, n.Format,'',1,0,''
                                FROM Nomenclature n
                                WHERE n.s_ObjectID = @docid
                                      AND n.s_Deleted = 0 
                                      AND n.s_ActualVersion = 1 

                                WHILE 1=1
                                BEGIN

                                  INSERT INTO #TmpVspec 
                                  SELECT DISTINCT nh.s_ObjectID,
                                         @level+1,
                                         nh.s_ParentID,
                                         n.s_ClassID,
                                         nh.Amount,
                                         n.Denotation,
                                         n.Name,
                                         n.Format, 
                                         nh.MeasureUnit,
                                         #TmpVspec.TotalCount*nh.Amount,
                                         nh.Position,
                                         LEFT (nh.Remarks, 255)
                                  FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                           AND #TmpVspec.[level]=@level
                                                                         
                                       INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                  WHERE
                                                                              nh.s_ActualVersion = 1
                                                                           AND nh.s_Deleted = 0
                                                                           AND n.s_ActualVersion = 1
                                                                           AND n.s_Deleted = 0
                                  SET @insertcount = @@ROWCOUNT 
                                  SET @level = @level + 1 
   
                                  IF @insertcount = 0 
                                  GOTO end1

                                END
                                end1:

                                SELECT * FROM #TmpVspec", documentId);
            SqlCommand prepareCollectionCommand = new SqlCommand(prepareCollectionCommandText, conn);
            prepareCollectionCommand.CommandTimeout = 0;
            SqlDataReader reader = prepareCollectionCommand.ExecuteReader();
            #endregion

            SpecListParams pars = new SpecListParams(documentId);

            /* Чтение ID классов */
            int classAssembly = TFDClass.Assembly;
            int classComplement = TFDClass.Komplekt;
            int classComplex = TFDClass.Complex;
            int classComplexProgram = TFDClass.ComplexProgram;
            int classComponentProgram = TFDClass.ComponentProgram;
            List<TFDDocument> readObjects = new List<TFDDocument>();

            while (reader.Read())
            {
                int objectClass = GetSqlInt32(reader, 3);
               
                    int objectID = reader.GetInt32(0);
                    int parentID = GetSqlInt32(reader, 2);
                    TFDDocument doc = new TFDDocument(objectID, parentID, objectClass);
                    doc.Amount = GetSqlDouble(reader, 9) * documentCount;
                    doc.Oboznachenie = GetSqlString(reader, 5);
                    doc.Naimenovanie = GetSqlString(reader, 6);
                    doc.Format = GetSqlString(reader, 7);
                    if (GetSqlString(reader, 8).Trim() != string.Empty)
                        doc.MesuareUnit = GetSqlString(reader, 8);

                    //Проверка на дублированные объекты в структуре считанных данных
                    var dublicats = _allElements.Where(t => t.DocumentID == doc.Id && t.ParentDocID == doc.ParentId && t.Naimenovanie == doc.Naimenovanie && t.Oboznachenie == doc.Oboznachenie && t.ClassObject == doc.Class).ToList();
                    if (dublicats.Count > 0)
                    {
                        _allElements.Where(t => t.DocumentID == doc.Id && t.ParentDocID == doc.ParentId && t.Naimenovanie == doc.Naimenovanie && t.Oboznachenie == doc.Oboznachenie && t.ClassObject == doc.Class).ToList()[0].Count += doc.Amount;
                        continue;
                    }

                    // Добавление всех элементов в коллекцию
                    pars = FillDocumentMainParameters(doc, false);
                if (objectClass == TFDClass.Assembly && reportParameters.Assembly ||
                   objectClass == TFDClass.Detal && reportParameters.Detail ||
                   objectClass == TFDClass.Document && reportParameters.Documentation ||
                   objectClass == TFDClass.Complex && reportParameters.Complex ||
                   objectClass == TFDClass.ComplexProgram && reportParameters.ComplexProgram ||
                   objectClass == TFDClass.ComponentProgram && reportParameters.ComponentProgram ||
                   objectClass == TFDClass.Komplekt && reportParameters.Complement ||
                   objectClass == TFDClass.Material && reportParameters.Material ||
                   objectClass == TFDClass.OtherItem && reportParameters.OtherItem ||
                   objectClass == TFDClass.StandartItem && reportParameters.StandartItem)
                {
                    _allElements.Add(pars);
                }
                    //Добавление элементов с составом в коллекцию Элементов с составом

                    if ((objectClass == classAssembly) || (objectClass == classComplement) || (objectClass == classComplex) ||
                        (objectClass == classComplexProgram) || (objectClass == classComponentProgram))
                        _elementsWithStructure.Add(pars);
                    logDelegate(String.Format("Добавлен элемент: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));
                
            }          
            reader.Close();

            conn.Close();

        }

        public string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }

        public int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        public double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }

        public List<SpecListParams> TreeDirectByPass(List<SpecListParams> nodeList, LogDelegate logDelegate)
        {
            for (int i = 0; i < nodeList.Count; i++)
                nodeList[i].ReadChildren = true;

            Stack<SpecListParams> stack = new Stack<SpecListParams>();
            List<SpecListParams> nodeSummaryList = new List<SpecListParams>();

            SpecListParams node = new SpecListParams(0);

            SpecListParams litleBrother = new SpecListParams(0);

            stack.Push(nodeList[0]);
            nodeSummaryList.Add(nodeList[0]);

            while (stack.Count != 0)
            {

                node = stack.Peek();

                logDelegate(String.Format("Запись объекта: {0} {1}", node.Oboznachenie, node.Naimenovanie));

                if (node.IsVisitedChildren(nodeList))
                {
                    nodeList = node.SetVisitedChildren(nodeList);
                    node = node.GetBigChild(nodeList);
                    if (node != null)
                    {
                        stack.Push(node);
                        nodeSummaryList.Add(node);
                    }
                }
                else
                {
                    stack.Pop();

                    litleBrother = node.GetNextLitleBrother(nodeList);
                    if (litleBrother != null)
                    {
                        stack.Push(litleBrother);
                        nodeSummaryList.Add(litleBrother);
                    }
                }

            }
            return nodeSummaryList;
        }

        // Стиль отчета
        private static void PageSetup(Xls xls)
        {
            PageSetup pageSetup = xls.Worksheet.PageSetup;
            pageSetup.BottomMargin = 50.36;
            pageSetup.FooterMargin = 22.4;
            pageSetup.RightMargin = 5;
            pageSetup.LeftMargin = 50;
            pageSetup.TopMargin = 18 / 0.268 - 16.8;
            pageSetup.HeaderMargin = 22.4;
        }

        private static void SetValueCell(Range range, string value, bool bold, bool underline, bool center)
        {
            range.Font.Name = "Calibri";
            range.Font.Size = 11;
            if (bold)
                range.Font.Bold = true;
            if (underline)
                range.Font.Underline = true;
            if (center)
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignTop;
            range.SetValue(value);
        }

        int countStr = 10;  // количество строк на листе
        int countPage = 0;  // количество страниц
        int countWrap = 0;   // количество строк, получившееся в результате переноса строк в наименовании или обозначении
        bool newPage = false;

        public void MakeSummaryReport(Xls xls, LogDelegate logDelegate, int countDevice)
        {
            List<int> listRowForPage = new List<int>(); // спискок номеров строк, в которых необходимо писать номер страницы
            List<int> listCountPage = new List<int>(); //Список номеров страниц

            // Задание полей документа при печати
            PageSetup(xls);
            Font font = new Font("Calibri", 11f, GraphicsUnit.Document);

            TextFormatter _textFormatter = new TextFormatter(font);

            List<SpecListParams> structure = new List<SpecListParams>();
            List<SpecListParams> sortedStructure = new List<SpecListParams>();
            List<SpecListParams> summaryList = new List<SpecListParams>();
            List<SpecListParams> sortedElementsWithStructure = new List<SpecListParams>();
            SpecListParams baseElement = new SpecListParams(1);
            baseElement = _elementsWithStructure[0];
            SetValueCell(xls[1, row - 1], "Сводная ведомость для " + baseElement.Naimenovanie + " " + baseElement.Oboznachenie + " - " + countDevice + " шт.", true, false, true);
            xls[1, row - 1, 4, 1].Merge();
            xls[1, row - 1].Interior.Color = Color.LightGray;
            // формирование заголовков отчета сводной ведомости
            SetValueCell(xls[1, row], "Формат", true, false, false);
            SetValueCell(xls[2, row], "Обозначение", true, false, false);
            SetValueCell(xls[3, row], "Наименование", true, false, false);
            SetValueCell(xls[4, row], "Количество", true, false, false);
            xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                        XlBordersIndex.xlEdgeRight,
                                                                                        XlBordersIndex.xlInsideVertical);
             
            _elementsWithStructure.RemoveAt(0);
            sortedElementsWithStructure.Add(baseElement);
            sortedElementsWithStructure.AddRange(_elementsWithStructure);

            summaryList = TreeDirectByPass(sortedElementsWithStructure, logDelegate);
            int row1 = 0;

            foreach (SpecListParams element in summaryList)
            {
                if ((countPage % 30 == 0) && (countPage != 0))
                    countStr += 8;
                row1 = row;
                row++;
                SetValueCell(xls[1, row], element.Naimenovanie + " " + element.Oboznachenie, true, true, true);
                xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                               XlBordersIndex.xlEdgeBottom,
                                                                                               XlBordersIndex.xlEdgeLeft,
                                                                                               XlBordersIndex.xlEdgeRight,
                                                                                               XlBordersIndex.xlInsideVertical);
                xls[1, row, 4, 1].Merge();
                logDelegate(String.Format("Запись спецификации для объекта: {0} {1}", element.Oboznachenie, element.Naimenovanie));
                countStr++;

                foreach (SpecListParams entrance in _allElements)
                    if (entrance.ParentDocID == element.DocumentID)
                        structure.Add(entrance);
                sortedStructure = Sort(structure);
                structure.Clear();
                if (sortedStructure.Count == 0)
                {
                    row++;
                    countStr++;
                    SetValueCell(xls[1, row], "Состав отсутствует", false, false, true);
                    xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                   XlBordersIndex.xlEdgeRight,
                                                                                                   XlBordersIndex.xlInsideVertical);
                    xls[1, row, 4, 1].Merge();
                }
                foreach (SpecListParams param in sortedStructure)
                {
                    row++;
                    // заполнение формата
                    SetValueCell(xls[1, row], param.Format, false, false, false);
                    xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    string name = param.Naimenovanie;

                    string[] wrapTextNaimenovanie = _textFormatter.Wrap(name.Trim(), 34f); // разбиение наименования на несколько строк
                    name = text(wrapTextNaimenovanie);

                    string denotation = param.Oboznachenie;

                    string[] wrapTextDenotation = _textFormatter.Wrap(denotation.Trim(), 23f); // разбиение обозначения на несколько строк
                    denotation = text(wrapTextDenotation);

                    // заполнение обозначения
                    SetValueCell(xls[2, row], denotation, false, false, false);
                    xls[2, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    // заполнение наименования
                    SetValueCell(xls[3, row], name, false, false, false);

                    if (wrapTextDenotation.Count() >= wrapTextNaimenovanie.Count())
                        countWrap = wrapTextDenotation.Count();
                    else
                        countWrap = wrapTextNaimenovanie.Count();

                    countStr += countWrap;

                    if (countStr > 50) // максимальное количество строк на листе 
                    {

                        countPage++;
                        int countRowDec = 2;

                        // Для случая когда  первая строка страницы - заголовок раздела
                        if (((countStr == 51) && (countWrap == 1)) || ((countStr == 52) && (countWrap == 2)) || ((countStr == 53) && (countWrap == 3))
                            || ((countStr == 54) && (countWrap == 4)) || ((countStr == 55) && (countWrap == 5)))
                        {
                            countRowDec = countWrap + 2;
                        }


                        if (wrapTextDenotation.Count() >= wrapTextNaimenovanie.Count())
                            countStr = wrapTextDenotation.Count();
                        else
                            countStr = wrapTextNaimenovanie.Count();

                        listRowForPage.Add(row - countRowDec);
                        listCountPage.Add(countPage);

                        // корректировка смещения строк
                        if (countPage % 30 == 0)
                            countStr += 4;


                        newPage = true;
                    }
                    else
                    {
                        newPage = false;
                    }

                    if (param.Count != 0)
                    {
                        if (param.MeasureUnit != null)
                        {
                            SetValueCell(xls[4, row], param.Count + " " + param.MeasureUnit, false, false, false);
                            xls[4, row].HorizontalAlignment = XlHAlign.xlHAlignRight;
                        }
                        else
                        {
                            SetValueCell(xls[4, row], param.Count.ToString(), false, false, false);
                            xls[4, row].HorizontalAlignment = XlHAlign.xlHAlignRight;
                        }
                    }

                    xls[1, row, 4, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                   XlBordersIndex.xlEdgeBottom,
                                                                                                   XlBordersIndex.xlEdgeLeft,
                                                                                                   XlBordersIndex.xlEdgeRight,
                                                                                                   XlBordersIndex.xlInsideVertical);
                }






            }

            //Заполнение номеров страниц
//----------------------------------------------------------------------------------------------------------------------

            //// заполнение ячеек номерами страниц
            //for (int i = 0; i < listRowForPage.Count; i++)
            //{
            //    if (listRowForPage.Count > 1)
            //        SetValueCell(xls[5, listRowForPage[i]], "с" + listCountPage[i].ToString() + "из" +
            //            (listCountPage[listCountPage.Count - 1] + 1), false, false, false);
            //}

            //// Заполнение ячейки номером страницы последней строки 
            //if (listRowForPage.Count > 1)
            //    SetValueCell(xls[5, row], "с" + (listCountPage[listCountPage.Count - 1] + 1) + "из" +
            //        (listCountPage[listCountPage.Count - 1] + 1), false, false, false);
            row = row + 3;
        }


        public int row = 2;

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
        public static List<SpecListParams> Sort(List<SpecListParams> structure)
        {
            List<SpecListParams> sortedStructure = new List<SpecListParams>();
            List<SpecListParams> sortedByClass = new List<SpecListParams>();

            //Добавление Документов в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Document)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Комплексов в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Complex)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Сборочных единиц в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Assembly)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Деталей в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Detal)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Стандартных изделий в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.StandartItem)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Прочих изделий в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.OtherItem)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Материалов в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Material)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Комплектов в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.Komplekt)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление  Комплексов (программы) в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.ComplexProgram)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            //Добавление Компонентов (программы) в коллекцию
            foreach (SpecListParams param in structure)
                if (param.ClassObject == TFDClass.ComponentProgram)
                    sortedByClass.Add(param);
            SortData(sortedByClass);
            sortedStructure.AddRange(sortedByClass);
            sortedByClass.Clear();

            return sortedStructure;
        }

        static public void SortData(List<SpecListParams> data)
        {
            if (data == null)
                return;

            if (data.Count < 2)
                return;
            FindDublicates(data);
            data.Sort(new NaimenAndOboznachComparer()); //сортируем записи по обозначению и наименованию
        }


        internal class NaimenAndOboznachComparer : IComparer<SpecListParams>
        {
            public int Compare(SpecListParams ob1, SpecListParams ob2)
            {
                string designation1 = ob1.Oboznachenie.Replace(" ", "");
                string designation2 = ob2.Oboznachenie.Replace(" ", "");
                int ob = String.Compare(designation1, designation2);
                if (ob != 0)
                    return ob;
                else
                    return String.Compare(ob1.Naimenovanie, ob2.Naimenovanie);
            }
        }

        static private void FindDublicates(List<SpecListParams> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (String.Equals(data[mainID].Naimenovanie, data[slaveID].Naimenovanie) &&
                        String.Equals(data[mainID].Oboznachenie, data[slaveID].Oboznachenie))
                    {
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
        }
    }

    internal class TFDDocument
    {
        /* Порядок полей в запросе чтения документа */
        public int Id;
        public int ParentId;
        public int Class;
        public double Amount;
        public string Naimenovanie;
        public string Oboznachenie;
        public string Format;
        public string MesuareUnit;
       
        public TFDDocument(int id, int parent, int docClass)
        {
            Id = id;
            ParentId = parent;
            Class = docClass;            
        }
        public TFDDocument()
        {
          
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

        public double GetTextWidth(string text, Font font)
        {
            Size size = TextRenderer.MeasureText(text, font);
            double width = 31f * widthFactor * (double)size.Width / (double)_displayResolutionX;
            return width;
        }




        public string[] Wrap(string text, double maxWidth)
        {
            return Wrap(text, _font, maxWidth);
        }


        public string[] WrapFullName(string textDesignation, string text, double maxWidth, string docName, string docDenotation)
        {
            List<string> wrappedLines = new List<string>();
            int indx = text.IndexOf(docName);
            int indxDesignation = textDesignation.IndexOf(docDenotation);

            if (indx == 0)
            {

                text = text.Remove(indx, docName.Length);
                wrappedLines.AddRange(Wrap(docName, _font, maxWidth));
                wrappedLines.AddRange(Wrap(text, _font, maxWidth));
                return wrappedLines.ToArray();

            }
            else
                if (indxDesignation == 0)
                {
                    wrappedLines.AddRange(Wrap(docName, _font, maxWidth));
                    wrappedLines.AddRange(Wrap(text, _font, maxWidth));
                    return wrappedLines.ToArray();
                }
                else
                    return Wrap(text, _font, maxWidth);

        }

        public string[] WrapNotFullName(string text, double maxWidth, string docName)
        {
            List<string> wrappedLines = new List<string>();
            int indx = text.IndexOf(docName);
            if (indx == 0)
            {
                text = text.Remove(indx, docName.Length);
                wrappedLines.AddRange(Wrap(text, _font, maxWidth));
                return wrappedLines.ToArray();
            }


            return Wrap(text, _font, maxWidth);
        }
        //Разбиение строки примечания на подстроки
        public string[] WrapNote(string text, double maxWidth)
        {
            //                string patternThreeDots = @"(?:(\d[.,][0-9]{1,3})?\D[0-9]{1,3}[\.]{3}\D[0-9]{1,3}\,?)*";
            string patternThreeDots = @"(?:\D[0-9]{1,3}[\.]{3}\D[0-9]{1,3}\,?((?=\D[0-9]{1,3}[\.]{3}\D[0-9]{1,3}\,?))?)";
            Regex regex = new Regex(patternThreeDots);
            MatchCollection match;
            match = regex.Matches(text);

            int indx;
            List<string> rowThreeDots = new List<string>();
            List<string> wrappedLines = new List<string>();
            string beginPart;



            for (int i = 0; i < match.Count; i++)
            {
                if (match[i].Value.Trim() != string.Empty)
                    rowThreeDots.Add(match[i].Value);
            }



            if (rowThreeDots.Count > 0)
            {
                foreach (string row in rowThreeDots)
                {
                    indx = text.IndexOf(row);
                    beginPart = text.Remove(indx);
                    if (beginPart.Trim() != string.Empty)
                    {
                        text = text.Remove(0, beginPart.Length);
                        text = text.Trim();
                        wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                    }
                    wrappedLines.Add(row);
                    text = text.Remove(0, row.Length);
                }
                if (text.Trim() != string.Empty)
                {
                    text = text.Trim();
                    int index = text.IndexOf(",");
                    if (index == 0)
                        text = text.Remove(0, 1);
                    wrappedLines.AddRange(Wrap(text, maxWidth));
                }
                string str = string.Empty;

                return wrappedLines.ToArray();
            }
            else return Wrap(text, _font, maxWidth);
        }

        /// Разбивка на строки по символам перевода строки \n
        string[] Wrap(string text, Font font, double maxWidth)
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

        //***
        //Разбивка наименований Прочих, Стандартных изделий и материалов на подстроки
        public string[] Wrap(string text, double maxWidth, string sectionName)
        {
            List<string> wrappedLines = new List<string>();

            string patternBeginPart = @"^(\S{2,}(\s+(?:[А-Я]{1})?[а-яё]{2,}(?=\s|$))*)";
            string patternEndPart = String.Empty;
            string patternBrackets = @"(?:(\(по+[\s,\S]+\)))$";
            Regex regex;
            Match match;
            string endPart = String.Empty;
            string beginPart = String.Empty;
            string remPart = String.Empty;
            string bracketsPart = String.Empty;
            int indx;


            if (sectionName == "Прочие изделия")
                patternEndPart = @"(?:(\S+(\s)*ТУ)|(ТУ(\s)*\S+)|(\S+ТУ\S+))$";

            if (sectionName == "Стандартные изделия")
                patternEndPart = @"(?<standard>(ИСО|ГОСТ|ОСТ).+)?$";

            regex = new Regex(patternEndPart);
            match = regex.Match(text);
            endPart = match.Value;
            endPart = endPart.Trim();
            indx = text.IndexOf(endPart);
            remPart = text.Remove(indx);

            if (endPart.Trim() != String.Empty)
            {
                if (GetTextWidth(remPart, _font) > maxWidth)
                {
                    regex = new Regex(patternBeginPart);
                    match = regex.Match(remPart);
                    beginPart = match.Value;
                    beginPart = beginPart.Trim();

                    if (GetTextWidth(remPart, _font) > maxWidth)
                        wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                    else wrappedLines.Add(beginPart);

                    remPart = remPart.Remove(0, beginPart.Length);
                }
                remPart = remPart.Trim();

                if (remPart.Trim() != String.Empty)
                {
                    if (GetTextWidth(remPart, _font) > maxWidth + maxWidth / 4)
                        wrappedLines.AddRange(Wrap(remPart, maxWidth));
                    else
                    {
                        remPart = remPart.Trim();
                        wrappedLines.Add(remPart);
                    }
                }
                if (wrappedLines.Count >= 2)
                {
                    if (((GetTextWidth(wrappedLines[wrappedLines.Count - 1] + wrappedLines[wrappedLines.Count - 2], _font)) < maxWidth + maxWidth / 5) && ((wrappedLines[wrappedLines.Count - 2] != "Резистор") && (wrappedLines[wrappedLines.Count - 2] != "Конденсатор") && (wrappedLines[wrappedLines.Count - 2] != "Транзисторная матрица")))
                    {
                        wrappedLines[wrappedLines.Count - 2] = wrappedLines[wrappedLines.Count - 2] + " " + wrappedLines[wrappedLines.Count - 1];
                        if (remPart.Trim() != String.Empty)
                            wrappedLines[wrappedLines.Count - 1] = endPart;
                        return wrappedLines.ToArray();
                    }
                    else if (remPart.Trim() != String.Empty)
                        wrappedLines.Add(endPart);
                }

                else if (remPart.Trim() != String.Empty)
                    wrappedLines.Add(endPart);
            }
            else
            {
                regex = new Regex(patternBrackets);
                match = regex.Match(text);
                bracketsPart = match.Value;
                bracketsPart = bracketsPart.Trim();

                if (bracketsPart != String.Empty)
                {
                    indx = text.IndexOf(bracketsPart);
                    text = text.Remove(indx);
                    wrappedLines.AddRange(Wrap(text, maxWidth));
                    wrappedLines.Add(bracketsPart);
                }
                else wrappedLines.AddRange(Wrap(text, maxWidth));

                if ((wrappedLines.Count >= 2) && (GetTextWidth(wrappedLines[wrappedLines.Count - 1] + wrappedLines[wrappedLines.Count - 2], _font) < maxWidth + maxWidth / 5) && ((wrappedLines[wrappedLines.Count - 2] != "Резистор") && (wrappedLines[wrappedLines.Count - 2] != "Конденсатор")))
                {
                    wrappedLines[wrappedLines.Count - 2] = wrappedLines[wrappedLines.Count - 2] + " " + wrappedLines[wrappedLines.Count - 1];
                    wrappedLines.RemoveAt(wrappedLines.Count - 1);
                }
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
        string[] WrapBySyllables(string text, Font font, double maxWidth)
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
        string[] WrapByAlphaAndDelimiter(string text, Font font, double maxWidth)
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
        string[] WrapByCharacters(string text, Font font, double maxWidth)
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

    public static class TFDClass
    {
        public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find(new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c")).Id;
        public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find(new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea")).Id;
        public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
        public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find(new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340")).Id;
        public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find(new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74")).Id;
        public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find(new Guid("f50df957-b532-480f-8777-f5cb00d541b5")).Id;
        public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find(new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5")).Id;
        public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find(new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987")).Id;
        public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find(new Guid("a1862b2c-032c-48af-9c9f-ab7ace0d5b2f")).Id;
        public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find(new Guid("b7f7df88-eefa-4d73-a4dc-c08c46d584d1")).Id;
    }
}

   
