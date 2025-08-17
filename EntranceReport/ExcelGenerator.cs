using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Excel;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Nomenclature;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Application = Microsoft.Office.Interop.Excel.Application;
using ReportHelpers;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;

namespace EntranceReport
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

        AddFilterForm addFilterForm = new AddFilterForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            using (new WaitCursorHelper(false))
            {
                addFilterForm.DotsEntry(context);
            }
        }
    }

    public delegate void LogDelegate(string line);
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
    public class WaitCursorHelper : IDisposable
    {
        private bool _waitCursor;

        public WaitCursorHelper(bool useWaitCursor)
        {
            _waitCursor = System.Windows.Forms.Application.UseWaitCursor;
            System.Windows.Forms.Application.UseWaitCursor = useWaitCursor;
        }

        public void Dispose()
        {
            System.Windows.Forms.Application.UseWaitCursor = _waitCursor;
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
        public void Make(IReportGenerationContext context,  ReportParameters filterParameters)
        {
            bool isCatch = false;
            context.CopyTemplateFile();    // Создаем копию шаблона
            Xls xls = new Xls();

            var m_form = new MainForm();
            m_form.Visible = true;
            LogDelegate m_writeToLog;
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);
                MakeExcelReport(context, xls, m_form, m_writeToLog, filterParameters);
                xls[1, 1].Select();
                xls.AutoWidth();
                // Сохранение документа
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

        public void MakeExcelReport(IReportGenerationContext context, Xls xls, MainForm m_form, LogDelegate m_writeToLog, ReportParameters filterParameters)
        {
            MakeReport(context, xls, m_form, m_writeToLog, filterParameters);
        }


        public void MakeReport(IReportGenerationContext context, Xls xls, MainForm m_form, LogDelegate m_writeToLog, ReportParameters filterParameters)
        {
            try
            {
                // Получение ID объекта, на который получаем ВС
                int baseDocumentID = Initialize(context, m_form, m_writeToLog);
                if (baseDocumentID == -1) return;

                m_form.setStage(MainForm.Stages.DataAcquisition);
                m_writeToLog("=== Получение и обработка данных ===");
                m_form.progressBarParam(2);

                // Заполнить отчет данными
                ReportData reportData = new ReportData();
                //ReportData.SpecListParams baseDocumentSpecListParams = 
                reportData.FillData(baseDocumentID, m_writeToLog, xls, filterParameters);
                m_form.progressBarParam(4);

                // Создание отчета на основе сформированных данных
                m_form.setStage(MainForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                m_form.progressBarParam(2);

                m_form.setStage(MainForm.Stages.Done);
                m_writeToLog("=== Работа завершена ===");
                m_form.progressBarParam(2);
                m_writeToLog("Ведомость входимостей сформирована");
                m_form.progressBarParam(2);
                m_form.Close();
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
    }

    //

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

        public TFDDocument(int id, int parent, int docClass)
        {
            Id = id;
            ParentId = parent;
            Class = docClass;
        }
    }

    class ReportData
    {
        //раздел сборочные единицы
        public SpecListParams FillData(int baseDocumentID, LogDelegate logDelegate, Xls xls, ReportParameters filterParameters)
        {
            MakeEntranceReport(xls, logDelegate, ReadReportDocuments(baseDocumentID, null, logDelegate,filterParameters), filterParameters);
            return BaseDocument;
        }

        internal class CollectionAllElements
        {
            public List<SpecListParams> documentation = new List<SpecListParams>();
            public List<SpecListParams> complexes= new List<SpecListParams>();
            public List<SpecListParams> assemblies= new List<SpecListParams>();
            public List<SpecListParams> detalies= new List<SpecListParams>();
            public List<SpecListParams> standartItems=new List<SpecListParams>();
            public List<SpecListParams> otherItems = new List<SpecListParams>();
            public List<SpecListParams> materials = new List<SpecListParams>();
            public List<SpecListParams> complects = new List<SpecListParams>();
            public List<SpecListParams> complementsPrograms = new List<SpecListParams>();
            public List<SpecListParams> complexesPrograms = new List<SpecListParams>();
        }
        internal class SpecListParams
        {
            public class Entrances
            {
                public string Obozna4enie; //обозначение головного изделия
                public string Naimenovanie;  //наименование головного изделия
                public string UnitMeasure;  //Единица измерения
                public string Remarks;      //Примечание
                public double Count; //число вхождений в головное изделие
                public double TotalCount; //общее число вхождений
                public int Position;       //Позиция
                public string Zone;       //Зона

                public Entrances(string obozna4enie, string naimenovanie, string unitMeasure, string remarks, double count, double commonCount, int position, string zone)
                {
                    Obozna4enie = obozna4enie;
                    Naimenovanie = naimenovanie;
                    UnitMeasure = unitMeasure;
                    Remarks = remarks;
                    Count = count;
                    TotalCount = commonCount;
                    Position = position;
                    Zone = zone;                                                      
                }
            }

            private int _s_ObjectID; //ID документа
            private int _parentDocID; //ID родителя
            private double _count; //число изделий
            private double _totalcount; //общее количество изделий
            private string _oboznachenie; //обозначение документа
            private string _naimenovanie; //наименование документа
          //  private string _prime4anie; //примечание спецификации
            private bool _vhoditVSostav; //составная часть входит в состав изделия
            public List<Entrances> Include; //список вхождений обозначений спецификаций
            private string _format; //формат документа
            private int _positions; //позиция
            private string _zone; //зона
            private string _remarks; //коментарий
           // private string _unitMeasure; //единица измерения

            public SpecListParams(int DocumentID)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 0;
                _oboznachenie = "";
                _naimenovanie = "";
             //   _prime4anie = "";
             //   _positions = 0;
           //     _zone = "";
            //    _remarks = "";
            //    _unitMeasure = "";
                _vhoditVSostav = false;
                Include = new List<Entrances>();
            }

            public SpecListParams(int DocumentID, string obozna4enie, string naimenovanie)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 1;
                _oboznachenie = obozna4enie;
                _naimenovanie = naimenovanie;
             //   _prime4anie = "";
                _vhoditVSostav = false;
                Include = new List<Entrances>();
            }

            public override string ToString()
            {
                return _oboznachenie
                    + "(id=" + _s_ObjectID.ToString()
                    + ", наим=" + _naimenovanie
                  //  + ", прим=" + _prime4anie
                    + ", входит=" + _vhoditVSostav.ToString()
                    + ", кол=" + _count
                    + ", parentID=" + _parentDocID.ToString()
                    + ")";
            }

            // ID документа
            public int DocumentID
            {
                get { return _s_ObjectID; }
            }

            public int ParentDocID //ID родителя
            {
                get { return _parentDocID; }
                set { _parentDocID = value; }
            }

            public double TotalCount
            {
                get { return _totalcount; }
                set { _totalcount = value; }
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

            //составная часть входит в состав изделия
            public bool VhoditVSostav
            {
                get { return _vhoditVSostav; }
                set { _vhoditVSostav = value; }
            }

            // формат документа
            public string Format
            {
                get { return _format; }
                set { _format = value; }
            }

            //public int Positions
            //{
            //    get { return _positions;}
            //    set { _positions = value;}
            //}

            //public string Zone
            //{
            //    get{ return _zone;}
            //    set{ _zone = value;}
            //}

            //public string Remarks
            //{
            //    get{ return _remarks;}
            //    set{ _remarks = value;}
            //}

            //public string UnitMeasure
            //{
            //    get{ return _unitMeasure;}
            //    set{ _unitMeasure = value;}
            //}
        }

        public static SpecListParams BaseDocument;

        //------------------------------------------------------
        public CollectionAllElements ReadReportDocuments(
            int documentId,
            SpecListParams baseDocumentSpecListParams,
            LogDelegate logDelegate,
            ReportParameters filterParameters)
        {

            // соединяемся с БД T-FLEX DOCs 2012
            SqlConnection conn = null;

            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
                //Создаем ссылку на справочник
                ReferenceInfo info = ServerGateway.Connection.ReferenceCatalog.Find(TFlex.DOCs.Model.Structure.SpecialReference.GlobalParameters);
                Reference reference = info.CreateReference();
                var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
                sqlConStringBuilder.DataSource = nameSqlServer;
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

            // находим первичную применяемость
            SqlCommand getFirstUseCommand = new SqlCommand(String.Format(@"SELECT ConstructorParams.FirstUse     
                                                                            FROM  ConstructorParams 
                                                                            WHERE ConstructorParams.s_ObjectID = {0}
                                                                                AND ConstructorParams.s_ActualVersion = 1
                                                                                AND ConstructorParams.s_Deleted = 0", documentId), conn);
            getFirstUseCommand.CommandTimeout = 0;
            string firstUse = Convert.ToString(getFirstUseCommand.ExecuteScalar());


            SqlCommand getBaseDocumentNameCommand = new SqlCommand(String.Format(@"SELECT [Name]
                        FROM Nomenclature
                        WHERE Nomenclature.s_ObjectID={0}
                            AND Nomenclature.s_ActualVersion = 1
                            AND Nomenclature.s_Deleted = 0", documentId), conn);
            getBaseDocumentNameCommand.CommandTimeout = 0;
            string Naimenovanie = Convert.ToString(getBaseDocumentNameCommand.ExecuteScalar());

            BaseDocument = new SpecListParams(documentId, Oboznachenie, Naimenovanie);

            SpecListParams pars = new SpecListParams(documentId);

            /* Чтение ID классов */
            int classDocumentation = TFDClass.Document;
            int classComplex = TFDClass.Complex;
            int classAssembly = TFDClass.Assembly;
            int classDetal = TFDClass.Detal;
            int classStandartItems = TFDClass.StandartItem;
            int classOtherItems = TFDClass.OtherItem;
            int classMaterial = TFDClass.Material;
            int classComplement = TFDClass.Komplekt;
            int classComplexProgram = TFDClass.ComplexProgram;
            int classComponentProgram = TFDClass.ComponentProgram;

            #region Чтение полей ВС
            //Читаем все поля ВС
            string getDocTreeCommandText = String.Format(@"
                           declare @docid INT
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
                                                        Position INT,
                                                        MeasureUnit NVARCHAR(255),
                                                        Zone NVARCHAR(20),
                                                        EntranceDenotation NVARCHAR(255),
                                                        EntranceName NVARCHAR(255),
                                                        TotalCount FLOAT,
                                                        Remarks NVARCHAR(MAX))
                                ELSE DELETE FROM #TmpVspec

                                INSERT INTO #TmpVspec
                                SELECT n.s_ObjectID,0,0,n.s_ClassID,1,n.Denotation,n.Name,'',0,'','','','',1,''
                                FROM Nomenclature n
                                WHERE n.s_ObjectID = @docid
                                      AND n.s_Deleted = 0 
                                      AND n.s_ActualVersion = 1 

                                WHILE 1=1
                                BEGIN

                                  INSERT INTO #TmpVspec 
                                  SELECT  nh.s_ObjectID,
                                         @level+1,
                                         nh.s_ParentID,
                                         n.s_ClassID,
                                         nh.Amount,
                                         n.Denotation,
                                         n.Name,
                                         n.Format,
                                         nh.Position,
                                         nh.MeasureUnit,
                                         nh.Zone,
                                         #TmpVspec.Denotation,
                                         #TmpVspec.Name,
                                         #TmpVspec.TotalCount*nh.Amount,
                                         nh.Remarks
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

                                SELECT s_ObjectID, s_ParentID, s_ClassID,  Denotation, Name, EntranceDenotation,Format, Amount,  SUM(TotalCount) as Summa, Remarks,  EntranceName, Position, Zone, MeasureUnit
                                FROM #TmpVspec 
                                ----Вставка условий для поиска
                                GROUP BY Denotation, Name,  EntranceDenotation,EntranceName,s_ObjectID, s_ParentID, s_ClassID, Format, Amount, Remarks, Position, Zone, MeasureUnit  
                                ORDER BY Denotation, Name,  EntranceDenotation,EntranceName,s_ObjectID, s_ParentID, s_ClassID, Format, Amount, Remarks, Position, Zone, MeasureUnit ", documentId);


            int indx = 0;
            if (filterParameters.AddLikeName.Count > 0)
            {
                indx = getDocTreeCommandText.IndexOf("----") - 1;
                getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "WHERE");
            }

            for (int i = 0; i < filterParameters.AddLikeName.Count; i++)
            {
                indx = getDocTreeCommandText.IndexOf("----") - 1;


                if (filterParameters.AddLikeDenotation[i] != null && filterParameters.AddLikeDenotation[i] != string.Empty)
                {
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                          string.Format(
                                                                              @" (Denotation Like N'%{0}%' ",
                                                                              filterParameters.AddLikeDenotation[i]));


                    if (filterParameters.AddLikeName[i] != null && filterParameters.AddLikeName[i] != string.Empty)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" AND Name Like N'%{0}%') ",
                                                                                  filterParameters.AddLikeName[i]));
                    }
                    else
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx, ")");
                    }
                }
                else
                {
                    if (filterParameters.AddLikeName[i] != null && filterParameters.AddLikeName[i] != string.Empty)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" Name Like N'%{0}%' ",
                                                                                  filterParameters.AddLikeName[i]));
                    }
                }
                if (i < filterParameters.AddLikeName.Count - 1)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "OR");
                }
            }



            SqlCommand docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
            docTreeCommand.CommandTimeout = 0;
            SqlDataReader reader = docTreeCommand.ExecuteReader();

            bool entrDocumentation = false;
            bool entrComplex = false;
            bool entrAssembly = false;
            bool entrDetal = false;
            bool entrStandartItem = false;
            bool boolOtherItem = false;
            bool entrMaterial = false;
            bool entrComplect = false;
            bool entrComponentProgram = false;
            bool entrComplexProgram = false;

            var collectionAllElements = new CollectionAllElements();
            int objectId = 0;


            while (reader.Read())
            {
                if (GetSqlInt32(reader, 0) != documentId)
                {
                    if (objectId != GetSqlInt32(reader, 0))
                    {
                        //Документация
                        if (GetSqlInt32(reader, 2) == classDocumentation && filterParameters.Documentation)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.documentation.Add(pars);
                            logDelegate(String.Format("Добавлен Документ: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = true;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                                                 GetSqlString(reader, 13),
                                                                                 GetSqlString(reader, 9),
                                                                                 GetSqlDouble(reader, 7),
                                                                                 GetSqlDouble(reader, 8),
                                                                                 GetSqlInt32(reader, 11),
                                                                                 GetSqlString(reader, 12));
                            collectionAllElements.documentation[collectionAllElements.documentation.Count - 1].Include.Add(include);
                        }

                        else
                            if (GetSqlInt32(reader, 2) == classComplex && filterParameters.Complex)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.complexes.Add(pars);
                            logDelegate(String.Format("Добавлен Комплекс: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = true;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                                            GetSqlString(reader, 13),
                                                                            GetSqlString(reader, 9),
                                                                            GetSqlDouble(reader, 7),
                                                                            GetSqlDouble(reader, 8),
                                                                            GetSqlInt32(reader, 11),
                                                                            GetSqlString(reader, 12));
                            collectionAllElements.complexes[collectionAllElements.complexes.Count - 1].Include.Add(include);
                        }
                        else
                                if (GetSqlInt32(reader, 2) == classAssembly && filterParameters.Assembly)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.assemblies.Add(pars);
                            logDelegate(String.Format("Добавлена СЕ: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = true;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                                       GetSqlString(reader, 13),
                                                                       GetSqlString(reader, 9),
                                                                       GetSqlDouble(reader, 7),
                                                                       GetSqlDouble(reader, 8),
                                                                       GetSqlInt32(reader, 11),
                                                                       GetSqlString(reader, 12));
                            collectionAllElements.assemblies[collectionAllElements.assemblies.Count - 1].Include.Add(include);
                        }
                        else
                                    if (GetSqlInt32(reader, 2) == classDetal && filterParameters.Detail)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.detalies.Add(pars);
                            logDelegate(String.Format("Добавлена Деталь: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = true;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;


                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                                   GetSqlString(reader, 13),
                                                                   GetSqlString(reader, 9),
                                                                   GetSqlDouble(reader, 7),
                                                                   GetSqlDouble(reader, 8),
                                                                   GetSqlInt32(reader, 11),
                                                                   GetSqlString(reader, 12));
                            collectionAllElements.detalies[collectionAllElements.detalies.Count - 1].Include.Add(include);
                        }
                        else
                                        if (GetSqlInt32(reader, 2) == classStandartItems && filterParameters.Standart)
                        {
                            pars = WriteParameters(reader);

                            collectionAllElements.standartItems.Add(pars);
                            logDelegate(String.Format("Добавлено Стандартное изделие: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = true;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                                GetSqlString(reader, 13),
                                                                GetSqlString(reader, 9),
                                                                GetSqlDouble(reader, 7),
                                                                GetSqlDouble(reader, 8),
                                                                GetSqlInt32(reader, 11),
                                                                GetSqlString(reader, 12));
                            collectionAllElements.standartItems[collectionAllElements.standartItems.Count - 1].Include.Add(include);
                        }
                        else
                                            if (GetSqlInt32(reader, 2) == classOtherItems && filterParameters.Other)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.otherItems.Add(pars);
                            logDelegate(String.Format("Добавлено Прочее изделие: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = true;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            double amountByRemark = GetSqlDouble(reader, 7);
                            if ((pars.Naimenovanie.Contains("Резистор") || pars.Naimenovanie.Contains("Конденсатор")))
                            {
                                amountByRemark = CalculateAmountByRemark(GetSqlString(reader, 9));
                                if (GetSqlDouble(reader, 7) > amountByRemark)
                                    amountByRemark = GetSqlDouble(reader, 7);
                            }
                            double countParent = GetSqlDouble(reader, 8) / GetSqlDouble(reader, 7);

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                            GetSqlString(reader, 13),
                                                            GetSqlString(reader, 9),
                                                            amountByRemark,
                                                            countParent * amountByRemark,
                                                            GetSqlInt32(reader, 11),
                                                            GetSqlString(reader, 12));
                            collectionAllElements.otherItems[collectionAllElements.otherItems.Count - 1].Include.Add(include);
                        }
                        else
                                                if (GetSqlInt32(reader, 2) == classMaterial && filterParameters.Materials)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.materials.Add(pars);
                            logDelegate(String.Format("Добавлен Материал: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = true;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                        GetSqlString(reader, 13),
                                                        GetSqlString(reader, 9),
                                                        GetSqlDouble(reader, 7),
                                                        GetSqlDouble(reader, 8),
                                                        GetSqlInt32(reader, 11),
                                                        GetSqlString(reader, 12));
                            collectionAllElements.materials[collectionAllElements.materials.Count - 1].Include.Add(include);
                        }
                        else
                                                    if (GetSqlInt32(reader, 2) == classComplement && filterParameters.Complement)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.complects.Add(pars);
                            logDelegate(String.Format("Добавлен Комплект: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = true;
                            entrComponentProgram = false;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                   GetSqlString(reader, 13),
                                                   GetSqlString(reader, 9),
                                                   GetSqlDouble(reader, 7),
                                                   GetSqlDouble(reader, 8),
                                                   GetSqlInt32(reader, 11),
                                                   GetSqlString(reader, 12));
                            collectionAllElements.complects[collectionAllElements.complects.Count - 1].Include.Add(include);
                        }
                        else
                                                        if (GetSqlInt32(reader, 2) == classComplexProgram && filterParameters.ComplexProgram)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.complexesPrograms.Add(pars);
                            logDelegate(String.Format("Добавлен Комплекc (программы): {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = false;
                            entrComplexProgram = true;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                GetSqlString(reader, 13),
                                                GetSqlString(reader, 9),
                                                GetSqlDouble(reader, 7),
                                                GetSqlDouble(reader, 8),
                                                GetSqlInt32(reader, 11),
                                                GetSqlString(reader, 12));
                            collectionAllElements.complexesPrograms[collectionAllElements.complexesPrograms.Count - 1].Include.Add(include);
                        }
                        else
                                                            if (GetSqlInt32(reader, 2) == classComponentProgram && filterParameters.ComponentProgram)
                        {
                            pars = WriteParameters(reader);
                            collectionAllElements.complementsPrograms.Add(pars);
                            logDelegate(String.Format("Добавлен Компонент (программы): {0} {1}", pars.Oboznachenie, pars.Naimenovanie));

                            entrDocumentation = false;
                            entrComplex = false;
                            entrAssembly = false;
                            entrDetal = false;
                            entrStandartItem = false;
                            boolOtherItem = false;
                            entrMaterial = false;
                            entrComplect = false;
                            entrComponentProgram = true;
                            entrComplexProgram = false;

                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                             GetSqlString(reader, 13),
                                             GetSqlString(reader, 9),
                                             GetSqlDouble(reader, 7),
                                             GetSqlDouble(reader, 8),
                                             GetSqlInt32(reader, 11),
                                             GetSqlString(reader, 12));
                            collectionAllElements.complementsPrograms[collectionAllElements.complementsPrograms.Count - 1].Include.Add(include);
                        }
                    }
                    else
                    {
                        double amountByRemark = GetSqlDouble(reader, 7);
                        if (GetSqlInt32(reader, 2) == classOtherItems)
                            if ((GetSqlString(reader, 4).Contains("Резистор") || GetSqlString(reader, 4).Contains("Конденсатор")))
                            {
                                amountByRemark = CalculateAmountByRemark(GetSqlString(reader, 9));
                                if (GetSqlDouble(reader, 7) > amountByRemark)
                                    amountByRemark = GetSqlDouble(reader, 7);
                            }

                        double countParent = GetSqlDouble(reader, 8) / GetSqlDouble(reader, 7);

                        var include = new SpecListParams.Entrances(GetSqlString(reader, 5), GetSqlString(reader, 10),
                                                        GetSqlString(reader, 13),
                                                        GetSqlString(reader, 9),
                                                        amountByRemark,
                                                        countParent * amountByRemark,
                                                        GetSqlInt32(reader, 11),
                                                        GetSqlString(reader, 12));


                        if (entrDocumentation && GetSqlInt32(reader, 2) == classDocumentation)
                            collectionAllElements.documentation[collectionAllElements.documentation.Count - 1].Include.Add(include);
                        else if (entrComplex && GetSqlInt32(reader, 2) == classComplex)
                            collectionAllElements.complexes[collectionAllElements.complexes.Count - 1].Include.Add(include);
                        else if (entrAssembly && GetSqlInt32(reader, 2) == classAssembly)
                            collectionAllElements.assemblies[collectionAllElements.assemblies.Count - 1].Include.Add(include);
                        else if (entrDetal && GetSqlInt32(reader, 2) == classDetal)
                            collectionAllElements.detalies[collectionAllElements.detalies.Count - 1].Include.Add(include);
                        else if (entrStandartItem && GetSqlInt32(reader, 2) == classStandartItems)
                            collectionAllElements.standartItems[collectionAllElements.standartItems.Count - 1].Include.Add(include);
                        else if (boolOtherItem && GetSqlInt32(reader, 2) == classOtherItems)
                            collectionAllElements.otherItems[collectionAllElements.otherItems.Count - 1].Include.Add(include);
                        else if (entrMaterial && GetSqlInt32(reader, 2) == classMaterial)
                            collectionAllElements.materials[collectionAllElements.materials.Count - 1].Include.Add(
                                include);
                        else if (entrComplect && GetSqlInt32(reader, 2) == classComplement)
                            collectionAllElements.complects[collectionAllElements.complects.Count - 1].Include.Add(include);
                        else if (entrComponentProgram && GetSqlInt32(reader, 2) == classComponentProgram)
                            collectionAllElements.complementsPrograms[collectionAllElements.complementsPrograms.Count - 1].Include.Add(include);
                        else if (entrComplexProgram && GetSqlInt32(reader, 2) == classComplexProgram)
                            collectionAllElements.complexesPrograms[collectionAllElements.complexesPrograms.Count - 1].Include.Add(include);

                    }

                    objectId = GetSqlInt32(reader, 0);
                }
            }


            reader.Close();
            #endregion


            conn.Close();

            return collectionAllElements;
        }

        private static SpecListParams WriteParameters(SqlDataReader reader)
        {
            SpecListParams pars = new SpecListParams(GetSqlInt32(reader, 0));
            pars.ParentDocID = GetSqlInt32(reader, 1);
            pars.Oboznachenie = GetSqlString(reader, 3);
            pars.Naimenovanie = GetSqlString(reader, 4);
            pars.Format = GetSqlString(reader, 6);
        //    pars.UnitMeasure = GetSqlString(reader, 13);
            pars.VhoditVSostav = true;
            return pars;
        }

        public int CalculateAmountByRemark(string remark)
        {
            int amount = 0;
            string threeDotsPattern = @"[RCС]\d{1,3}\.{2,3}[RCС]\d{1,3}";
            string digitsPattern = @"\d{1,3}";
            string itemsPattern = @"[RCС]\d{1,3}";

            // Вычисление количества в структуре типа RX..RZ amount = (Z-X)+1
            var threeDotsFragments = Regex.Matches(remark, threeDotsPattern);
            foreach (Match match in threeDotsFragments)
            {
                var digitsFragments = Regex.Matches(match.Value, digitsPattern);
                if (Convert.ToInt32(digitsFragments[0].Value) < Convert.ToInt32(digitsFragments[1].Value))
                    amount += Convert.ToInt32(digitsFragments[1].Value) - Convert.ToInt32(digitsFragments[0].Value) + 1;
                else
                    amount += Convert.ToInt32(digitsFragments[0].Value) - Convert.ToInt32(digitsFragments[1].Value) + 1;
                remark = remark.Remove(remark.IndexOf(match.Value), match.Value.Length);
            }

            // Вычисление количества одиночных элементов
            amount += Regex.Matches(remark, itemsPattern).Count;

            return amount;
        }

        private static string GetSqlString(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return String.Empty;

            return reader.GetString(field);
        }

        private static int GetSqlInt32(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0;

            return reader.GetInt32(field);
        }

        private static double GetSqlDouble(SqlDataReader reader, int field)
        {
            if (reader.IsDBNull(field))
                return 0d;

            return reader.GetDouble(field);
        }

        public int DataListCategoryToExcel(Xls xls, List<SpecListParams> listCategory, string category, int row, int col, LogDelegate logDelegate, ReportParameters reportParameters)
        {        
            // запись элементов в отчет
            logDelegate("Запись всех элементов раздела спецификации - " + category);

            if (listCategory.Count > 0)
            {
                //заголовок раздела спецификации
                row++;
                int row1 = 0;
                xls[1, row].SetValue(category);
                xls[1, row].Font.Bold = true;
                xls[1, row].Font.Underline = true;
                xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                xls[1, row, col, 1].Merge();
                xls[1, row - 1, col, 2].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                  XlBordersIndex.xlEdgeBottom,
                                                                                                  XlBordersIndex.xlEdgeLeft,
                                                                                                  XlBordersIndex.xlEdgeRight,
                                                                                                  XlBordersIndex.xlInsideVertical);
                int lenthTable = col;

                foreach (SpecListParams param in listCategory)
                {
                    int colEntrance = 0;
                    var colCountOnKnot = 0;
                    var colCountOnProduct = 0;
                    row1 = row;
                    row++;
                    col = 1;
                    col = (reportParameters.Positions ) ? col + 1 : col + 0;
                    col = (reportParameters.Zone) ? col + 1 : col + 0; 


                    col = InsertParameter(xls, col, row, reportParameters.Format, param.Format);
                    col = InsertParameter(xls, col, row, reportParameters.Denotation, param.Oboznachenie);
                    col = InsertParameter(xls, col, row, reportParameters.Name, param.Naimenovanie);

                    int mainCol = col;
                    //запись первого вхождения элемента в отчет
                    colEntrance = col;
                    col = 1;
                    col = InsertParameter(xls, col, row, reportParameters.Positions, param.Include[0].Position);
                    col = InsertParameter(xls, col, row, reportParameters.Zone, param.Include[0].Zone);
                    colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.Remarks, param.Include[0].Remarks);
                    colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.UnitMeasure, param.Include[0].UnitMeasure);

                    if (param.Include[0].Obozna4enie != BaseDocument.Oboznachenie)
                    {
                        //входимость  
                        if (reportParameters.EntranceDenotation && reportParameters.EntranceName)
                        {
                            colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.EntranceDenotation && reportParameters.EntranceName, param.Include[0].Obozna4enie + " " + param.Include[0].Naimenovanie);
                        }
                        else
                        {
                            colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.EntranceDenotation, param.Include[0].Obozna4enie);
                            colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.EntranceName, param.Include[0].Naimenovanie);
                        }
                    }
                    else
                    {
                        colEntrance = (reportParameters.EntranceDenotation || reportParameters.EntranceName) ? colEntrance + 1 : colEntrance + 0;

                    }

                    //количество на узел
                    if (param.Include[0].Count != 0)
                    {
                        if (reportParameters.CountOnKnot)
                            colCountOnKnot = colEntrance;
                        colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.CountOnKnot, param.Include[0].Count);
                    }
                    //количество на изделие
                    if (param.Include[0].Count != 0)
                    {
                        if (reportParameters.CountOnProduct)
                            colCountOnProduct = colEntrance;
                        colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.CountOnProduct, param.Include[0].TotalCount);
                    }

                    //  xls[1, row].SetValue(param.Format);
                    //   xls[1, row].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    //обозначение

                    //  xls[2, row].SetValue(param.Oboznachenie);
                    //наименование
                    //    xls[3, row].SetValue(param.Naimenovanie);

                    //запись первого вхождения элемента в отчет
                    //if (param.Include[0].Obozna4enie != BaseDocument.Oboznachenie)
                    //{
                    //    //входимость  
                    //    if (reportParameters.EntranceDenotation)
                    //        xls[4, row].SetValue(param.Include[0].Obozna4enie);
                    //    //входимость  
                    //    if (reportParameters.EntranceName)
                    //        xls[4, row].SetValue(param.Include[0].Naimenovanie);
                    //    //входимость  
                    //    if (reportParameters.EntranceDenotation && reportParameters.EntranceName)
                    //        xls[4, row].SetValue(param.Include[0].Naimenovanie + " " + param.Include[0].Obozna4enie);
                    //}
                    ////количество на узел
                    //if (param.Include[0].Count != 0)
                    //    xls[5, row].SetValue(param.Include[0].Count);
                    ////количество на изделие
                    //if (param.Include[0].Count != 0)
                    //    xls[6, row].SetValue(param.Include[0].TotalCount);
                    double IncludeCount = 0;
                    double SummCommonCount = 0;

                    IncludeCount = param.Include.Count;
                    SummCommonCount = param.Include[0].TotalCount;

                    //заносит все связанные с головным документы в очередь
                    int indx = 0;

                    for (int i = 1; i < IncludeCount; i++)
                    {
                        indx = i;
                        colEntrance = mainCol; 
                        //row++;
                        ////входимость
                        //if (param.Include[i].Obozna4enie != BaseDocument.Oboznachenie)
                        //    //входимость
                        //    if (reportParameters.AddLikeName.Count > 0)
                        //        xls[4, row].SetValue(param.Include[i].Obozna4enie + " " + param.Include[0].Naimenovanie);
                        //    else
                        //    {
                        //        //входимость  
                        //        if (reportParameters.EntranceDenotation)
                        //            xls[4, row].SetValue(param.Include[i].Obozna4enie);
                        //        //входимость  
                        //        if (reportParameters.EntranceName)
                        //            xls[4, row].SetValue(param.Include[i].Naimenovanie);
                        //        //входимость  
                        //        if (reportParameters.EntranceDenotation && reportParameters.EntranceName)
                        //            xls[4, row].SetValue(param.Include[i].Naimenovanie + " " + param.Include[i].Obozna4enie);
                        //    }

                        row++;
                        col = 1;
                        col = InsertParameter(xls, col, row, reportParameters.Positions, param.Include[i].Position);
                        col = InsertParameter(xls, col, row, reportParameters.Zone, param.Include[i].Zone);
                        colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.Remarks, param.Include[i].Remarks);
                        colEntrance = InsertParameter(xls, colEntrance, row, reportParameters.UnitMeasure, param.Include[i].UnitMeasure);
                        //входимость
                        if (param.Include[i].Obozna4enie != BaseDocument.Oboznachenie)
                            //входимость
                            if (reportParameters.AddLikeName.Count > 0)
                            {
                                if (reportParameters.EntranceDenotation && reportParameters.EntranceName)
                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                        reportParameters.EntranceDenotation && reportParameters.EntranceName, param.Include[i].Obozna4enie + " " + param.Include[i].Naimenovanie);
                                else
                                {
                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                            reportParameters.EntranceDenotation, param.Include[i].Obozna4enie);

                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                        reportParameters.EntranceName, param.Include[i].Naimenovanie);
                                }
                            }
                            else
                            {
                                //входимость  
                                if (reportParameters.EntranceDenotation && reportParameters.EntranceName)
                                {
                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                        reportParameters.EntranceDenotation && reportParameters.EntranceName, param.Include[i].Obozna4enie + " " + param.Include[i].Naimenovanie);
                                }
                                else
                                {
                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                            reportParameters.EntranceDenotation, param.Include[i].Obozna4enie);
                                    colEntrance = InsertParameter(xls, colEntrance, row,
                                            reportParameters.EntranceName, param.Include[i].Naimenovanie);
                                }
                            }

                        //количество на узел
                        if (param.Include[i].Count != 0)
                        {
                            colEntrance = InsertParameter(xls, colEntrance, row,
                                       reportParameters.CountOnKnot, param.Include[i].Count);
                            colCountOnKnot = colEntrance;
                        }
                        //количество на изделие
                        if (param.Include[i].Count != 0)
                        {
                            colEntrance = InsertParameter(xls, colEntrance, row,
                                      reportParameters.CountOnProduct, param.Include[i].TotalCount);
                            colCountOnProduct = colEntrance - 1;
                        }

                        //если изделие входит в состав головного изделия - подсчитываем общее число изделий
                        if (param.VhoditVSostav)
                        {
                            SummCommonCount += param.Include[i].TotalCount;
                        }

                        //добавляем строку с общим количеством для изделия являющихся частью головного
                        if ((i + 1 == IncludeCount) && param.VhoditVSostav && (SummCommonCount != 0) && category != "Документация")
                        {
                            row++;
                            xls[colCountOnProduct, row].NumberFormat = "@";
                            xls[colCountOnProduct, row].SetValue(SummCommonCount);
                            xls[colCountOnProduct, row - 1].Font.Underline = true;
                        }
                    }
                    xls[1, row1 + 1, lenthTable, row - row1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                            XlBordersIndex.xlEdgeBottom,
                                                                                                            XlBordersIndex.xlEdgeLeft,
                                                                                                            XlBordersIndex.xlEdgeRight,
                                                                                                            XlBordersIndex.xlInsideVertical);
                }
            }
            return row;
        }

        private int InsertParameter(Xls xls, int col, int row, bool typeExist, string parameter)
        {
            if (typeExist)
            {
                xls[col, row].SetValue(parameter.ToString());
                col++;
            }
            return col;
        }

        private int InsertParameter(Xls xls, int col, int row, bool typeExist, double parameter)
        {
            if (typeExist)
            {
                if (parameter != 0)
                {
                    xls[col, row].NumberFormat = "@";
                    xls[col, row].SetValue(parameter.ToString());
                }
                col++;
            }
            return col;
        }

        public static int InsertHeader(bool isCheck, string header, int col, int row, Xls xls)
        {
            if (isCheck)
            {
                xls[col, row].Font.Name = "Calibri";
                xls[col, row].Font.Size = 11;
                xls[col, row].SetValue(header.ToString());
                xls[col, row].Font.Bold = true;
                col++;
            }
            return col;
        }


        public void MakeEntranceReport(Xls xls, LogDelegate logDelegate, CollectionAllElements collectionAllElements, ReportParameters reportParameters)
        {
            int col = 1;
            int row = 1;
            col = InsertHeader(reportParameters.Positions, "Позиция", col, row, xls);
            col = InsertHeader(reportParameters.Zone, "Зона", col, row, xls);
            col = InsertHeader(reportParameters.Format, "Формат", col, row, xls);
            col = InsertHeader(reportParameters.Denotation, "Обозначение", col, row, xls);

            col = InsertHeader(reportParameters.Name, "Наименование", col, row, xls);
            col = InsertHeader(reportParameters.Remarks, "Примечание", col, row, xls);
            col = InsertHeader(reportParameters.UnitMeasure, "Единица измерения", col, row, xls);
            col = InsertHeader(reportParameters.EntranceDenotation || reportParameters.EntranceName, "Входимость", col, row, xls);
            col = InsertHeader(reportParameters.CountOnKnot, "Количество на узел", col, row, xls);
            col = InsertHeader(reportParameters.CountOnProduct, "Количество на изделие", col, row, xls);

            col--;

            // формирование заголовков отчета сводной ведомости
            //xls[1, 1].SetValue("Формат");
            //xls[1, 1].Font.Bold = true;
            //xls[2, 1].SetValue("Обозначение");
            //xls[2, 1].Font.Bold = true;
            //xls[3, 1].SetValue("Наименование");
            //xls[3, 1].Font.Bold = true;
            //xls[4, 1].SetValue("Входимость");
            //xls[4, 1].Font.Bold = true;
            //xls[5, 1].SetValue("Количество на узел");
            //xls[5, 1].Font.Bold = true;
            //xls[6, 1].SetValue("Количество на изделие");
            //xls[6, 1].Font.Bold = true;

            xls[1, 1, col, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                        XlBordersIndex.xlEdgeBottom,
                                                                                        XlBordersIndex.xlEdgeLeft,
                                                                                        XlBordersIndex.xlEdgeRight,
                                                                                        XlBordersIndex.xlInsideVertical);
           
            xls[1, 2].SetValue(BaseDocument.Oboznachenie + " " + BaseDocument.Naimenovanie);
            xls[1, 2].Font.Bold = true;
            xls[1, 2].Font.Underline = true;
            xls[1, 2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 2, col, 1].Merge();
            logDelegate("Заполняется шапка документа");
            row = 2;
            row = DataListCategoryToExcel(xls, collectionAllElements.documentation, "Документация", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.complexes, "Комплексы", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.assemblies, "Сборочные единицы", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.detalies, "Детали", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.standartItems, "Стандартные изделия", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.otherItems, "Прочие изделия", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.materials, "Материалы", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.complects, "Комплекты", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.complementsPrograms, "Компоненты (программы)", row, col, logDelegate, reportParameters);
            row = DataListCategoryToExcel(xls, collectionAllElements.complexesPrograms, "Комплексы (программы)", row, col, logDelegate, reportParameters);
        }
    }

    public static class TFDClass
    {
        public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
        public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
        public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
        public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
        public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
        public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
        public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
        public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find(new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987")).Id;
        public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
        public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;
    }
}
