using System;
using System.Collections.Generic;
using System.Linq;
using Globus.DOCs.Technology.Reports;
using Globus.DOCs.Technology.Reports.CAD;
using TFlex;
using TFlex.DOCs.Model.References.Files;
using TFlex.Drawing;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Reporting;
using System.Windows.Forms;
using System.Data.SqlClient;
using TFlex.DOCs.Model.References.Nomenclature;
using System.Text.RegularExpressions;
using Application = TFlex.Application;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;

namespace TechnologicalDocumentationReport
{
    public class CadGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".grb";
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

                if(selectionForm.ShowDialog() == DialogResult.OK)
                {
                    TechnologicalDocumentationReport.MakeReport(context, selectionForm.reportParameters);
                }

                //string str = @" taskkill /IM TFlex.DOCs.FilePreview.Provider.exe /F";
                //ExecuteCommand(str);

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                //if (document != null)
                // document.CloseDocument();
            }
        }

    }

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);

    public class ReportParameters
    {
        // список удаляемых объектов из состава объекта отчета
        public List<int> listDeleteObjectId;

        public List<int> listObjectId;

        // параметр добавление условия содержит наименование
        public List<string> AddLikeName = new List<string>();

        // параметр добавление условия содержит обозначение
        public List<string> AddLikeDenotation = new List<string>();

        public string NameDocument;        // Наименование документа
        public string DenotationDocument;  // Обозначение документа
        public string LiteraDocument;      // Литера документа
        public string Author;              // Разработчик
        public string NControl;            // Нормоконтроль
        public string Denotation;          // Обозначение
        public string Checkman1;          // Проверил1
        public string Checkman2;          // Проверил2
        public string NameBoss;          // Наименование должности начальника
        public string Head;               // Шапка документа

        public bool AddUpak;

        public List<TFDDocument> ListTTP = new List<TFDDocument>();           //Перечень ТТП
        public List<TFDDocument> ListComplement = new List<TFDDocument>();    //Комплекты
    }

    public class TFDDocument
    {
        // Параметры объекта

        public int ObjectID;

        public int Class;
        public double Amount;

        public string Naimenovanie;
        public string Denotation;


    }

    public class ApiDocs
    {
        public static SqlConnection GetConnection()
        {
            SqlConnection conn;
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            sqlConStringBuilder.DataSource = nameSqlServer;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";
            // string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=S2"; //SRV1";
            conn = new SqlConnection(sqlConStringBuilder.ToString());
            conn.Open();
            return conn;
        }
    }


    public class TechnologicalDocumentationReport : IDisposable
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, ReportParameters reportParameters)
        {
            TechnologicalDocumentationReport report = new TechnologicalDocumentationReport();
            report.Make(context, reportParameters);
        }

        private IReportGenerationContext _context;

        public void Make(IReportGenerationContext context, ReportParameters reportParameters)
        {
            _context = context;
            context.CopyTemplateFile();    // Создаем копию шаблона


            m_form.Visible = true;
            LogDelegate m_writeToLog;
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.setStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            List<TFDDocument> reportData = new List<TFDDocument>();
            List<TFDDocument> sortedReportData;

            try
            {
                using (TechnologicalDocumentationReport report = new TechnologicalDocumentationReport())
                {
                    //int docID = report.GetDocsDocumentID(context);
                    if (reportParameters.listObjectId.Count == 0)
                    {
                        MessageBox.Show("Выберите объект для формирования отчета");
                        return;
                    }

                    m_form.setStage(IndicationForm.Stages.DataAcquisition);
                    m_writeToLog("=== Получение данных ===");

                    //Чтение данных для отчета    
                    foreach (var id in reportParameters.listObjectId)
                    {
                        reportData.AddRange(report.ReadData(id, context, reportParameters, m_writeToLog));
                    }


                    if (reportData.Count == 1 || reportData.Count == 0)
                    {
                        MessageBox.Show("Объект выбранный для формирования отчета не имеет состава или выборка с дополнительными условиями возвращает нулевой состав.", "Внимание!",
                                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        this.m_form.Close();
                        return;
                        // Environment.Exit(0);
                    }
                    m_form.setStage(IndicationForm.Stages.DataProcessing);
                    m_writeToLog("=== Обработка данных ===");

                    sortedReportData = SumDublicates(SortData(reportData));

                    m_form.setStage(IndicationForm.Stages.ReportGenerating);
                    m_writeToLog("=== Формирование отчета ===");

                    //Формирование отчета
                    MakeSelectionByTypeReport(sortedReportData, reportParameters, m_writeToLog);

                    m_form.progressBarParam(2);

                    m_form.setStage(IndicationForm.Stages.Done);
                    m_writeToLog("=== Завершение работы ===");
                    System.Diagnostics.Process.Start(context.ReportFilePath);


                }
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






        // Чтение данных для отчета
        public List<TFDDocument> ReadData(int documentID, IReportGenerationContext context, ReportParameters reportParameters, LogDelegate logDelegate)
        {
            // соединяемся с БД T-FLEX DOCs 2010
            var conn = ApiDocs.GetConnection();

            List<TFDDocument> objects = new List<TFDDocument>();

            // Формирование списка объектов, которые исключаются из состава объекта для формирования отчета
            string listDelObj = string.Empty;

            if (reportParameters.listDeleteObjectId.Count > 1)
            {
                for (int i = 0; i < reportParameters.listDeleteObjectId.Count - 1; i++)
                {
                    listDelObj += reportParameters.listDeleteObjectId[i] + ",";
                }
            }
            listDelObj += reportParameters.listDeleteObjectId[reportParameters.listDeleteObjectId.Count - 1].ToString();



            #region запрос для отчета с разузловкой
            //запрос для отчета с разузловкой

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
                                                        Position INT,
                                                        Zone NVARCHAR(20),
                                                        Format NVARCHAR(20),
                                                        Denotation NVARCHAR(255),
                                                        Name NVARCHAR(255),
                                                        Amount FLOAT,
                                                        Remarks NVARCHAR(255),
                                                        letterEx NVARCHAR(50),
                                                        TotalCount FLOAT)
                                ELSE DELETE FROM #TmpVspec

                                INSERT INTO #TmpVspec
                                SELECT n.s_ObjectID,0,0,n.s_ClassID,0,'','',n.Denotation,n.Name,0,'',n.letterEx,1
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
                                         nh.Position,
                                         nh.Zone,
                                         n.Format,
                                         n.Denotation,
                                         n.Name,
                                         nh.Amount,
                                         nh.Remarks,
                                         n.letterEx,
                                         #TmpVspec.TotalCount*nh.Amount
                                  FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                           
                                       INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                                  WHERE
                                      #TmpVspec.[level]=@level
                                      AND nh.s_ActualVersion = 1
                                      AND nh.s_Deleted = 0
                                      AND n.s_ActualVersion = 1
                                      AND n.s_Deleted = 0
                                      AND n.s_ObjectID NOT IN ({1})
                                      AND n.s_ClassID IN (499,500, 502, 501)
                                  
                                  SET @insertcount = @@ROWCOUNT 
                                  SET @level = @level + 1 
   
                                  IF @insertcount = 0 
                                  GOTO end1

                                END
                                end1:

                                   SELECT tv.s_ObjectID,  
                                      
                                       tv.s_ClassID, 
                                       
                                       tv.Denotation, 
                                       tv.Name 
                                      
                                      
                          
                                FROM #TmpVspec tv 
                                ----AddWhereDenotation
                                --AddWhereName
                                GROUP BY   
                                         
                                         tv.s_ClassID, 
                                         
                                         tv.Denotation, 
                                         tv.Name ,
                                         tv.s_ObjectID
                                         ", documentID, listDelObj);

            int indx = 0;
            if (reportParameters.AddLikeName.Count > 0)
            {
                indx = getDocTreeCommandText.IndexOf("----") - 1;
                getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "WHERE");
            }

            for (int i = 0; i < reportParameters.AddLikeName.Count; i++)
            {
                indx = getDocTreeCommandText.IndexOf("----") - 1;


                if (reportParameters.AddLikeDenotation[i] != null && reportParameters.AddLikeDenotation[i] != string.Empty)
                {
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                          string.Format(
                                                                              @" (Denotation Like N'%{0}%' ",
                                                                              reportParameters.AddLikeDenotation[i]));


                    if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" AND Name Like N'%{0}%') ",
                                                                                  reportParameters.AddLikeName[i]));
                    }
                    else
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx, ")");
                    }
                }
                else
                {
                    if (reportParameters.AddLikeName[i] != null && reportParameters.AddLikeName[i] != string.Empty)
                    {
                        indx = getDocTreeCommandText.IndexOf("----") - 1;
                        getDocTreeCommandText = getDocTreeCommandText.Insert(indx,
                                                                              string.Format(
                                                                                  @" Name Like N'%{0}%' ",
                                                                                  reportParameters.AddLikeName[i]));
                    }
                }
                if (i < reportParameters.AddLikeName.Count - 1)
                {
                    indx = getDocTreeCommandText.IndexOf("----") - 1;
                    getDocTreeCommandText = getDocTreeCommandText.Insert(indx, "OR");
                }
            }


            var docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
            docTreeCommand.CommandTimeout = 0;
            SqlDataReader reader = docTreeCommand.ExecuteReader();

            TFDDocument doc;

            while (reader.Read())
            {
                doc = new TFDDocument();
                doc.ObjectID = GetSqlInt32(reader, 0);

                doc.Class = GetSqlInt32(reader, 1);
                doc.Denotation = GetSqlString(reader, 2);
                doc.Naimenovanie = GetSqlString(reader, 3);

                if (doc.Class == TFDClass.Komplekt && doc.Naimenovanie.Contains("Упаковка") && reportParameters.AddUpak)
                {
                    doc.Class = TFDClass.Assembly;
                }

             

                objects.Add(doc);
                logDelegate(String.Format("Добавлен объект: {0} {1}", doc.Denotation, doc.Naimenovanie));

            }
            m_form.progressBarParam(2);
            reader.Close();
            #endregion


            return objects;
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


        // Группировка элементов по типам с сортировкой внутри групп
        public  List<TFDDocument> SortData(List<TFDDocument> reportData)
        {
            List<TFDDocument> sortedReportData = new List<TFDDocument>();


            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Assembly));
            sortedReportData.AddRange(SortOneTypeObjects(reportData, TFDClass.Detal));


            return sortedReportData;
        }

        //Сортировка внутри группы объектов одного и того же типа
        public  List<TFDDocument> SortOneTypeObjects(List<TFDDocument> reportData, int type)
        {
            List<TFDDocument> oneTypeObjects = new List<TFDDocument>();
            foreach (TFDDocument doc in reportData)
                if (doc.Class == type)
                    oneTypeObjects.Add(doc);

            oneTypeObjects.Sort(new NaimenAndOboznachComparer()); //сортируем записи по обозначению и наименованию
            return oneTypeObjects;
        }

        public void MakeSelectionByTypeReport(List<TFDDocument> reportData, ReportParameters reportParameters, LogDelegate logDelegate)
        {
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

            var doc = Application.OpenDocument(_context.ReportFilePath);

            //var titul = doc.GetPages().First();
            //var secondPage = doc.GetPages().Skip(1).First();

            doc.BeginChanges("Заполнение данными");
            var reportText = doc.GetTextByName("REP_CONTENT");
            reportText.BeginEdit();
            var contentTable = reportText.GetFirstTable();

            logDelegate("Количество элементов отчета: " + reportData.Count);

            if (contentTable.HasValue)
            {
                var cadTable = new CadTable(contentTable.Value, reportText, 52, 61);

                var assemblies = reportData.Where(t => t.Class == TFDClass.Assembly).ToList();
                logDelegate("Количество сборочных единиц: " + assemblies.Count);
                if (assemblies.Count > 0)
                {
                    FillSection(cadTable, "Сборочные единицы", assemblies);
                }

                var details = reportData.Where(t => t.Class == TFDClass.Detal).ToList();
                logDelegate("Количество деталей: " + details.Count);
                if (details.Count > 0)
                {
                    cadTable.NewPage();
                    FillSection(cadTable, "Детали", details);
                
                }
               
                cadTable.NewPage();
                FillSection(cadTable, "Перечень ТТП", reportParameters.ListTTP);
                cadTable.CreateRow();
                FillSection(cadTable, "Комплекты", reportParameters.ListComplement);
                //MessageBox.Show("Сделал!");
                cadTable.Apply();
            }
            reportText.EndEdit();


            // вставка листа регистрации изменений
            var changelogPage = new Page(doc);
            changelogPage.Rectangle = new Rectangle(0,0,210,297);
            // временная страница для удаления фрагментов
            var tempPage = new Page(doc);
            foreach (Fragment fragment in changelogPage.GetFragments())
            {
                fragment.Page = tempPage;
            }
            doc.DeletePage(tempPage, new DeleteOptions(true) { DeletePageObjects = true });

            var fileReference = new FileReference();

            var fileObject = fileReference.FindByRelativePath(
                @"Служебные файлы\Шаблоны технологической документации Глобус\Лист регистрации изменений ГОСТ 2.503.grb");
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента
            var changelogFragment = new Fragment(new FileLink(doc, fileObject.LocalPath, true));
            changelogFragment.Page = changelogPage;


            var pages = new List<Page>();
            foreach (Page page in doc.GetPages())
            {
                pages.Add(page);
            }

            //foreach (var variable in doc.GetVariables())
            //{
            //    switch (variable.Name)
            //    {
            //        case "$graph_04":
            //            variable.Expression = "\"" + reportParameters.DenotationDocument + "\"";
            //            break;
            //    }
            //}

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
                            case "$graph_05":
                                if (i == 0)
                                {
                                    variable.TextValue = reportParameters.LiteraDocument;
                                }
                                break;
                            case "$graph_10":
                                if (i == 1)
                                {
                                    variable.TextValue = reportParameters.LiteraDocument;
                                }
                                break;
                            case "$Razrab_FIO":
                                variable.TextValue = reportParameters.Author;
                                break;
                            case "$NKontr_FIO":
                                variable.TextValue = reportParameters.NControl;
                                break;
                            case "$Prov1_FIO":
                                variable.TextValue = reportParameters.Checkman1;
                                break;
                            case "$Prov2_FIO":
                                variable.TextValue = reportParameters.Checkman2;
                                break;
                            case "$Head":
                                variable.TextValue = reportParameters.Head;
                                break;
                            case "$Boss":
                                variable.TextValue = reportParameters.NameBoss;
                                break;
                            case "$Denotation":
                                variable.TextValue = reportParameters.Denotation;
                                break;
                            case "$oboznach":
                                variable.TextValue = reportParameters.Denotation;
                                break;
                            case "$graph_04":
                                variable.TextValue = reportParameters.DenotationDocument;
                                break;
                            case "$name":
                                variable.TextValue = reportParameters.NameDocument;
                                break;
                        }
                    }
                }
            }




            doc.EndChanges();

            //doc.BeginChanges("Изменение переменных");

            //foreach (var variable in doc.GetVariables())
            //{
            //    MessageBox.Show(variable.Name);
            //    switch (variable.Name)
            //    {
            //        case "$Razrab_FIO":
            //            variable.Expression = reportParameters.Author;
            //            break;
            //        case "$NKontr_FIO":
            //            variable.Expression = reportParameters.NControl;
            //            break;
            //    }
            //}



            //doc.EndChanges();
            doc.Save();
            doc.Close();
        }

        public class ReportItem
        {
            public string denotation; // базовое обозначение
            public string naimenovanie; // наименование
            public string stringPartMake = string.Empty; // строковая часть исполнения (Например: Сп - ЫК6.490.099-17 Сп)
            public string denotationWithOutStringPartMake;

            public List<int> variants = new List<int>(); // исполнения
            public List<int> variantsStr = new List<int>(); // исполнения строковые
            private Regex regexStringPartMake = new Regex(@"(?<=-\d{1,3}\s)\D{1,}");
            private Regex regexNumberMake = new Regex(@"(?<=-)\d{1,3}");
            // создание элемента отчета
            public ReportItem(TFDDocument item)
            {

                naimenovanie = item.Naimenovanie;
                var index = item.Denotation.LastIndexOf('-');
                if (index < 0)
                {
                    denotation = item.Denotation;
                    variants.Add(0);
                }
                else
                {
                    if (regexStringPartMake.Match(item.Denotation).Value.Trim() != string.Empty)
                    {
                        denotation = item.Denotation.Substring(0, index) + " " +
                                     regexStringPartMake.Match(item.Denotation).Value.Trim();
                        stringPartMake = regexStringPartMake.Match(item.Denotation).Value.Trim();
                       
                        denotationWithOutStringPartMake = item.Denotation.Substring(0, index);
                    }
                    else
                        denotation = item.Denotation.Substring(0, index);
                    try
                    {
                        variants.Add(int.Parse(item.Denotation.Substring(index + 1).Trim()));
                    }
                    catch
                    {
                      variants.Add(int.Parse(regexNumberMake.Match(item.Denotation).Value));
                    }

                }
            }

            public string Naimenovanie
            {
                get
                {
                    return naimenovanie;
                }
            }

            // формируемое обозначение с учетом группировки исполнений
            public string Denotation
            {
                get
                {

                   variants = variants.Distinct().OrderBy(t => t).ToList();
                        //variants.Sort();

                    var newDenotation = string.Empty;
                    if (stringPartMake == string.Empty)
                    {
                        newDenotation = denotation;
                        if (variants.Count == 2)
                        {
                            if (variants[1] == 1)
                                newDenotation += ", ";
                        }
                    }
                    else
                    {
                        newDenotation = denotationWithOutStringPartMake;

                        if (variants[0] == 0)
                        {
                            newDenotation = denotation;

                            if (variants.Count == 2)
                            {
                                if (variants[1] == 1)
                                    newDenotation += ", ";
                            }

                        }
                    }
                    if (variants.Count == 1)
                        {
                            if (variants[0] != 0)
                            {
                                if (stringPartMake == string.Empty)
                                    newDenotation += "-" + VariantToString(variants[0]);
                                else
                                    newDenotation +=  "-" + VariantToString(variants[0]) + " " + stringPartMake;
                            }
                        }
                        else
                        {
                            // создание цепочек исполнений
                            var variantChains = new List<List<int>>();
                            var lastChain = new List<int>();
                                lastChain.Add(variants[0]);
                                variantChains.Add(lastChain);

                            for (int i = 1; i < variants.Count; i++)
                            {
                              
                                if (variants[i] - lastChain[lastChain.Count - 1] == 0)
                                {
                                   
                                    continue;
                                }

                                if (variants[i] - lastChain[lastChain.Count - 1] == 1)
                                {
                                    lastChain.Add(variants[i]);
                                   
                                }
                                else
                                {
                                    lastChain = new List<int>();
                                    lastChain.Add(variants[i]);
                                    variantChains.Add(lastChain);
                                }
                            }

                            // формирование строки обозначения с учетом цепочек исполнений
                            //if (variants.Count>1)
                            //    MessageBox.Show(string.Join(", ", ", "+variantChains.Select(ChainToString).ToArray()));
                           
                            if (stringPartMake == string.Empty)
                                newDenotation += string.Join(", ", variantChains.Select(ChainToString).ToArray());
                            else
                                newDenotation += string.Join(", ", variantChains.Select(ChainToStringWithPartMake).ToArray());
                          
                        }

                        return newDenotation;
                   
                }
            }

           

            // преобразование цепочки вариантов в строковое представление
            private string ChainToString(List<int> chain)
            {
               
                if (chain.Count == 1)
                {
                    if (chain[0] !=0)
                        return "-" + VariantToString(chain[0]);
                    else return string.Empty;
                }
                if (chain.Count == 2)
                {
                    if (chain[0] != 0)
                    {
                        return "-" + VariantToString(chain[0]) + ", -" + VariantToString(chain[1]);
                    }
                    else
                    {
                        return "-" + VariantToString(chain[1]);
                    }
                }
                if (chain.Count == 3)
                {
                    if (chain[0] != 0)
                    {
                        return "-" + VariantToString(chain[0]) + "...-" + VariantToString(chain[chain.Count - 1]);
                    }
                    else
                    {
                        return "-" + VariantToString(chain[1]) + ", -" + VariantToString(chain[2]);
                    }
                }
               

                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + "...-" + VariantToString(chain[chain.Count - 1]);
                }
                else
                {
                    return "-" + VariantToString(chain[1]) + "...-" + VariantToString(chain[chain.Count - 1]);
                }
            }

            private string ChainToStringWithPartMake(List<int> chain)
            {
                
                if (chain.Count == 1)
                {
                    if (chain[0] != 0)
                        return "-" + VariantToString(chain[0]) +" " + stringPartMake;
                    else return string.Empty;
                }
                if (chain.Count == 2)
                {
                    if (chain[0] != 0)
                    {
                        return "-" + VariantToString(chain[0]) + " " + stringPartMake + ", -" + VariantToString(chain[1]) + " " + stringPartMake;
                    }
                    else
                    {
                        return "-" + VariantToString(chain[1]) + " " + stringPartMake;
                    }
                }
                if (chain.Count == 3)
                {
                    if (chain[0] != 0)
                    {
                        return "-" + VariantToString(chain[0]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
                    }
                    else
                    {
                        return "-" + VariantToString(chain[1]) + " " + stringPartMake + ", -" + VariantToString(chain[2]) + " " + stringPartMake;
                    }
                }


                if (chain[0] != 0)
                {
                    return "-" + VariantToString(chain[0]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
                }
                else
                {
                    return "-" + VariantToString(chain[1]) + " " + stringPartMake + "...-" + VariantToString(chain[chain.Count - 1]) + " " + stringPartMake;
                }
            }


            // преобразование варианта в строкове представление
            private string VariantToString(int variant)
            {
                if (variant > 99)
                {
                    return variant.ToString("000");
                }
                return variant.ToString("00");
            }

            // проверка нового элемента на соответствие текущему (является ли элемент исполнением)
            public bool IsValid(TFDDocument item)
            {
                if (item.Naimenovanie != naimenovanie) // если не соответствует наименование
                {
                    return false;
                }

                var index = item.Denotation.LastIndexOf('-');
                var itemDenotationBase = ""; // обозначение без части исполнения

               // var regexStringPartMake = new Regex(@"(?<=-\d{1,3}\s)\D{1,}");
               
                if (index < 0)
                {
                    itemDenotationBase = item.Denotation;
                }
                else
                {
                    if (regexStringPartMake.Match(item.Denotation).Value.Trim() != string.Empty)
                        itemDenotationBase = item.Denotation.Substring(0, index) +" "+ regexStringPartMake.Match(item.Denotation).Value.Trim();
                    else
                        itemDenotationBase = item.Denotation.Substring(0, index);
                }

               

                   if ((itemDenotationBase != denotation) ) // обозначения не соответствуют
                {
                    
                    return false;
                }
                

                if (index >= 0)
                {
                    var variantString = item.Denotation.Substring(index + 1).Trim();
                    //if (variantString.Length == 1)
                    //{
                    //    MessageBox.Show(denotation);
                    //    return false; // если одна цифра исполнения 
                    //}
                    //for (int i = 0; i < variantString.Length; i++) // если часть исполнения содержит не цифры
                    //{
                    //    if (!char.IsDigit(variantString[i]))
                    //    {
                    //        MessageBox.Show(denotation);
                    //        return false;
                    //    }
                    //}
                }

                return true;
            }

            // функция добавления исполнения
            public void Add(TFDDocument item)
            {
                var index = item.Denotation.LastIndexOf('-');
                if (index < 0)
                {
                    variants.Add(0);
                }
                else
                {
                    try
                    {
                        variants.Add(int.Parse(item.Denotation.Substring(index + 1)));
   

                    }
                    catch
                    {
                       // variants.Add(999999);
                        variants.Add(int.Parse(regexNumberMake.Match(item.Denotation).Value));
 
                       // MessageBox.Show("Не удалось добавить вариант " + item.Denotation + " к " + denotation + "!");
                    }
                }
            }
        }

        public static List<string> GetLines(string text, int countSymbolsOnLine)
        {
            var textWords = text.Split(new char[]{' ', '\r', '\n'}, StringSplitOptions.RemoveEmptyEntries);

            var lines = new List<string>();
            var line = "";
            foreach (string word in textWords)
            {
                if (line.Length + word.Length > countSymbolsOnLine)
                {
                    if (line.Length == 0)
                    {
                        lines.Add(word);
                    }
                    else
                    {
                        lines.Add(line.Trim());
                        line = word;
                    }
                }
                else
                {
                    line += " " + word;
                }
            }
            if (!string.IsNullOrEmpty(line))
            {
                lines.Add(line.Trim());
            }
            return lines;
        }

        private void FillSection(CadTable table, string section, List<TFDDocument> items)
        {
            var row = table.CreateRow();

            row.AddText(section, 2, CadTableCellJust.Center, TextStyle.BoldUnderline);
            table.CreateRow();
            table.CreateRow();

            var reportItems = new List<ReportItem>();

            ReportItem lastItem = null;

            //// формирование списка на вывод с учетом группировки по исполнениям
            //foreach (var item in items)
            //{

               
            //    if (lastItem == null) // первый элемент отчета
            //    {
            //        lastItem = new ReportItem(item);
            //        reportItems.Add(lastItem);
            //    }
            //    else
            //    {
            //        // если последний элемент отчета соответствует текущему, то добавляем исполнение
            //        if (lastItem.IsValid(item))
            //        {
                        
            //                lastItem.Add(item);

            //        }
            //        else
            //        {
                       
            //            lastItem = new ReportItem(item);
            //            reportItems.Add(lastItem);
            //        }
            //    }
            //}

            //Избавляемся от суммирования исполнений
            foreach (var tfdDocument in items)
            {
                lastItem = new ReportItem(tfdDocument);
                reportItems.Add(lastItem);
            }
           

            foreach (var item in reportItems)
            {

                // перенос строк если обозначение длиннее 55 и наименование длинеее 45
                var denotationLines = GetLines(item.Denotation, 55);
                var naimenovanieLines = GetLines(item.Naimenovanie, 45);

                var linesCount = Math.Max(denotationLines.Count, naimenovanieLines.Count);

                for (int i = 0; i < linesCount; i++)
                {
                    row = table.CreateRow();
                    if (denotationLines.Count>i)
                    {
                        row.AddText(denotationLines[i]);
                    }
                    else
                    {
                        row.AddText("");
                    }

                    if (naimenovanieLines.Count > i)
                    {
                        row.AddText(naimenovanieLines[i]);
                    }
                    else
                    {
                        row.AddText("");
                    }
                }
            }

        }

        // Сложение количеств повторяющихся элементов отсортированной коллекции
        static private List<TFDDocument> SumDublicates(List<TFDDocument> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ObjectID == data[slaveID].ObjectID)
                    {
                        data[mainID].Amount += data[slaveID].Amount;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
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

public static class TFDClass
{
    public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
    public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
    public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
    public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
    public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
    public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
    public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
    public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find("Материал").Id;
    public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
    public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;

    public static readonly string ComplementName = "Комплекты";
    public static readonly string OtherItemsName = "Прочие изделия";
    public static readonly string StandartItemsName = "Стандартные изделия";
    public static readonly string AssemblyName = "Сборочные единицы";
    public static readonly string DetailName = "Детали";
    public static readonly string MaterialName = "Материалы";
    public static readonly string DocumentName = "Документы";
    public static readonly string ComplexName = "Комплексы";
    public static readonly string ComponentProgramName = "Компоненты (ЕСПД)";
    public static readonly string ComplexProgramName = "Комплексы (ЕСПД)";

    public static string InsertSection(int classId)
    {
        string sectionName = string.Empty;
        if (classId == Komplekt)
            sectionName = ComplementName;
        if (classId == OtherItem)
            sectionName = OtherItemsName;
        if (classId == StandartItem)
            sectionName = StandartItemsName;
        if (classId == Assembly)
            sectionName = AssemblyName;
        if (classId == Detal)
            sectionName = DetailName;
        if (classId == Material)
            sectionName = MaterialName;
        if (classId == Document)
            sectionName = DocumentName;
        if (classId == Complex)
            sectionName = ComplexName;
        if (classId == ComponentProgram)
            sectionName = ComponentProgramName;
        if (classId == ComplexProgram)
            sectionName = ComplexProgramName;

        return sectionName;
    }
}

