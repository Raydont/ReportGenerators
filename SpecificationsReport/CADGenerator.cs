using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Files;
using SpecificationsReport.Vspecnew;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Structure;

namespace SpecificationsReport
{
    public class CADGenerator : IReportGenerator
    {
       
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
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
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, IWin32Window owner)
        {
            return false;
        }

        public Attributes Attr = new Attributes();
        public MainForm MForm;
        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {

            //Инициализация API CAD
            var applicationSessionSetup =
                new ApplicationSessionSetup
                {
                    Enable3D = true,
                    ReadOnly = false,
                    PromptToSaveModifiedDocuments = false,
                    EnableMacros = true,
                    EnableDOCs = true,
                    DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version12,
                    ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
                };

            if (!TFlex.Application.InitSession(applicationSessionSetup))
            {
                return;
            }
            TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD
            context.CopyTemplateFile();  //Создаем копию шаблона
            TFlex.Model.Document document = null;    //Документ CAD 
            MForm = new MainForm();

            LogDelegate mWriteToLog = new LogDelegate(MForm.writeToLog);
       
            try
            {
                Attr.DotsEntry(context, MForm, mWriteToLog);
            }
            catch (Exception e)
            {
               MessageBox.Show(e.Message, "Exception!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (document != null)
                    document.Close();
            }
        }

    }

    /// <summary>
    /// Класс отвечающий за резрешение ссылок на сборки
    /// </summary>
    public class AssemblyResolver
    {
        private List<string> _directories = new List<string>();

        private AssemblyResolver()
        {
            AddDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));            
        }

        private static AssemblyResolver _resolver;

        public static AssemblyResolver Instance
        {
            get
            {
                if (_resolver == null)
                {
                    AppDomain.CurrentDomain.AssemblyResolve += ResolveEventHandler;
                    _resolver = new AssemblyResolver();
                }
                return _resolver;
            }
        }

        internal static Assembly ResolveEventHandler(Object sender, ResolveEventArgs args)
        {
            Assembly result = null;            

            try
            {
                var name = new AssemblyName();
                name = new AssemblyName(args.Name);

                string resolveAsmName = null;

                int index = args.Name.IndexOf(",");
                if (index >= 0)
                    resolveAsmName = args.Name.Substring(0, index);
                else
                    resolveAsmName = args.Name;                

                if (name.CultureInfo != null && name.CultureInfo.TwoLetterISOLanguageName != "iv" &&
                    (resolveAsmName.StartsWith("TFlex.") || resolveAsmName.StartsWith("DevExpress.")))
                {
                    resolveAsmName = Path.Combine(name.CultureInfo.TwoLetterISOLanguageName, resolveAsmName);                    
                }

                foreach (string directory in _resolver._directories)
                {
                    var builder = new StringBuilder(directory);
                    builder.Append(@"\");
                    builder.Append(resolveAsmName);
                    builder.Append(".dll");
                    string path = builder.ToString();
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        var fullName = new AssemblyName();
                        fullName.CodeBase = path;
                        result = AppDomain.CurrentDomain.Load(fullName);
                        break;
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        public ReadOnlyCollection<string> Directories
        {
            get { return _directories.AsReadOnly(); }
        }

        public void AddDirectory(string directoryPath)
        {            
            directoryPath = directoryPath.ToLower();
            if (!_directories.Contains(directoryPath))
                _directories.Add(directoryPath);            
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

    /// <summary>
    /// Класс для передачи аттрибутов из формы в отчет
    /// </summary>
    public class DocumentAttributes
    {
            // для вкладки Добавление изменений
            private bool _addModify = false; 
            private string numberModify = string.Empty;                 // номер изменения
            private string numberDocument = string.Empty; // номер документа
            private string dateModify = string.Empty;           // дата изменения
            private List<int> listPage = new List<int>(); // список страниц
            
            private string listZam = string.Empty;
            private string listNov = string.Empty;
            private List<int> listPage2 = new List<int>();

        public int countReservPages = 0; //Количество резервных страниц

           public List<int> ListPage2
            {
                get { return listPage2; }
                set { listPage2 = value; }
            }

            public string ListNov
            {
                get { return listNov; }
                set { listNov = value; }
            }
            public string ListZam
            {
                get { return listZam; }
                set { listZam = value; }
            }
            public List<int> ListPage
            {
                get { return listPage; }
                set { listPage = value; }
            }

            public string DateModify
            {
                get { return dateModify; }
                set { dateModify = value; }
            }

            public string NumberDocument
            {
                get { return numberDocument; }
                set { numberDocument = value; }
            }

            public string NumberModify
            {
                get { return numberModify; }
                set { numberModify = value; }
            }

            public bool AddModify
            {
                get { return _addModify; }
                set { _addModify = value; }
            }
    }

    //
    // Класс для запуска макроса
    //
    public class Specification
    {
        //======================================================================
        //
        // Точка входа
        //
        //======================================================================
       
        public static void MakeSpecListReport(IReportGenerationContext context,DocumentAttributes DocAttr,MainForm m_form, LogDelegate m_writeToLog)
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
                ReportData.SpecListParams baseDocumentSpecListParams = ReportData.FillData(baseDocumentID,m_writeToLog);

                m_form.setStage(MainForm.Stages.DataProcessing);
                m_writeToLog("=== Обработка данных ===");
                m_form.progressBarParam(2);
                 
                // Переформатирует данные под CAD
                Queue<ReportData.FormatedData> dataForCAD =
                        ReportData.MakeDataForCAD(baseDocumentSpecListParams,
                                   ReportData.m_sboro4nyeEdinicy,
                                   ReportData.m_komplekty,
                                   ReportData.m_vedomosti,
                                   m_writeToLog);
              
                // Создание отчета на основе сформированных данных
                m_form.progressBarParam(2);
                m_form.setStage(MainForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");
                m_form.progressBarParam(2);

                // Создание отчета CAD
                ReportData.MakeCADReport(dataForCAD, baseDocumentSpecListParams,context , DocAttr);

                m_form.setStage(MainForm.Stages.Done);
                m_writeToLog("=== Работа завершена ===");
                m_form.progressBarParam(2);
                m_writeToLog("Ведомость спецификаций сформирована");
                m_form.progressBarParam(2);
                m_form.Close();
            }
            catch (SLBuilderFatalException e)
            {
                string message = String.Format("ФАТАЛЬНАЯ ОШИБКА: {0}\r\nДля разрешения ситуации обратитесь к системному администратору", e.Message);
                if (m_writeToLog != null)
                    m_writeToLog(message);
               MessageBox.Show(message, "Фатальная ошибка",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (SLBuilderException e)
            {
                string message = String.Format("ОШИБКА ПРИЛОЖЕНИЯ: {0}", e.Message);
                if (m_writeToLog != null)
                   m_writeToLog(message);
               MessageBox.Show(message, "Ошибка приложения",
               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception e)
            {
                string message = String.Format("ОШИБКА: {0}", e.Message);
                if (m_writeToLog != null)
                   m_writeToLog(message);
                MessageBox.Show(message, "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Нахождение ID объекта для отчета ВС
        static int Initialize(IReportGenerationContext context, MainForm mForm, LogDelegate m_writeToLog)
        {
            // --------------------------------------------------------------------------
            // Инициализация
            // --------------------------------------------------------------------------

            int documentID;

             mForm.setStage(MainForm.Stages.Initialization);
             m_writeToLog("=== Инициализация ===");
             mForm.progressBarParam(2);
            // Получение ID обрабатываемого документа

             m_writeToLog("Получение идентификатора корневого документа");
             mForm.progressBarParam(2);
             
            // Получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }             
            
            m_writeToLog("Идентификатор корневого документа: " + documentID);
            mForm.progressBarParam(2);
            
            m_writeToLog("Подготовка к загрузке данных из T-FLEX DOCs");
            mForm.progressBarParam(2);
    
            return documentID;
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
        public double TotalCount;
       
        public TFDDocument(int id, int parent, int docClass)
        {
            Id = id;
            ParentId = parent;
            Class = docClass;            
        }
    }

    class ReportData
    {
        public enum RowType
        {
            Undefined,
            Empty,
            Header,
            FirstRecordRow,
            Row,
            RowBeforeSummary,
            Summary
        }       

        static public Queue<FormatedData> MakeDataForCAD(
                  SpecListParams baseDocument,
                  List<SpecListParams> sboro4nyeEdinicy,
                  List<SpecListParams> komplekty,
                  List<SpecListParams> vedomosti,
                  LogDelegate m_writeToLog
              )
        {
           
            Queue<FormatedData> data = new Queue<FormatedData>();

            if (baseDocument != null) //заносим данные о спецификации головного документа
            {
              
                if (baseDocument.Naimenovanie.Length > 0 && baseDocument.Oboznachenie.Length > 0)
                {
                    FormatedData.FormatDataToCad(baseDocument, true, data);
                    data.Enqueue(new FormatedData());
                }
            }
            if (sboro4nyeEdinicy != null) //заносим данные обо всех сборочных единицах
                if (sboro4nyeEdinicy.Count > 0)
                {
                    
                    
                    foreach (SpecListParams param in sboro4nyeEdinicy)
                    {
                        FormatedData.FormatDataToCad(param, false, data);
                    }
                    
                    data.Enqueue(new FormatedData());
                }
            
            if (komplekty != null)//заносим данные обо всех комплектах
                if (komplekty.Count > 0)
                {
                    foreach (SpecListParams param in komplekty)
                    {
                        FormatedData.FormatDataToCad(param, false, data);
                    }
                }

            if (vedomosti != null) //вставляем в хвост данные о документах, имеющих свои ведомости спецификаций
                if (vedomosti.Count > 0)
                {
                    data.Enqueue(new FormatedData());

                    var temp = new FormatedData();
                    temp.Naimenovanie = "Ведомости спецификаций:";
                    data.Enqueue(temp);

                    foreach (SpecListParams param in vedomosti)
                    {
                        temp = new FormatedData();
                        temp.Naimenovanie = param.Naimenovanie;
                        if (param.Oboznachenie.Length > 0)
                            temp.Naimenovanie += " (" + param.Oboznachenie + ")";

                        data.Enqueue(temp);
                    }
                }
            
            return data;
        }

        internal class FormatedData
        {
            // Перенесение данных из формата, использованного для DOCs в формат для вывода в таблицу CAD
            public static void FormatDataToCad(SpecListParams param, bool BaseDoc, Queue<FormatedData> outData)
            {
                var textFormatter = new TextFormatter(GetFont());
                string[] formatName = textFormatter.Wrap(param.Naimenovanie, 97);
                

                var formatedData = new FormatedData();
                formatedData.Obozna4enie = param.Oboznachenie;
                if (param.Include.Count <= 1 && formatName.Length > 1)
                {
                    formatedData.Naimenovanie = string.Join("\n", formatName);
                }
                else
                    formatedData.Naimenovanie = formatName[0];// string.Join("\n", formatName);
             
                //для главного документа не обрабатываем связи с другими документами
                if (param.Include.Count < 1)
                {
                    if (BaseDoc)
                        formatedData.TotalCount = param.Count.ToString();          
                    outData.Enqueue(formatedData);
                    return;
                }

                //заносим вхождения               
                formatedData.SecondObozna4enie = param.Include[0].Obozna4enie;
                formatedData.Count = param.Include[0].Count.ToString();
                formatedData.TotalCount = param.Include[0].TotalCount.ToString();
                formatedData.Prime4anie = String.Empty;//param.Prime4anie;
                formatedData.UnderlinedCommonCount = false;
                outData.Enqueue(formatedData);
                int IncludeCount = param.Include.Count;
                int SummCommonCount = param.Include[0].TotalCount;

                for (int i = 1; i < IncludeCount; i++)
                {
                    if (formatName.Length>1)
                    for (int j = 0; j < formatName.Length-1; j++)
                    {
                        if (j + 1 == i)
                        {
                            formatedData = new FormatedData();
                            formatedData.Naimenovanie = formatName[i];
                            formatedData.SecondObozna4enie =  param.Include[i].Obozna4enie;
                            formatedData.Count =  param.Include[i].Count.ToString();
                            formatedData.TotalCount =  param.Include[i].TotalCount.ToString();

                            //если изделие входит в состав головного изделия - подсчитываем общее число изделий

                            if (param.VhoditVSostav)
                            {
                                SummCommonCount += param.Include[i].TotalCount;
                            }

                            //добавляем строку с общим количеством для изделия являющихся частью головного
                            if ((i + 1 == IncludeCount) && param.VhoditVSostav && formatName.Length >= IncludeCount)
                            {
                                if (formatName.Length > IncludeCount)
                                {

                                    formatedData.UnderlinedCommonCount = true;
                                    outData.Enqueue(formatedData);
                                    formatedData = new FormatedData();
                                    var endPartName = new List<string>();
                                    for (int k = i + 1; k < formatName.Length; k++)
                                    {
                                        endPartName.Add(formatName[k]);
                                    }
                                    formatedData.Naimenovanie = string.Join("\n",endPartName);
                                    formatedData.TotalCount =   SummCommonCount.ToString();
                                  
                                    
                                }
                                else
                                {
                                    formatedData.UnderlinedCommonCount = true;
                                    outData.Enqueue(formatedData);
                                    formatedData = new FormatedData();
                                    formatedData.TotalCount = SummCommonCount.ToString();
                                    
                                    
                                }
                                outData.Enqueue(formatedData);
                                return;

                            }
                            else
                            {
                                outData.Enqueue(formatedData);
                            }
                        }         
                    }
                }

                


                //заносит все связанные с головным документы в очередь
                for (int i = formatName.Length; i < IncludeCount; i++)
                {
                    formatedData = new FormatedData();


                    {
                        formatedData.SecondObozna4enie = param.Include[i].Obozna4enie;
                        formatedData.Count = param.Include[i].Count.ToString();
                        formatedData.TotalCount = param.Include[i].TotalCount.ToString();
                    }

                    //если изделие входит в состав головного изделия - подсчитываем общее число изделий
                    
                    if (param.VhoditVSostav)
                    {
                        SummCommonCount += param.Include[i].TotalCount;
                    }

                    //добавляем строку с общим количеством для изделия являющихся частью головного
                    if ((i + 1 == IncludeCount) && param.VhoditVSostav)
                    {
                        formatedData.UnderlinedCommonCount = true;
                        outData.Enqueue(formatedData);
                        formatedData = new FormatedData();
                        formatedData.TotalCount = SummCommonCount.ToString();

                        outData.Enqueue(formatedData);
                        break;
                    }
                   
                    outData.Enqueue(formatedData);                    
                }
                
            }


            private string _obozna4enie; //столбец "Обозначение"
            private string _naimenovanie; //столбец "Наименование"
            private string _secondObozna4enie; //столбец "Куда входит: Обозначение"
            private string _count; //столбец "Куда входит: Количество"
            private string _commonCount; //столбец "Куда входит: Общее кол-во"
            private string _prime4anie; //столбец "Примечание"
            //нужно ли подчеркивать значение в поле "Куда входит: Общее кол-во" для изделий входящих в состав
            private bool _underlinedCommonCount;

            public string Obozna4enie
            {
                get { return _obozna4enie; }
                set
                {
                    if (value == null)
                        _obozna4enie = "";
                    else
                        _obozna4enie = value;
                }
            }
            public string Naimenovanie
            {
                get { return _naimenovanie; }
                set
                {
                    if (value == null)
                        _naimenovanie = "";
                    else
                        _naimenovanie = value;
                }
            }
            public string SecondObozna4enie
            {
                get { return _secondObozna4enie; }
                set
                {
                    if (value == null)
                        _secondObozna4enie = "";
                    else
                        _secondObozna4enie = value;
                }
            }
            public string Count
            {
                get { return _count; }
                set
                {
                    if (value == null || value == "0")
                        _count = "";
                    else
                        _count = value;
                }
            }
            public string TotalCount
            {
                get { return _commonCount; }
                set
                {
                    if (value == null || value == "0")
                        _commonCount = "";
                    else
                        _commonCount = value;
                }
            }
            public string Prime4anie
            {
                get { return _prime4anie; }
                set
                {
                    if (value == null)
                        _prime4anie = "";
                    else
                        _prime4anie = value;
                }
            }
            public bool UnderlinedCommonCount
            {
                get { return _underlinedCommonCount; }
                set { _underlinedCommonCount = value; }
            }
            public FormatedData()
            {
                Obozna4enie = "";
                Naimenovanie = "";
                SecondObozna4enie = "";
                Count = "";
                TotalCount = "";
                Prime4anie = "";
                UnderlinedCommonCount = false;
            }
            public FormatedData(string obozna4enie, string naimenovanie, string secondObozna4enie, string count, string commonCount, string prime4anie, bool underlinedCommonCount)
            {
                Obozna4enie = obozna4enie;
                Naimenovanie = naimenovanie;
                SecondObozna4enie = secondObozna4enie;
                Count = count;
                TotalCount = commonCount;
                Prime4anie = prime4anie;
                UnderlinedCommonCount = underlinedCommonCount;
            }

            static private Font GetFont()
            {
                return new Font("T-FLEX Type B", 2.5f, GraphicsUnit.Millimeter);
            }
        }

        //создает отчет в CAD
        static public void MakeCADReport(Queue<FormatedData> data, SpecListParams BaseDocument, IReportGenerationContext context , DocumentAttributes  attribute)
        {
          
            FileReference fileReference = new FileReference();

            if (BaseDocument == null)
                BaseDocument = new SpecListParams(0);

            if (data == null)
                data = new Queue<FormatedData>();
            TFlex.Model.Document document = null;    //Документ CAD            
            document = TFlex.Application.OpenDocument(context.ReportFilePath);            
            document.BeginChanges("Заполнение данными первой страницы");
            //получаем указатель на шаблон текста
            ParagraphText TextTablePg = null;
                       
            foreach (Text text in document.Texts) //TFlex.Application.ActiveDocument.Texts)
            {
                if (text == null)
                    continue;

                if (text.Name == "$spec_list_table")
                    TextTablePg = text as ParagraphText;    
                
                if (TextTablePg != null)
                    break;

            }

            if (TextTablePg == null)
                return;

            TextTablePg.BeginEdit();            

            //подключаем таблицу из шаблона текста
            Table TablePg = TextTablePg.GetTableByIndex(0);
            uint RecordsCount = (uint)data.Count;
            // получение аппаратурной входимости
            string AppVhodimost = BaseDocument.Oboznachenie;
            
            for (uint i = 0; i < RecordsCount; i++)             
                InsertRowInTable(TextTablePg, TablePg, data.Dequeue(), AppVhodimost);

            if(attribute.countReservPages > 0)
            {
                for (uint i = 0; i < 28* (attribute.countReservPages); i++)
                    InsertEmtyRowInTable(TextTablePg, TablePg);
            }
            
            // string PathToTemplates = @"D:\Шаблоны отчетов\"; // @"D:\Тест - отчет\";
            string PathToTemplates = @"Служебные файлы\Шаблоны отчётов\Ведомость спецификаций ГОСТ 2.106-96\";

            var fileObject = fileReference.FindByRelativePath(PathToTemplates + "Ведомость спецификаций Форма 3 ГОСТ 2.106-96.grb");
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

            FileLink link1 = new FileLink(document, fileObject.LocalPath, true);

            fileObject = fileReference.FindByRelativePath(PathToTemplates + "Ведомость спецификаций Форма 3а ГОСТ 2.106-96.grb");
            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

            FileLink link1a = new FileLink(document, fileObject.LocalPath, true);

            //вставляем в документ фрагменты форматок
            Fragment fr; //вставляемый фрагмент
            int List = 3; //значение поля Лист для последующих листов спецификации
            string pattern = @"^(?<g1>[А-Я]{4})\.(?<g2>[0-9]{6})\.(?<g3>[0-9]{3})(?<vard>-(?<var>[0-9]{2,3}))?";
            Regex r = new Regex(pattern, RegexOptions.IgnoreCase);
            int pageNumber = 1; 
            foreach (Page page in document.Pages)
            {
                pageNumber++;
                if (page.Name == "Страница 1")
                {
                    //вставляем фрагмент форматки
                    fr = new Fragment(link1);


                    //заполняем переменные фрагмена первой страницы
                    foreach (FragmentVariableValue var in fr.VariableValues)
                    {
                        switch (var.Name)
                        {
                            case "$oboznach":
                                //убираем регулярку обозначения, в связи с новым ГОСТом 
                               // Match m = r.Match(BaseDocument.Oboznachenie);
                                var.TextValue = BaseDocument.Oboznachenie;//m.ToString();
                                // var.TextValue = BaseDocument.Oboznachenie;
                                break;
                            case "$per_prim":                                                             // вставил - Veleszmey
                                var.TextValue = BaseDocument.Oboznachenie;
                                break;
                            case "$list":
                                if (document.Pages.Count > 1)
                                    var.TextValue = "2";
                                break;
                            case "$listov":
                                uint CountList = document.Pages.Count + 2;
                                var.TextValue = CountList.ToString();
                                break;
                            case "$naimen1":
                                var.TextValue = BaseDocument.Naimenovanie;
                                break;
                            case "$izm2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberModify;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberModify;
                                }
                                break;
                            case "$il2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.ListZam;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.ListNov;
                                }
                                break;
                            case "$ido2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberDocument.ToString();

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberDocument.ToString();
                                }
                                break;
                            case "$idata2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.DateModify;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.DateModify;
                                }
                                break;
                        }
                    }
                }
                else
                {
                    fr = new Fragment(link1a);

                    //заполняем переменные фрагмента последующих страниц
                    foreach (FragmentVariableValue var in fr.VariableValues)
                    {
                        switch (var.Name)
                        {
                            case "$oboznach":
                                //убираем регулярку обозначения, в связи с новым ГОСТом
                               // Match m = r.Match(BaseDocument.Oboznachenie);
                                var.TextValue = BaseDocument.Oboznachenie;// m.ToString();
                                // var.TextValue = BaseDocument.Oboznachenie;
                                break;
                            case "$list":
                                var.TextValue = List.ToString();
                                List++;
                                break;
                            case "$naimen1":
                                var.TextValue = BaseDocument.Naimenovanie;
                                break;
                            case "$izm2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberModify;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberModify;
                                }
                                break;
                            case "$il2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.ListZam;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.ListNov;
                                }
                                break;
                            case "$ido2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberDocument.ToString();

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.NumberDocument.ToString();
                                }
                                break;
                            case "$idata2":
                                if (attribute.AddModify)
                                {
                                    foreach (int nPage in attribute.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.DateModify;

                                    foreach (int nPage in attribute.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = attribute.DateModify;
                                }
                                break;
                        }
                    }
                }

                fr.Page = page;
            }            

            TextTablePg.EndEdit();
            document.EndChanges();
            document.Save();         // сохраняем измененный документ

            document.Close();
            System.Diagnostics.Process.Start(context.ReportFilePath);
        }

        //заносит одну строку данных в таблицу
        static private void InsertRowInTable(ParagraphText parText, Table table, FormatedData pars, string AVhodimost)
        {
            table.InsertRows(1, (uint)(table.CellCount - 1), Table.InsertProperties.After);
            CharFormat ch;
            uint NewCellsCount = (uint)table.CellCount;
            uint CollsCount = (uint)table.ColumnCount;            
            table.InsertText(NewCellsCount - CollsCount, 0, pars.Obozna4enie);
            table.InsertText(NewCellsCount - CollsCount + 1, 0, pars.Naimenovanie);
            table.InsertText(NewCellsCount - CollsCount + 5, 0, pars.Prime4anie);
            // аппаратурную входимость не указывать

            if (pars.SecondObozna4enie == AVhodimost)
            {
                table.InsertText(NewCellsCount - CollsCount + 2, 0, "");
                table.InsertText(NewCellsCount - CollsCount + 3, 0, "");
            }
            else
            {
                table.InsertText(NewCellsCount - CollsCount + 2, 0, pars.SecondObozna4enie);
                table.InsertText(NewCellsCount - CollsCount + 3, 0, pars.Count);
            }

            if (pars.UnderlinedCommonCount)
            {
                //подчеркивает значение в столбце Общее кол-во
                table.SetSelection(NewCellsCount - CollsCount + 4, NewCellsCount - CollsCount + 4);
                ch = parText.CharacterFormat;
                ch.isUnderline = true;
                table.InsertText(NewCellsCount - CollsCount + 4, 0, pars.TotalCount, ch);

                return;
            }
            table.InsertText(NewCellsCount - CollsCount + 4, 0, pars.TotalCount);
        }

        static private void InsertEmtyRowInTable(ParagraphText parText, Table table)
        {
            table.InsertRows(1, (uint)(table.CellCount - 1), Table.InsertProperties.After);

        }
        //раздел сборочные единицы
        public static List<SpecListParams> m_sboro4nyeEdinicy;
        // раздел комплексы
        public static List<SpecListParams> m_komplekty;
        // составные части изделия, имеющие свои ведомости спецификаций
        public static List<SpecListParams> m_vedomosti;


        public static SpecListParams FillData(int baseDocumentID,LogDelegate logDelegate)
        {

            SpecListParams baseDocument = FillData(baseDocumentID,
                                            out m_sboro4nyeEdinicy, out m_komplekty, out m_vedomosti, logDelegate);
            return baseDocument;
        }


        public static SpecListParams FillData(
                            int baseDocumentID,
                            out List<SpecListParams> sboro4nyeEdinicy,
                            out List<SpecListParams> komplekty,
                            out List<SpecListParams> vedomosti, LogDelegate logDelegate)
        {

            // Инициализация выходных переменных

            // раздел Сборочные единицы
            sboro4nyeEdinicy = new List<SpecListParams>();
            // раздел Комплекты
            komplekty = new List<SpecListParams>();
            // части изделия, имеющие свои ведомости спецификации 
            vedomosti = new List<SpecListParams>();

           
            // Получение параметров документов
            ReadReportDocuments(baseDocumentID, null, sboro4nyeEdinicy, komplekty, vedomosti, logDelegate);
            
            return BaseDocument;
        }

        internal class SpecListParams
        {
            public class Entrances
            {
                public int s_ObjectID; //ID головного изделия
                public string Obozna4enie; //обозначение головного изделия
                public int Count; //число вхождений в головное изделие
                public int TotalCount; //общее число вхождений

                public Entrances(int iD)
                {
                    s_ObjectID = iD;
                    Obozna4enie = "";
                    Count = 0;
                    TotalCount = 0;
                }

                public Entrances(string obozna4enie, int count, int commonCount)
                {
                    Obozna4enie = obozna4enie;
                    Count = count;
                    TotalCount = commonCount;
                }
                // ID документа
                public int DocumentID
                {
                    get { return s_ObjectID; }
                }

                public int TotalCountEntr //ID родителя
                {
                    get { return TotalCount; }
                    set { TotalCount = value; }
                }

                public int CountEntr //кол-во изделий
                {
                    get { return Count; }
                    set { Count = value; }
                }

                // Обозначение документа
                public string Oboznachenie
                {
                    get { return Obozna4enie; }
                    set { Obozna4enie = value; }
                }
            }
                        
            private int _s_ObjectID; //ID документа
            private int _parentDocID; //ID родителя
            private int _count; //число изделий
            private int _totalcount; //общее количество изделий
            private string _oboznachenie; //обозначение документа
            private string _naimenovanie; //наименование документа
            private string _prime4anie; //примечание спецификации
            private bool _vhoditVSostav; //составная часть входит в состав изделия
            public List<Entrances> Include; //список вхождений обозначений спецификаций
            private string _firstUse; //первичная применяемость


            public SpecListParams(int DocumentID)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 0;
                _oboznachenie = "";
                _naimenovanie = "";
                _prime4anie = "";
                _vhoditVSostav = false;
                Include = new List<Entrances>();
            }

            public SpecListParams(int DocumentID, string obozna4enie, string naimenovanie, string firstUse)
            {
                _s_ObjectID = DocumentID;
                _parentDocID = 0;
                _count = 1;
                _oboznachenie = obozna4enie;
                _naimenovanie = naimenovanie;
                _prime4anie = "";
                _vhoditVSostav = false;
                Include = new List<Entrances>();
                _firstUse = firstUse;
            }

            public override string ToString()
            {
                return _oboznachenie
                    + "(id=" + _s_ObjectID.ToString()
                    + ", наим=" + _naimenovanie
                    + ", прим=" + _prime4anie
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

            public int Count //кол-во изделий
            {
                get { return _count; }
                set { _count = value; }
            }

            public int TotalCount
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
            // первичная применяемость
            public string FirstUse
            {
                get { return _firstUse; }
                set { _firstUse = value; }
            }
        }

        public static SpecListParams BaseDocument;
    
        //------------------------------------------------------
        public static void ReadReportDocuments(
            int documentId,
            SpecListParams baseDocumentSpecListParams,
            List<SpecListParams> sboro4nyeEdinicy,
            List<SpecListParams> komplekty,
            List<SpecListParams> vedomosti,
            LogDelegate logDelegate)
        {   
            // соединяемся с БД T-FLEX DOCs 2010
            SqlConnection conn;
            //string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=S2"; //SRV1";
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            var info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            var reference = info.CreateReference();
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
                                                                                AND ConstructorParams.s_Deleted = 0",documentId),conn);
            getFirstUseCommand.CommandTimeout = 0;                
            string firstUse = Convert.ToString(getFirstUseCommand.ExecuteScalar());
                 

            SqlCommand getBaseDocumentNameCommand = new SqlCommand(String.Format(@"SELECT [Name]
                        FROM Nomenclature
                        WHERE Nomenclature.s_ObjectID={0}
                            AND Nomenclature.s_ActualVersion = 1
                            AND Nomenclature.s_Deleted = 0", documentId), conn);
            getBaseDocumentNameCommand.CommandTimeout = 0;
            string Naimenovanie = Convert.ToString(getBaseDocumentNameCommand.ExecuteScalar());
                               
            BaseDocument = new SpecListParams(documentId, Oboznachenie, Naimenovanie, firstUse);
                
            #region Чтение полей ВС

            /* Чтение ID классов */
            int classAssembly = TFDClass.Assembly;
            int classComplement = TFDClass.Komplekt;
            int classComplexProgram = TFDClass.ComplexProgram;
            int classComponentProgram = TFDClass.ComponentProgram;

            #region Читаем все поля ВС
            string getDocTreeCommandText = String.Format(
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
                                                        EntranceDenotation NVARCHAR(255),
                                                        TotalCount FLOAT)
                                ELSE DELETE FROM #TmpVspec

                                INSERT INTO #TmpVspec
                                SELECT n.s_ObjectID,0,0,n.s_ClassID,1,n.Denotation,n.Name,'',1
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
                                         nh.Amount,
                                         n.Denotation,
                                         n.Name,
                                         #TmpVspec.Denotation,
                                         #TmpVspec.TotalCount*nh.Amount
                                  FROM NomenclatureHierarchy nh INNER JOIN #TmpVspec ON nh.s_ParentID=#TmpVspec.s_ObjectID 
                                                                          
                                       INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID 
                                    WHERE 
                                           #TmpVspec.[level]=@level          
                                           AND n.s_Deleted = 0
                                           AND n.s_ActualVersion = 1    
                                           AND nh.s_ActualVersion = 1
                                           AND nh.s_Deleted = 0  
                                           AND n.s_ClassID IN ({1},{2},{3},{4})                                        
     
                                  SET @insertcount = @@ROWCOUNT 
                                  SET @level = @level + 1 
   
                                  IF @insertcount = 0 
                                  GOTO end1

                                END
                                end1:

                                SELECT s_ObjectID, s_ParentID, s_ClassID,  Denotation, Name, EntranceDenotation, Amount, SUM(TotalCount) as TotalCount 
                                FROM #TmpVspec 
                                GROUP BY Denotation, Name, EntranceDenotation, s_ObjectID, s_ParentID, s_ClassID, Amount                        
                                ", documentId, classAssembly,classComplement,
                                               classComplexProgram, classComponentProgram);
            SqlCommand docTreeCommand = new SqlCommand(getDocTreeCommandText, conn);
            docTreeCommand.CommandTimeout = 0;
            SqlDataReader reader = docTreeCommand.ExecuteReader();
            #endregion

            int objectId = 0;
            bool typeDoc = false;

            var pars = new SpecListParams(documentId);

            while (reader.Read())
            {
                if (GetSqlInt32(reader, 0) != documentId)
                {
                    if (objectId != GetSqlInt32(reader, 0))
                    {
                        if (GetSqlInt32(reader, 2) == classAssembly)
                        {
                            pars = new SpecListParams(GetSqlInt32(reader, 0));
                            pars.ParentDocID = GetSqlInt32(reader, 1);
                            pars.Oboznachenie = GetSqlString(reader, 3);
                            pars.Naimenovanie = GetSqlString(reader, 4);
                            pars.VhoditVSostav = true;
                            sboro4nyeEdinicy.Add(pars);
                            logDelegate(String.Format("Добавлена СЕ: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));
                            typeDoc = false;
                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5),
                                                                       Convert.ToInt32(GetSqlDouble(reader, 6)),
                                                                       Convert.ToInt32(GetSqlDouble(reader, 7)));
                            sboro4nyeEdinicy[sboro4nyeEdinicy.Count - 1].Include.Add(include);
                        }

                        else
                        {
                            pars = new SpecListParams(GetSqlInt32(reader, 0));
                            pars.ParentDocID = GetSqlInt32(reader, 1);
                            pars.Oboznachenie = GetSqlString(reader, 3);
                            pars.Naimenovanie = GetSqlString(reader, 4);
                            pars.VhoditVSostav = true;
                            komplekty.Add(pars);
                            logDelegate(String.Format("Добавлен комплект: {0} {1}", pars.Oboznachenie, pars.Naimenovanie));
                            typeDoc = true;
                            var include = new SpecListParams.Entrances(GetSqlString(reader, 5),
                                                                       Convert.ToInt32(GetSqlDouble(reader, 6)),

                                                                       Convert.ToInt32(GetSqlDouble(reader, 7)));
                            komplekty[komplekty.Count - 1].Include.Add(include);
                        }
                    }
                    else
                    {
                        var include = new SpecListParams.Entrances(GetSqlString(reader, 5),
                                                                    Convert.ToInt32(GetSqlDouble(reader, 6)),
                                                                    Convert.ToInt32(GetSqlDouble(reader, 7)));
                        if (typeDoc)
                            komplekty[komplekty.Count - 1].Include.Add(include);
                        else sboro4nyeEdinicy[sboro4nyeEdinicy.Count - 1].Include.Add(include);
                    }

                    objectId = GetSqlInt32(reader, 0);
                }
            }
            
            reader.Close();
            #endregion

            conn.Close();            
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
        
    }

    //Форма
    namespace Vspecnew
    {
        partial class MainForm
        {
            /// <summary>
            /// Required designer variable.
            /// </summary>
            private System.ComponentModel.IContainer components = null;

            /// <summary>
            /// Clean up any resources being used.
            /// </summary>
            /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
            protected override void Dispose(bool disposing)
            {
                if (disposing && (components != null))
                {
                    components.Dispose();
                }
                base.Dispose(disposing);
            }

            #region Windows Form Designer generated code

            /// <summary>
            /// Required method for Designer support - do not modify
            /// the contents of this method with the code editor.
            /// </summary>
            private void InitializeComponent()
            {
                this.canselButton = new System.Windows.Forms.Button();
                this.progressBar1 = new System.Windows.Forms.ProgressBar();
                this.LogTextBox = new System.Windows.Forms.TextBox();
                this.label1 = new System.Windows.Forms.Label();
                this.Processinglabel = new System.Windows.Forms.Label();
                this.SuspendLayout();
                // 
                // canselButton
                // 
                this.canselButton.Location = new System.Drawing.Point(382, 291);
                this.canselButton.Name = "canselButton";
                this.canselButton.Size = new System.Drawing.Size(75, 23);
                this.canselButton.TabIndex = 0;
                this.canselButton.Text = "Отмена";
                this.canselButton.UseVisualStyleBackColor = true;
                this.canselButton.Click += new System.EventHandler(this.canselButton_Click);
                // 
                // progressBar1
                // 
                this.progressBar1.Location = new System.Drawing.Point(20, 262);
                this.progressBar1.Name = "progressBar1";
                this.progressBar1.Size = new System.Drawing.Size(437, 23);
                this.progressBar1.TabIndex = 2;
                // 
                // LogTextBox
                // 
                this.LogTextBox.BackColor = System.Drawing.SystemColors.ButtonFace;
                this.LogTextBox.Location = new System.Drawing.Point(21, 46);
                this.LogTextBox.Multiline = true;
                this.LogTextBox.Name = "LogTextBox";
                this.LogTextBox.Size = new System.Drawing.Size(436, 164);
                this.LogTextBox.TabIndex = 3;
                // 
                // label1
                // 
                this.label1.AutoSize = true;
                this.label1.Location = new System.Drawing.Point(18, 18);
                this.label1.Name = "label1";
                this.label1.Size = new System.Drawing.Size(101, 13);
                this.label1.TabIndex = 4;
                this.label1.Text = "Журнал операций:";
                // 
                // Processinglabel
                // 
                this.Processinglabel.AutoSize = true;
                this.Processinglabel.Location = new System.Drawing.Point(23, 235);
                this.Processinglabel.Name = "Processinglabel";
                this.Processinglabel.Size = new System.Drawing.Size(0, 13);
                this.Processinglabel.TabIndex = 5;
                // 
                // MainForm
                // 

                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(485, 332);
                this.Controls.Add(this.Processinglabel);
                this.Controls.Add(this.label1);
                this.Controls.Add(this.LogTextBox);
                this.Controls.Add(this.progressBar1);
                this.Controls.Add(this.canselButton);
                this.Name = "MainForm";
                this.Text = "Формирование ведомости спецификаций (ВС)";
                this.ResumeLayout(false);
                this.PerformLayout();

            }

            #endregion

            private Button canselButton;
            private ProgressBar progressBar1;
            private TextBox LogTextBox;
            private Label label1;
            private Label Processinglabel;
        }
    }


    namespace Vspecnew
    {
        public partial class MainForm : Form
        {

            public enum Stages
            {
                Initialization,
                DataAcquisition,
                DataProcessing,
                ReportGenerating,
                Done
            }

            public MainForm()
            {
                Application.EnableVisualStyles();
                InitializeComponent();
            }

            public void writeToLog(string line)
            {
                DateTime now = DateTime.Now;
                string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
                LogTextBox.AppendText(s);
                LogTextBox.Invalidate();
                this.Update();
                Application.DoEvents();
            }

            public void progressBarParam(int step)
            {

                progressBar1.PerformStep();
            }
            public void setStage(Stages stage)
            {
                switch (stage)
                {
                    case Stages.Initialization:
                        Processinglabel.Text = "Инициализация...";
                        break;
                    case Stages.DataAcquisition:
                        Processinglabel.Text = "Получение данных";
                        break;
                    case Stages.DataProcessing:
                        Processinglabel.Text = "Обработка данных";
                        break;
                    case Stages.ReportGenerating:
                        Processinglabel.Text = "Создание отчёта";
                        break;
                    case Stages.Done:
                        break;
                    default:
                        System.Diagnostics.Debug.Assert(false);
                        break;
                }
                this.Update();
            }

            private void canselButton_Click(object sender, EventArgs e)
            {
                MainForm.ActiveForm.Close();
            }
        }
    }    
}
