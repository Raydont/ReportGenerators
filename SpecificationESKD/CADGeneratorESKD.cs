using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using SpecificationESKD.Globus.TFlexDocs.SpecificationReport.GOST_2_106;
using TFlex;
using TFlex.DOCs.Model;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Reporting;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Structure;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using Rectangle = TFlex.Drawing.Rectangle;

//using tfdAPI;
//using Globus.TFlexDocs;

namespace SpecificationESKD
{

    public class CADGenerator : IReportGenerator
    {
        //static CADGenerator()
        //{
        //   // Добавляем ссылку на директории с API CAD 11
        //    AssemblyResolver.Instance.AddDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + @"\T-FLEX\T-FLEX CAD 11\Program\");
        //}

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
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        public Attributes attr = new Attributes();

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

            try
            {
                attr.DotsEntry(context);

                //string str = @" taskkill /IM TFlex.DOCs.FilePreview.Provider.exe /F";
                //ExecuteCommand(str);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("" + e, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (document != null)
                    document.Close();
            }
        }

        public void ExecuteCommand(string str)
        {
            ProcessStartInfo startInfo;
            Process process = null;
            string command;

            command = str;

            try
            {
                //Проверка 
                if (command == null || command.Trim().Length == 0)
                    throw new ArgumentNullException("command");

                startInfo = new ProcessStartInfo();

                //Запускаем через cmd
                startInfo.FileName = "cmd.exe";
                //Ключ /c - выполнение команды
                startInfo.Arguments = "/C " + command;

                //Не используем shellexecute
                startInfo.UseShellExecute = false;
                //Перенаправить вывод на обычную консоль
                startInfo.RedirectStandardOutput = false;
                //Стартуем
                process = Process.Start(startInfo);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (process != null)
                    //Закрываем
                    process.Close();
                //Освобождаем
                process = null;
                startInfo = null;
                // GC.Collect();
                GC.GetTotalMemory(true);
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
                    AppDomain.CurrentDomain.AssemblyResolve += new System.ResolveEventHandler(AssemblyResolver.ResolveEventHandler);
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
                AssemblyName name = new AssemblyName();
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
                    StringBuilder builder = new StringBuilder(directory);
                    builder.Append(@"\");
                    builder.Append(resolveAsmName);
                    builder.Append(".dll");
                    string path = builder.ToString();
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        AssemblyName fullName = new AssemblyName();
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

    public class TFDDocument
    {
        // Порядок полей в запросе чтения документа

        public int ObjectID;
        public int ParentID;
        public int Class;
        public string FirstUse;
        public double Amount;
        public string Zone;
        public int Position;
        public string Format;
        public string Naimenovanie;
        public string Denotation;
        public string Remarks;
        public string BomSection = string.Empty;
        public string MeasureUnit;
        public TFDDocument rootDocument;
        public ReferenceObject ReplacingObject;
        //private List<TFDDocument> _childDocuments;

        public TFDDocument()
        {
        }

        public TFDDocument(int documentID)
        {

            NomenclatureReference refNomenclature = new NomenclatureReference();


            //// TFlex.DOCs.Model.Structure.ParameterInfo classInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
            //TFlex.DOCs.Model.Structure.ParameterInfo nameInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
            //TFlex.DOCs.Model.Structure.ParameterInfo denotationInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation); ;
            ////TFlex.DOCs.Model.Structure.ParameterInfo bomSectionInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.);
            rootDocument = null;
            ParentID = 0;
            ObjectID = documentID;

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomenclatureReferenceInfo.CreateReference();
            var nomObj = reference.Find(documentID);


            Naimenovanie = nomObj[new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")].Value.ToString();
            Denotation = nomObj[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].Value.ToString();

            if (Denotation.Contains("ПИ"))
            {

                // Class = nomenclatureObject[classInfo].Value.ToString();

                Amount = 1;

                //Naimenovanie = nomenclatureObject[nameInfo].Value.ToString();

                //Denotation = nomenclatureObject[denotationInfo].Value.ToString();
                // Вырезаем из обозначения последние два символа (ПИ) по требованию Баскаковой

                if (Denotation.Contains("ПИ"))
                {
                    string pi = Denotation.Trim().Remove(Denotation.Length - 2, 2);
                    if (Denotation.Contains("ПИ") && !pi.Contains("ПИ"))
                    {
                        Denotation = Denotation.Trim().Remove(Denotation.Length - 2, 2);
                    }

                    if (Denotation.Contains("ПИ") && !(Denotation.Trim().Remove(Denotation.Length - 5, 2)).Contains("ПИ"))
                    {
                        Denotation = Denotation.Trim().Remove(Denotation.Length - 5, 2);

                    }
                    if (Denotation.Contains("ПИ") && !(Denotation.Trim().Remove(Denotation.Length - 6, 2)).Contains("ПИ"))
                    {
                        Denotation = Denotation.Trim().Remove(Denotation.Length - 6, 2);

                    }

                    Remarks = String.Empty;
                }
            }

        }
    }

    internal static class ApiDocs
    {
        public static SqlConnection GetConnection(bool open)
        {
            //  string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=S2"; //SRV1";
            // string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=SRV1";

            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            //sqlConStringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;

            ReferenceInfo info = ServerGateway.Connection.ReferenceCatalog.Find(SpecialReference.GlobalParameters);
            Reference reference = info.CreateReference();
            var nameSqlServer = reference.Find(new Guid("36d4f5df-360e-4a44-b3b7-450717e24c0c"))[new Guid("dfda4a37-d12a-4d18-b2c6-359207f407d0")].ToString();
            sqlConStringBuilder.DataSource = nameSqlServer;
            sqlConStringBuilder.InitialCatalog = "TFlexDOCs";
            sqlConStringBuilder.Password = "reportUser";
            sqlConStringBuilder.UserID = "reportUser";

            try
            {
                //подключение к БД T-FLEX DOCs
                SqlConnection connection = new SqlConnection(sqlConStringBuilder.ToString());
                SqlConnection.ClearPool(connection);
                if (open)
                    connection.Open();

                return connection;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }

            return null;
        }
    }

    // ===================================== ТЕКСТ МАКРОСА =====================================

    namespace Globus.TFlexDocs
    {
        using System;
        using System.Collections.Generic;
        using System.Text;

        /** Работа с объектом класса Атрибуты документов.
         *  Класс поддерживает:
         *  - чтение специфичных для объекта этого класса параметров;
         *  - копирование объектов из состава объекта класса Атрибуты документов
         *    в состав заданного объекта
         */
        public class DocumentAttributes
        {
            TFlexDocs _tflexdocs;
            int _id;

            // Значения параметров
            string _designedBy = "";
            string _controlledBy = "";
            string _nControlBy = "";
            string _approvedBy = "";
            int _firstPageNumber = 2;
            bool _addChangelogPage = false;
            string _sidebarSignHeader1 = "";
            string _sidebarSignName1 = "";
            string _sidebarSignHeader2 = "";
            string _sidebarSignName2 = "";
            string _sidebarSignHeader3 = "";
            string _sidebarSignName3 = "";
            int _countEmptyLinesBefore;
            int _countEmptyLinesAfter;
            bool _fullNames = false;
            bool _zeroToDash = true;
            bool _addMake = false;
            bool _addModify = false;
            bool _sortByAlphabet = false;

            private bool _addPagesForExecution = false;

            public bool ThroughNumbering = true;


            private decimal _documentationPages = 0;
            private decimal _assemblyPages = 0;
            private decimal _detailsPages = 0;
            private decimal _stdProductsPages = 0;
            private decimal _otherProductsPages = 0;
            private decimal _materialsPages = 0;
            private decimal _komplektPages = 0;

            public bool specificationForPI = false;
            public bool doc = false;
            public bool assem = false;
            public bool det = false;
            public bool std = false;
            public bool oth = false;
            public bool mat = false;
            public bool kom = false;
            public bool newPageOIAfterKindProducts = false;
            public bool newPageOIAfterKindProductsVariosPart = false;
            public bool checkChiefMillitary = false;
            public bool checkVPMO = false;
            public List<ReferenceObject> ReplacingObjects = new List<ReferenceObject>();

            private string goev = string.Empty;
            private string militaryChief = string.Empty;
            private string tech = string.Empty;
            private string otdel45 = string.Empty;
            private string norm = string.Empty;
            private string vPMO = string.Empty;
            private string chiefTeam = string.Empty;
            private string yearMake = string.Empty;
            private string fullName = string.Empty;

            // для вкладки Добавление изменений
            private string numberModify = string.Empty;                 // номер изменения
            private string numberDocument = string.Empty; // номер документа
            private string dateModify = string.Empty;           // дата изменения
            private List<int> listPage = new List<int>(); // список страниц
            private bool addTitulList = false;
            private string listZam = string.Empty;
            private string listNov = string.Empty;
            private List<int> listPage2 = new List<int>();

            static readonly DocumentAttributes _empty = new DocumentAttributes();
            public bool docVarPart;
            public bool assemVarPart;
            public bool detVarPart;
            public bool stdVarPart;
            public bool othVarPart;
            public bool matVarPart;
            public bool komVarPart;
            public bool addSectionProgramSoft = true;
            public bool programToLower;

            private decimal _documentationPagesVarPart = 0;
            private decimal _assemblyPagesVarPart = 0;
            private decimal _detailsPagesVarPart = 0;
            private decimal _stdProductsPagesVarPart = 0;
            private decimal _otherProductsPagesVarPart = 0;
            private decimal _materialsPagesVarPart = 0;
            private decimal _komplektPagesVarPart = 0;

            public DocumentAttributes()
            {
            }

            //* Конструктор, выполняющий чтение значений параметров из объекта
            public DocumentAttributes(/*TFlexDocs tflexdocs,*/ TFDDocument document)
            {
                //_tflexdocs = tflexdocs;
                _id = document.ObjectID;
                ReadAttributes(/*tflexdocs,*/ document);
            }

            //* Чтение значений параметров из объекта
            private void ReadAttributes(/*TFlexDocs tflexdocs,*/ TFDDocument document)
            {
                //tflexdocs.LoadDocumentParameters(document, LOAD_PARAMETERS.ALL);

                // string designedBy = tflexdocs.GetParameterValue(document, ParametersGUIDs.DesignedBy);
                // _designedBy = (designedBy == null) ? "" : designedBy;

                //  string controlledBy = tflexdocs.GetParameterValue(document, ParametersGUIDs.ControlledBy);
                //  _controlledBy = (controlledBy == null) ? "" : controlledBy;

                //  string nControlBy = tflexdocs.GetParameterValue(document, ParametersGUIDs.NControlBy);
                //  _nControlBy = (controlledBy == null) ? "" : nControlBy;

                //string approvedBy = tflexdocs.GetParameterValue(document, ParametersGUIDs.ApprovedBy);
                //  _approvedBy = (approvedBy == null) ? "" : approvedBy;

                _firstPageNumber = 1;// tflexdocs.GetIntegerParameterValue(document, ParametersGUIDs.FirstPageNumber);

                _addChangelogPage = true;// tflexdocs.GetBooleanParameterValue(document, ParametersGUIDs.AddChangelogPage);

                // string sidebarSignHeader1 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader1);
                // _sidebarSignHeader1 = (sidebarSignHeader1 == null) ? "" : sidebarSignHeader1;

                //  string sidebarSignName1 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName1);
                //  _sidebarSignName1 = (sidebarSignName1 == null) ? "" : sidebarSignName1;

                // string sidebarSignHeader2 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader2);
                //  _sidebarSignHeader2 = (sidebarSignHeader2 == null) ? "" : sidebarSignHeader2;

                //  string sidebarSignName2 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName2);
                // _sidebarSignName2 = (sidebarSignName2 == null) ? "" : sidebarSignName2;

                //  string sidebarSignHeader3 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader3);
                // _sidebarSignHeader3 = (sidebarSignHeader3 == null) ? "" : sidebarSignHeader3;

                //string sidebarSignName3 = tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName3);
                // _sidebarSignName3 = (sidebarSignName3 == null) ? "" : sidebarSignName3;

                _fullNames = false;// tflexdocs.GetBooleanParameterValue(document, ParametersGUIDs.FullNames);
            }

            /** Копирование объектов из состава объекта класса Атрибуты документов
             *  в состав указанного объекта
             */
            /*  public void CopyContent(TFDDocument destination)
              {
                  if (_tflexdocs == null || _id == 0)
                      return;

                  TFDDocument attrDoc = _tflexdocs.GetDocument(_id);
                  int result = attrDoc.LoadComposition(LOAD_TYPE.DOCUMENT_CONTENT, LOAD_MODE.ITEMS,
                                                       LOAD_PARAMETERS.ALL);
                  if (result != 0)
                  {
                      string error = String.Format("Ошибка при загрузке состава объекта {0}",
                                                   _tflexdocs.GetDocumentInfoString(attrDoc));
                      throw new Exception(error);
                  }

                  ITFDEnumerator childDocEnum = attrDoc.Documents.GetEnumerator();
                  childDocEnum.Reset();
                  result = childDocEnum.MoveNext();
                  while (result > 0)
                  {
                      TFDDocument childDocument = (TFDDocument)(childDocEnum.Current);
                      if (childDocument == null)
                      {
                          string error = String.Format("Ошибка при загрузке вложенного объекта из состава {0}",
                                                       _tflexdocs.GetDocumentInfoString(attrDoc));
                          throw new Exception(error);
                      }
                      _tflexdocs.CopyOnDesktop(childDocument, destination);
                      result = childDocEnum.MoveNext();
                  }
              }
              */


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
            public static DocumentAttributes Empty
            {
                get { return _empty; }
            }

            public string DesignedBy
            {
                set { _designedBy = value; }
                get { return _designedBy; }
            }

            public string ControlledBy
            {
                set { _controlledBy = value; }
                get { return _controlledBy; }
            }

            public string NControlBy
            {
                set { _nControlBy = value; }
                get { return _nControlBy; }
            }

            public string ApprovedBy
            {
                set { _approvedBy = value; }
                get { return _approvedBy; }
            }

            public int FirstPageNumber
            {
                set { _firstPageNumber = value; }
                get { return _firstPageNumber; }
            }

            public bool AddChangelogPage
            {
                set { _addChangelogPage = value; }
                get { return _addChangelogPage; }
            }

            public string SidebarSignHeader1
            {
                set { _sidebarSignHeader1 = value; }
                get { return _sidebarSignHeader1; }
            }

            public string SidebarSignName1
            {
                set { _sidebarSignName1 = value; }
                get { return _sidebarSignName1; }
            }

            public string SidebarSignHeader2
            {
                set { _sidebarSignHeader2 = value; }
                get { return _sidebarSignHeader2; }
            }

            public string SidebarSignName2
            {
                set { _sidebarSignName2 = value; }
                get { return _sidebarSignName2; }
            }

            public string SidebarSignHeader3
            {
                set { _sidebarSignHeader3 = value; }
                get { return _sidebarSignHeader3; }
            }

            public string SidebarSignName3
            {
                set { _sidebarSignName3 = value; }
                get { return _sidebarSignName3; }
            }

            public bool FullNames
            {
                set { _fullNames = value; }
                get { return _fullNames; }
            }

            public bool ZeroToDash
            {
                set { _zeroToDash = value; }
                get { return _zeroToDash; }
            }

            public int CountEmptyLinesBefore
            {
                set { _countEmptyLinesBefore = value; }
                get { return _countEmptyLinesBefore; }
            }

            public int CountEmptyLinesAfter
            {
                set { _countEmptyLinesAfter = value; }
                get { return _countEmptyLinesAfter; }
            }

            public decimal DocumentationPages
            {
                get
                {
                    if (_documentationPages > 50)
                    {
                        return 50;
                    }
                    else return _documentationPages;
                }
                set { _documentationPages = value; }
            }

            public decimal AssemblyPages
            {
                get
                {
                    if (_assemblyPages > 50)
                    {
                        return 50;
                    }
                    else return _assemblyPages;
                }
                set { _assemblyPages = value; }
            }

            public decimal DetailsPages
            {
                get
                {
                    if (_detailsPages > 50)
                    {
                        return 50;
                    }
                    else return _detailsPages;
                }
                set { _detailsPages = value; }
            }

            public decimal StdProductsPages
            {
                get
                {
                    if (_stdProductsPages > 50)
                    {
                        return 50;
                    }
                    else return _stdProductsPages;
                }
                set { _stdProductsPages = value; }
            }

            public decimal OtherProductsPages
            {
                get
                {
                    if (_otherProductsPages > 50)
                    {
                        return 50;
                    }
                    else return _otherProductsPages;
                }
                set { _otherProductsPages = value; }
            }

            public decimal MaterialsPages
            {
                get
                {
                    if (_materialsPages > 50)
                    {
                        return 50;
                    }
                    else return _materialsPages;
                }
                set { _materialsPages = value; }
            }

            public decimal KomplektPages
            {
                get
                {
                    if (_komplektPages > 50)
                    {
                        return 50;
                    }
                    else return _komplektPages;
                }
                set { _komplektPages = value; }
            }

            public string Goev
            {
                get { return goev; }
                set { goev = value; }
            }

            public string MilitaryChief
            {
                get { return militaryChief; }
                set { militaryChief = value; }
            }

            public string Tech
            {
                get { return tech; }
                set { tech = value; }
            }

            public string Norm
            {
                get { return norm; }
                set { norm = value; }
            }

            public string Otdel45
            {
                get { return otdel45; }
                set { otdel45 = value; }
            }

            public string VPMO
            {
                get { return vPMO; }
                set { vPMO = value; }
            }

            public string ChiefTeam
            {
                get { return chiefTeam; }
                set { chiefTeam = value; }
            }

            public string FullName
            {
                get { return fullName; }
                set { fullName = value; }
            }

            public string YearMake
            {
                get { return yearMake; }
                set { yearMake = value; }
            }

            public bool AddTitulList
            {
                get { return addTitulList; }
                set { addTitulList = value; }
            }

            public bool AddMake
            {
                get { return _addMake; }
                set { _addMake = value; }
            }

            public bool SortByAlphabet
            {
                get { return _sortByAlphabet; }
                set { _sortByAlphabet = value; }
            }



            public decimal DocumentationPagesVarPart
            {
                get
                {
                    if (_documentationPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _documentationPagesVarPart;
                }
                set { _documentationPagesVarPart = value; }
            }

            public decimal AssemblyPagesVarPart
            {
                get
                {
                    if (_assemblyPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _assemblyPagesVarPart;
                }
                set { _assemblyPagesVarPart = value; }
            }

            public decimal DetailsPagesVarPart
            {
                get
                {
                    if (_detailsPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _detailsPagesVarPart;
                }
                set { _detailsPagesVarPart = value; }
            }

            public decimal StdProductsPagesVarPart
            {
                get
                {
                    if (_stdProductsPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _stdProductsPagesVarPart;
                }
                set { _stdProductsPagesVarPart = value; }
            }

            public decimal OtherProductsPagesVarPart
            {
                get
                {
                    if (_otherProductsPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _otherProductsPagesVarPart;
                }
                set { _otherProductsPagesVarPart = value; }
            }

            public decimal MaterialsPagesVarPart
            {
                get
                {
                    if (_materialsPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _materialsPagesVarPart;
                }
                set { _materialsPagesVarPart = value; }
            }

            public decimal KomplektPagesVarPart
            {
                get
                {
                    if (_komplektPagesVarPart > 50)
                    {
                        return 50;
                    }
                    else return _komplektPagesVarPart;
                }
                set { _komplektPagesVarPart = value; }
            }
        }

    }

    namespace Globus.TFlexDocs
    {
        using System;
        using System.Collections;
        using System.Collections.Generic;
        using System.Diagnostics;
        using System.Data;
        using System.Data.SqlClient;
        using System.IO;
        using System.Security.Cryptography;
        using System.Text;
        using System.Threading;


        // Параметры документа в виде словаря "идентификатор-значение".
        //  Поддерживается создание текстового отчета.        
        public class DocumentParameters : IDictionary<int, string>, IEquatable<DocumentParameters>
        {
            int _id;
            int _parentId;
            string _name;
            string _designation;

            // Словарь с параметрами
            SortedDictionary<int, string> _values;

            // Конструктор, выполняющий чтение указанного набора параметров
            /*public DocumentParameters(TFlexDocs tflexdocs, TFDDocument document, int[] parameters)
            {
                if (tflexdocs == null)
                    throw new ArgumentException("Нулевой указатель на экземпляр TFlexDocs");
                Load(tflexdocs, document, parameters);
            }*/

            // Чтение набора параметров
            /*void Load(TFlexDocs tflexdocs, TFDDocument document, int[] parameters)
            {
                if (document == null)
                    throw new ArgumentException("Нулевой указатель на документ T-FLEX DOCs");
                if (parameters == null)
                    throw new ArgumentException("Нулевой массив идентификаторов параметров");

                _id = document.id;
                _parentId = document.Parent;
                string des = tflexdocs.GetStringParameterValue(document, ParametersGUIDs.Designation);
                if (des == null)
                    des = "";

                string name = tflexdocs.GetStringParameterValue(document, ParametersGUIDs.Name);
                if (name == null)
                    name = "";

                _designation = des;
                _name = name;

                _values = new SortedDictionary<int, string>();
                Dictionary<int, string> all = tflexdocs.GetDocumentParameters(document);
                foreach (int id in parameters)
                {
                    if (all.ContainsKey(id))
                        _values.Add(id, all[id]);
                }
            }*/

            // Преобразование в строку вида Наименование Обозначение [идентификатор]
            public override string ToString()
            {
                StringBuilder s = new StringBuilder();
                if (_name.Length > 0)
                    s.Append(_name);
                if (_designation.Length > 0)
                {
                    if (s.Length > 0)
                        s.Append(" ");
                    s.Append(_designation);
                }
                if (s.Length > 0)
                    s.Append(" ");
                s.AppendFormat("[{0}]", _id);
                return s.ToString();
            }

            // Сравнение документов по обозначению, а при его отсутствии - по наименованию
            public bool Equals(DocumentParameters other)
            {
                if (_designation.Length > 0 || other.Designation.Length > 0)
                    return _designation == other._designation;
                else
                    return _name == other._name;
            }

            public int ID { get { return _id; } }

            public int ParentID { get { return _parentId; } }

            public string Designation { get { return _designation; } }

            public string Name { get { return _name; } }

            // Получение значения параметра по идентификатору
            public string GetValue(int id)
            {
                return _values[id];
            }

            // Получение текстового отчета о параметрах
            /* public string[] GetReport(TFlexDocs tflexdocs)
             {
                 return GetReport(tflexdocs, false, false, false);
             }*/

            // Получение текстового отчета о параметрах
            // zeroAsNull - 0 рассматриваются как отсутствие значения
            // noAsNull - "Нет" рассматривается как отсутствие значения
            // excludeEmpty - не включать в отчет параметры без значений
            //
            // Отчет представляет собой строки вида "Наименование параметра Идентификатор параметра: Значение"
            // и используется для логов и отчетов иерархии. В общем, везде, где
            // требуется иметь дело с текстовым "дампом" определенного набора параметров
            // объекта
            /*public string[] GetReport(TFlexDocs tflexdocs, bool zeroAsNull, bool noAsNull, bool excludeEmpty)
            {
                List<string> text = new List<string>();
                foreach (KeyValuePair<int, string> pair in _values)
                {
                    int id = pair.Key;
                    string value = pair.Value;
                    if (zeroAsNull && value == "0")
                        value = "";
                    if (noAsNull && value == "Нет")
                        value = "";
                    if (excludeEmpty && String.IsNullOrEmpty(value))
                        continue;
                    TFDParameterHeader header = tflexdocs.GetParameterHeader(id);
                    string headerString = String.Format("{0} {1}:", header.Description, id);
                    string parameterLine = String.Format("{0,-45} {1}", headerString, value).TrimEnd();
                    text.Add(parameterLine);
                }
                return text.ToArray();
            }*/

            // IEnumerable
            // ===========

            IEnumerator IEnumerable.GetEnumerator()
            {
                return _values.GetEnumerator();
            }

            // IEnumerable<KeyValuePair<TKey,TValue>>
            // ======================================

            IEnumerator<KeyValuePair<int, string>> IEnumerable<KeyValuePair<int, string>>.GetEnumerator()
            {
                return _values.GetEnumerator();
            }

            // ICollection<KeyValuePair<TKey,TValue>>
            // ======================================

            void ICollection<KeyValuePair<int, string>>.Add(KeyValuePair<int, string> pair)
            {
                throw new Exception("Not allowed");
            }

            void ICollection<KeyValuePair<int, string>>.Clear()
            {
                throw new Exception("Not allowed");
            }

            bool ICollection<KeyValuePair<int, string>>.Contains(KeyValuePair<int, string> pair)
            {
                throw new Exception("Not allowed");
            }

            void ICollection<KeyValuePair<int, string>>.CopyTo(KeyValuePair<int, string>[] arr, int i)
            {
                throw new Exception("Not allowed");
            }

            bool ICollection<KeyValuePair<int, string>>.Remove(KeyValuePair<int, string> pair)
            {
                throw new Exception("Not allowed");
            }

            int ICollection<KeyValuePair<int, string>>.Count
            {
                get { return _values.Count; }
            }

            bool ICollection<KeyValuePair<int, string>>.IsReadOnly
            {
                get { return true; }
            }

            // IDictionary<TKey,TValue>
            // ========================

            string IDictionary<int, string>.this[int key]
            {
                get { return _values[key]; }
                set
                {
                    throw new Exception("Not allowed");
                }
            }

            public ICollection<int> Keys
            {
                get { return _values.Keys; }
            }

            public ICollection<string> Values
            {
                get { return _values.Values; }
            }

            public void Add(int key, string value)
            {
                throw new Exception("Not allowed");
            }

            public bool ContainsKey(int key)
            {
                return _values.ContainsKey(key);
            }

            public bool Remove(int key)
            {
                throw new Exception("Not allowed");
            }

            public bool TryGetValue(int key, out string value)
            {
                return TryGetValue(key, out value);
            }

            // IEquatable<DocumentParameters>
            // ==============================

            bool IEquatable<DocumentParameters>.Equals(DocumentParameters other)
            {
                if (_designation.Length > 0 || other.Designation.Length > 0)
                    return _designation == other._designation;
                else
                    return _name == other._name;
            }

        }

        internal static class TFDClass
        {
            public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
            public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
            public static readonly int OtherDocument = new NomenclatureReference().Classes.AllClasses.Find("Документ").Id;
            public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find(new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114")).Id;
            public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
            public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
            public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
            public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
            public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find("Материал").Id;
            public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
            public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;
        }

        internal static class TFDBomSection
        {
            public static readonly string Documentation = "Документация";
            public static readonly string Assembly = "Сборочные единицы";
            public static readonly string Details = "Детали";
            public static readonly string StdProducts = "Стандартные изделия";
            public static readonly string OtherProducts = "Прочие изделия";
            public static readonly string Materials = "Материалы";
            public static readonly string Komplekts = "Комплекты";
            public static readonly string Complex = "Комплексы";
            public static readonly string Components = "Компоненты";
            public static readonly string ProgramItems = "Программные изделия";

        }


        public class TFlexDocs : IDisposable
        {
            // TextNormalizer _normalizer;

            // BomGroupsReference _bomGroupsReference;
            // ClassSetReference _classSetReference;
            //OrderReference _orderReference;


            public TFlexDocs()
            {
            }

            ~TFlexDocs()
            {
                Dispose();
            }

            public void Dispose()
            {
            }

            public void GetBomEmptyLineCounts(int docID, out int countBefore, out int countAfter)
            {
                /* string sql = "SELECT CountBefore, CountAfter FROM DocPositionInBom WHERE DocID = @DocID";
                 SqlCommand command = new SqlCommand(sql, Connection);
                 command.Parameters.AddWithValue("DocID", docID);
                 SqlDataReader reader = command.ExecuteReader();
                 if (reader.Read())
                 {
                     countBefore = reader.GetInt32(0);
                     countAfter = reader.GetInt32(1);
                 }
                 else
                 {
                     countBefore = 0;
                     countAfter = 0;
                 }
                 reader.Close();*/
                countBefore = 0;
                countAfter = 0;
            }
        }

        /**
         * Позиционное обозначение на схеме
         */
        public class Refdes : IEquatable<Refdes>, IComparable, IComparable<Refdes>
        {
            string _refdes;
            object[] _elements;
            string[] _separators;

            static readonly Regex _elementRegex = new Regex(
                "(?<alpha>[A-Za-zА-Яа-я]+)"             // последовательность букв
                + "|(?<number>[0-9]+)"                  // число
                , RegexOptions.Compiled
            );

            public Refdes(string refdes)
            {
                _refdes = refdes;
                _elements = Slice(refdes, out _separators);
            }

            /**
             * Получение массива элементов строки.
             * 
             * Каждый элемент представляет собой строку или число с дробной часью.
             * Схема разделения строки:
             *      sep0 element0 sep1 element1 ... sepN elementN sepN+1
             * 
             * @param string str
             *      Входная строка
             * @param out string[] separators
             *      Разделители элементов. Их количество всегда на один больше количества элементов.
             * @return
             *      Массив элементов строки.
             */
            object[] Slice(string str, out string[] separators)
            {
                if (str == null)
                {
                    separators = new string[] { };
                    return new object[] { };
                }

                List<string> separatorsList = new List<string>();

                ArrayList elements = new ArrayList();
                MatchCollection mc = _elementRegex.Matches(str);

                int pos = 0;
                foreach (Match m in mc)
                {
                    Group group = null;

                    Group alphaGroup = m.Groups["alpha"];
                    if (alphaGroup.Success)
                    {
                        elements.Add(alphaGroup.Value);
                        group = alphaGroup;
                    }

                    Group numberGroup = m.Groups["number"];
                    if (numberGroup.Success)
                    {
                        var nom = m.Value.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        double element = Convert.ToDouble(nom);
                        elements.Add(element);
                        group = numberGroup;
                    }

                    Debug.Assert(group != null);

                    string separator = str.Substring(pos, group.Index - pos);
                    separatorsList.Add(separator);
                    pos = group.Index + group.Length;
                }

                separatorsList.Add(str.Substring(pos, str.Length - pos));
                separators = separatorsList.ToArray();

                Debug.Assert(elements.Count == separatorsList.Count - 1);

                return elements.ToArray();
            }

            public override bool Equals(object other)
            {
                if (other is Refdes == false)
                    throw new ApplicationException("Сравнение объекта класса Refdes с объектом другого класса");
                return Equals(other as Refdes);
            }

            public override int GetHashCode()
            {
                return _refdes.GetHashCode();
            }

            public bool Equals(Refdes other)
            {
                if (other == null)
                    return false;
                else
                {
                    if (_refdes == other._refdes)
                        return true;
                    else
                        return false;
                }
            }

            public int CompareTo(object other)
            {
                if (other is Refdes == false)
                    throw new ApplicationException("Сравнение объекта класса Refdes с объектом другого класса");
                return CompareTo(other as Refdes);
            }

            public int CompareTo(Refdes other)
            {
                IEnumerator e1 = _elements.GetEnumerator();
                bool next1 = e1.MoveNext();
                IEnumerator e2 = other._elements.GetEnumerator();
                bool next2 = e2.MoveNext();
                while (true)
                {
                    if (!next1 && next2)
                        // this имеет меньше элементов, чем other
                        return -1;
                    else if (next1 && !next2)
                        // this имеет больше элементов, чем other
                        return 1;
                    else if (!next1 && !next2)
                    {
                        // this и other имеют равное количество элементов
                        return _refdes.CompareTo(other._refdes);
                    }
                    else
                    {
                        // сравнение очередной пары элементов
                        object element1 = e1.Current;
                        object element2 = e2.Current;
                        if (element1 is string && element2 is string)
                        {
                            // Сравнение двух строк
                            int c = String.Compare((string)element1, (string)element2);
                            if (c != 0)
                                return c;
                        }
                        else if (element1 is double && element2 is string)
                        {
                            return -1;
                        }
                        else if (element1 is string && element2 is double)
                        {
                            return 1;
                        }
                        else if (element1 is double && element2 is double)
                        {
                            // Сравнение двух чисел
                            double num1 = (double)element1;
                            double num2 = (double)element2;
                            if (num1 < num2)
                                return -1;
                            else if (num1 > num2)
                                return 1;
                        }
                        else
                            Debug.Assert(false);
                    }
                    next1 = e1.MoveNext();
                    next2 = e2.MoveNext();
                }
            }

            public Refdes IncreaseLastNumber()
            {
                if (_elements.Length == 0)
                    return null;

                if (_elements[_elements.Length - 1] is double == false)
                    return null;

                StringBuilder nextRefdes = new StringBuilder();
                int i = 0;
                while (i < _elements.Length - 1)
                {
                    nextRefdes.Append(_separators[i]);
                    nextRefdes.Append(_elements[i]);
                    ++i;
                }

                nextRefdes.Append(_separators[i]);
                nextRefdes.Append((double)_elements[i] + 1);
                nextRefdes.Append(_separators[i + i]);

                return new Refdes(nextRefdes.ToString());
            }

            public override string ToString()
            {
                return _refdes;
            }
        }
    }

    namespace Globus.TFlexDocs.SpecificationReport.GOST_2_106
    {
        using System;
        using System.Collections;
        using System.Collections.Generic;
        using System.Diagnostics;
        using System.Drawing;
        using System.Text;
        using System.Text.RegularExpressions;

        using Globus.TFlexDocs;
        using Globus.TFlexDocs.SpecificationReport;
        using Globus.TFlexDocs.SpecificationReport.GOST_2_113;
        using System.Globalization;

        using Form1Namespace;
        using System.Threading;

        //
        //  Изделие, на которое оформляется спецификация.
        //
        public class Product
        {
            //TFlexDocs _tflexdocs;

            // Само изделие
            SpecDocument _productSpecDoc;

            // Состав (документы спецификации)
            SpecDocumentCollection _documents;

            // Хэши дочерних объектов
            SortedDictionary<int, byte[]> _childrenHashes;

            protected SortedDictionary<int, byte[]> ChildrenHashes
            {
                get { return _childrenHashes; }
            }

            SpecDocumentComparer _orderComparer;
            SpecDocumentBaseEqualityComparer _authenticityComparer;
            SpecDocumentBaseEqualityComparer _authenticityComparer2;
            SpecDocumentBaseEqualityComparer _aggregateComparer;

            // Первичная применяемость
            SpecDocument _primaryUsage;
            // Атрибуты документа
            Dictionary<string, DocumentAttributes> _attributes;

            public Product(/*TFlexDocs tflexdocs*/bool sortByAlphabet)
            {
                // _tflexdocs = tflexdocs;
                _documents = new SpecDocumentCollection();
                _attributes = new Dictionary<string, DocumentAttributes>();
                // Компараторы
                _orderComparer = GetSpecDocumentOrderComparer(sortByAlphabet);
                _authenticityComparer = GetSpecDocumentAuthenticityComparer();
                _authenticityComparer2 = GetSpecDocumentAuthenticityComparer2();
                _aggregateComparer = GetSpecDocumentAggregateComparer();

                //  _hashAlgorithm = GetDocumentHashAlgorithm();
                //   _childrenHashes = new SortedDictionary<int, byte[]>();
            }




            protected virtual SpecDocumentComparer GetSpecDocumentOrderComparer(bool sortByAlphabet)
            {
                var specDoc = new SpecDocumentComparerGost2_106();
                specDoc.sortByAlphabet = sortByAlphabet;
                return specDoc;
            }

            protected virtual SpecDocumentBaseEqualityComparer GetSpecDocumentAuthenticityComparer()
            {
                return new SpecDocumentGost2_106EqualityComparer();
            }
            protected virtual SpecDocumentBaseEqualityComparer GetSpecDocumentAuthenticityComparer2()
            {
                return new SpecDocumentGost2_106EqualityComparer2();
            }

            protected virtual SpecDocumentBaseEqualityComparer GetSpecDocumentAggregateComparer()
            {
                return new SpecDocumentGost2_106AggregateComparer();
            }

            //Корневой объект
            public TFDDocument rootDocument;

            public Dictionary<int, string> ConformityClassAndBomSection = new Dictionary<int, string>()
            {
                {TFDClass.Assembly, TFDBomSection.Assembly},
                {TFDClass.Detal, TFDBomSection.Details},
                {TFDClass.OtherItem, TFDBomSection.OtherProducts},
                {TFDClass.Document, TFDBomSection.Documentation},
                {TFDClass.Komplekt, TFDBomSection.Komplekts},
                {TFDClass.Complex, TFDBomSection.Complex},
                {TFDClass.StandartItem, TFDBomSection.StdProducts},
                {TFDClass.Material, TFDBomSection.Materials},
                {TFDClass.ComplexProgram, TFDBomSection.ProgramItems},
                {TFDClass.ComponentProgram, TFDBomSection.ProgramItems},
            };

            public virtual void ReadContent(TFDDocument document, DocumentAttributes documentAttr, ProgressForm progressForm)
            {
                //            _report = new HierarchyReport(_tflexdocs, HierarchyReportConsts.HierarchyReportParameters);
                //            _report.AddDocument(document, null);

                SqlConnection connection = ApiDocs.GetConnection(true);


                string objectListQuery = String.Format
                 (@"DECLARE @DocID INT;
                        SET @DocID = {0}

                        SELECT Nomenclature.s_ObjectID,
                        Nomenclature.[Name],
                        Nomenclature.Denotation,
                        NomenclatureHierarchy.Position,
                        NomenclatureHierarchy.Amount,
                        NomenclatureHierarchy.Zone,
                        Nomenclature.Format,
                        NomenclatureHierarchy.BomSection,
                        NomenclatureHierarchy.Remarks,
                        '',                          -- AS FirstUse,
                        NomenclatureHierarchy.MeasureUnit,
                        NomenclatureHierarchy.s_ParentID,
                        Nomenclature.s_ClassID
                        FROM Nomenclature INNER JOIN NomenclatureHierarchy ON
                        (Nomenclature.s_ObjectID = NomenclatureHierarchy.s_ObjectID)  
                        WHERE Nomenclature.s_ActualVersion = 1
                          AND Nomenclature.s_Deleted = 0
                          AND NomenclatureHierarchy.s_ActualVersion = 1
                          AND NomenclatureHierarchy.s_Deleted = 0 
                          AND NomenclatureHierarchy.s_ParentID = @DocID
                          AND Nomenclature.s_ObjectID NOT IN (SELECT ConstructorParams.s_ObjectID
                                                              FROM ConstructorParams
                                                              WHERE ConstructorParams.s_ActualVersion = 1
                                                                AND ConstructorParams.s_Deleted = 0)
                          
                          UNION ALL 
                          SELECT Nomenclature.s_ObjectID,
                                 Nomenclature.[Name],
                                 Nomenclature.Denotation,
                                 NomenclatureHierarchy.Position,
                                 NomenclatureHierarchy.Amount,
                                 NomenclatureHierarchy.Zone,
                                 Nomenclature.Format,
                                 NomenclatureHierarchy.BomSection,
                                 NomenclatureHierarchy.Remarks,
                                 ConstructorParams.FirstUse,
                                 NomenclatureHierarchy.MeasureUnit,
                                 NomenclatureHierarchy.s_ParentID,
                                 Nomenclature.s_ClassID
                          FROM Nomenclature INNER JOIN NomenclatureHierarchy ON (Nomenclature.s_ObjectID = NomenclatureHierarchy.s_ObjectID)
                               JOIN ConstructorParams ON (Nomenclature.s_ObjectID = ConstructorParams.s_ObjectID)
                          WHERE Nomenclature.s_ActualVersion = 1
                            AND Nomenclature.s_Deleted = 0
                            AND ConstructorParams.s_ActualVersion = 1
                            AND ConstructorParams.s_Deleted = 0
                            AND NomenclatureHierarchy.s_ParentID = @DocID
                            AND NomenclatureHierarchy.s_ActualVersion = 1
                            AND NomenclatureHierarchy.s_Deleted = 0     
                       
                         UNION ALL                        
                         SELECT Nomenclature.s_ObjectID,
                                Nomenclature.[Name],
                                Nomenclature.Denotation,
                                1,
                                1,
                                '',
                                '',
                                '',
                                '',
                                ConstructorParams.FirstUse,
                                '',
                                '',
                                Nomenclature.s_ClassID
                          FROM Nomenclature JOIN ConstructorParams ON (Nomenclature.s_ObjectID = ConstructorParams.s_ObjectID)
                          WHERE Nomenclature.s_ActualVersion = 1
                            AND Nomenclature.s_Deleted = 0
                            AND ConstructorParams.s_ActualVersion = 1
                            AND ConstructorParams.s_Deleted = 0
                            AND Nomenclature.s_ObjectID = @DocID", document.ObjectID);

                SqlCommand objectListCommand = new SqlCommand(objectListQuery, connection);
                objectListCommand.CommandTimeout = 0;
                SqlDataReader objectListReader = objectListCommand.ExecuteReader();

                progressForm.SetStepProgressBar(10);

                Dictionary<int, TFDDocument> documentByID = new Dictionary<int, TFDDocument>();
                documentByID.Add(document.ObjectID, document);
                rootDocument = document;
                string positionNull = "";
                _productSpecDoc = CreateSpecDocument(document, documentAttr);
                //Поиск корневогообъекта в номенклатуре
                ReferenceInfo nomenclatureInfo = ReferenceCatalog.FindReference(Guids.References.Nomenclature);
                // создание объекта для работы с данными
                Reference nomenclature = nomenclatureInfo.CreateReference();
                var rootObj = nomenclature.Find(rootDocument.ObjectID);

                //Полцчение всех объектов справочника Допустимые замены
                ReferenceInfo admissibleReplacementsInfo = ReferenceCatalog.FindReference(Guids.References.AdmissibleReplacements);
                Reference admissibleReplacements = admissibleReplacementsInfo.CreateReference();
                admissibleReplacements.Objects.Load();

                List<TFDDocument> listChildDoc = new List<TFDDocument>();
                //int i = 0;
                while (objectListReader.Read())
                {
                    TFDDocument childDocument = new TFDDocument();
                    childDocument.ObjectID = objectListReader.GetInt32(0);
                  
                    if (documentByID.ContainsKey(childDocument.ObjectID))
                    {
                        rootDocument.Class = objectListReader.GetInt32(12);
                        rootDocument.FirstUse = objectListReader.GetString(9);
                        document.FirstUse = rootDocument.FirstUse;
                        continue;
                    }

                    childDocument.Naimenovanie = objectListReader.GetString(1);               
                    childDocument.Denotation = objectListReader.GetString(2);
                    positionNull = objectListReader.GetSqlInt32(3).ToString();
                    if (positionNull == "Null")
                        childDocument.Position = 0;
                    else
                        childDocument.Position = objectListReader.GetInt32(3);
                    childDocument.Amount = objectListReader.GetDouble(4);
                    childDocument.Zone = objectListReader.GetString(5);
                    if ((objectListReader.GetString(7) == TFDBomSection.Details) && (objectListReader.GetInt32(12) == TFDClass.Document))
                        childDocument.Format = "";
                    childDocument.Format = objectListReader.GetString(6);
                    childDocument.BomSection = objectListReader.GetString(7);

                    childDocument.Remarks = objectListReader.GetString(8);
                    childDocument.FirstUse = objectListReader.GetString(9);
                    childDocument.MeasureUnit = objectListReader.GetString(10);
                    childDocument.ParentID = objectListReader.GetInt32(11);
                    childDocument.Class = objectListReader.GetInt32(12);

                    if (documentAttr.programToLower && (childDocument.Class == 771 || childDocument.Class == 772))
                    {
                        childDocument.Naimenovanie = childDocument.Naimenovanie.ToLower();
                    }

                    if ((childDocument.Class == 772) || (childDocument.Class == 771) || (childDocument.BomSection == TFDBomSection.ProgramItems))
                        childDocument.Amount = 0;

                    childDocument.rootDocument = rootDocument;

                    foreach (var obj in admissibleReplacements.Objects)
                    {
                        var devicesResolvingReplace = obj.Links.ToMany[Guids.AdmissibleReplacementsPars.DevicesResolvingReplaceLink].Objects.ToList();
                        if (devicesResolvingReplace.Contains(rootObj))
                        {
                            var childObject = nomenclature.Find(childDocument.ObjectID);
                            var replacedObject = obj.Links.ToOne[Guids.AdmissibleReplacementsPars.ObjectForReplaceLink].LinkedObject;
                            if (replacedObject == childObject && childDocument.Remarks != "См. примечание")
                            {
                                if (childDocument.Remarks.Trim() == string.Empty)
                                    childDocument.Remarks = "См. примечание";
                                else
                                    childDocument.Remarks += " См. примечание";

                                childDocument.ReplacingObject = obj;
                            }
                        }
                    }
                    try
                    {

                        ReadChildDocument(childDocument, document, documentAttr);
                    }
                    catch (Exception e)
                    { 
                         MessageBox.Show("Объект " + childDocument.Naimenovanie + " " + childDocument.Denotation + "сбойный и не будет добавлен в спецификацию. Обратитесь в отдел 911.\r\n" + e,
                            "Внимание Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    listChildDoc.Add(childDocument);
                }
                List<TFDDocument> notCorrectObjects = new List<TFDDocument>();
                foreach (var child in listChildDoc)
                {
                    if (!ConformityClassAndBomSection.Contains(new KeyValuePair<int, string>(child.Class, child.BomSection)) && !child.BomSection.Contains(new NomenclatureReference().Classes.AllClasses.Find(child.Class).ToString()))
                    {
                        notCorrectObjects.Add(child);
                    }
                }

                if (notCorrectObjects.Count > 0)
                    MessageBox.Show("В процессе формирования отчета выявлено несоответствие типа объекта и его раздела спецификации:" + "\r\n" +
                    string.Join("\r\n", notCorrectObjects.Select(t => t.Format + " " + t.Denotation + " " + t.Naimenovanie + " " + " " + new NomenclatureReference().Classes.AllClasses.Find(t.Class) + " -> " + t.BomSection)),
                    "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);


                objectListReader.Close();
                connection.Close();

                foreach (SpecDocumentGost2_106 doc1 in _documents)
                {
                    foreach (SpecDocumentGost2_106 doc2 in _documents)
                    {
                        if (doc1.BomSection == doc2.BomSection && doc1.BomSection == TFDBomSection.OtherProducts && doc1.Name == doc2.Name && doc1.Designation == doc2.Designation && doc1.PositionNumber != doc2.PositionNumber)
                        {
                            doc2.Name = doc2.Name + "###";
                        }
                    }
                }
                // Сортировка документов
                _documents.Sort(_orderComparer);
            }

            public static readonly List<string> sectionCategories = new List<string> {
            TFDBomSection.Documentation,
            TFDBomSection.Complex,
            TFDBomSection.Assembly,
            TFDBomSection.Details,
            TFDBomSection.StdProducts,
            TFDBomSection.OtherProducts,
            TFDBomSection.Materials,
            TFDBomSection.Komplekts,
            TFDBomSection.Components,
            TFDBomSection.ProgramItems};

            /**
             *  ProcessDocument child document
             */
            protected virtual void ReadChildDocument(TFDDocument doc, TFDDocument parent, DocumentAttributes documentAttr)
            {
                DocumentAttributes attributes = new DocumentAttributes(doc);
                SpecDocumentGost2_106 child = CreateSpecDocument(doc, documentAttr);
                // ВОЗМОЖНО НАДО ИСКЛЮЧИТЬ
                string name = doc.Naimenovanie;//_tflexdocs.GetParameterValue(doc, ParametersGUIDs.Name);
                if (!_attributes.ContainsKey(name))
                {
                    _attributes.Add(name, attributes);
                }

                _documents.Add(child);

            }

            protected virtual SpecDocumentGost2_106 CreateSpecDocument(TFDDocument doc, DocumentAttributes documentAttr)
            {
                SpecDocumentGost2_106 child = new SpecDocumentGost2_106(/*_tflexdocs,*/ doc, documentAttr /*, _productSpecDoc*/);
                return child;
            }

            /**
             *  Получение записей спецификации формы 1, 1а по ГОСт 2.106
             */
            public virtual void AddSpecRowsForm1(SpecForm1 specForm, bool fullNames, DocumentAttributes documentAttr, string docName, string denotation)
            {
                // Создание списка документов
                SpecDocumentGost2_106[] documents = (SpecDocumentGost2_106[])_documents.ToArray();

                // Добавление строк таблицы спецификации
                AddSpecSectionsAndRows(specForm, documents, documentAttr, fullNames, docName, denotation, false);
            }

            /**
             * Получение строк таблицы спецификации для одного документа
             * с учетом уже занесенных в таблицу документов, а также
             * документов, находящихся в очереди.
             */
            protected virtual SpecRowForm1[] CreateSpecRows(
                SpecDocumentGost2_106 document,
                List<SpecDocumentGost2_106> addedDocuments,
                List<SpecDocumentGost2_106> queueDocuments,
                SpecForm1 specForm,
                bool fullNames, string sectionName, bool zeroToDash, string docName, string denotation, DocumentAttributes documentAttr)
            {

                if (document.Name.Contains("###"))
                {
                    // MessageBox.Show(document.Name);
                    document.Name = document.Name.Remove(document.Name.Length - 3, 3);

                }
                // Очередной документ должен находиться в очереди документов
                Debug.Assert(queueDocuments.Contains(document));

                // Составление списка добавляемых в виде одной записи документов
                List<SpecDocumentGost2_106> addingDocuments = new List<SpecDocumentGost2_106>();
                addingDocuments.Add(document);
                foreach (SpecDocumentGost2_106 queueDocument in queueDocuments)
                {
                    if (queueDocument == document)
                        continue;
                    if (_aggregateComparer.Equals(document, queueDocument))
                        addingDocuments.Add(queueDocument);
                }

                // Удаление из очереди тех документов, которые сейчас будут добавлены
                foreach (SpecDocumentGost2_106 d in addingDocuments)
                    queueDocuments.Remove(d);

                // Получение значений полей спецификации на основе нескольких документов
                string format;
                string zone;
                string position;
                string designation;
                string name;
                string count;
                string countUnit;
                string refdes;
                string note;
                string noteSuperscriptSymbol;

                GetSpecForm1Fields(addingDocuments, addedDocuments, fullNames,
                    out format, out zone, out position, out designation, out name,
                    out count, out countUnit, out refdes, out note,
                    out noteSuperscriptSymbol, zeroToDash, documentAttr);

                // Добавление в список добавленных документов тех, которые были обработаны
                foreach (SpecDocumentGost2_106 d in addingDocuments)
                    addedDocuments.Add(d);

                SpecRowForm1[] record = SpecRowForm1.CreateRecord(format, zone, position,
                    designation, name, count, countUnit, refdes, note, noteSuperscriptSymbol,
                    specForm.TextFormatter, sectionName, fullNames, docName, denotation, documentAttr);
                return record;
            }

            /// Получение значений граф спецификации на основе списка документов.
            /// Документы списка подлежат объединению в одну запись спецификации
            /// (т.е. у них должны быть одинаковые позиции, наименование, обозначение
            /// и раздел спецификации).
            protected virtual void GetSpecForm1Fields(List<SpecDocumentGost2_106> addingDocuments,
                List<SpecDocumentGost2_106> addedDocuments,
                bool fullNames,
                out string format, out string zone, out string position,
                out string designation, out string name, out string count, out string countUnit,
                out string refdes, out string note, out string noteSuperscriptSymbol, bool zeroToDash, DocumentAttributes documentAttr)
            {
                List<string> resultFormats = new List<string>();
                List<string> resultZones = new List<string>();
                string resultPosition = null;
                string resultNote = "";
                string resultDesignation = null;
                string resultName = null;
                string resultNoteSuperscriptSymbol = null;
                List<string> resultRefdesList = new List<string>();
                double? totalCount = null;
                string resultCountUnit = "";
                string extraInfo = "";

                foreach (SpecDocumentGost2_106 document in addingDocuments)
                {
                    // Формат
                    foreach (string currentFormat in document.Formats)
                    {
                        if (!resultFormats.Contains(currentFormat))
                            resultFormats.Add(currentFormat);
                    }

                    // Зона
                    foreach (string currentZone in document.Zones)
                    {
                        if (currentZone != "МЭ" && !resultZones.Contains(currentZone))
                        {
                            resultZones.Add(currentZone);
                        }
                    }

                    // Номер позиции

                    string currentPosition = GetSpecDocumentPositionField(document, addedDocuments, zeroToDash);
                    if (resultPosition == null)
                        resultPosition = currentPosition;
                    else if (resultPosition != currentPosition)
                        resultPosition = "(?)";

                    // Обозначение
                    string currentDesignation = GetSpecDocumentDesignationField(document, addedDocuments); ;
                    if (resultDesignation == null)
                        resultDesignation = GetSpecDocumentDesignationField(document, addedDocuments);
                    else if (resultDesignation != currentDesignation)
                        resultDesignation += @"\n(?)" + currentDesignation;

                    // Наименование
                    string currentName = GetSpecDocumentNameField(document, fullNames, documentAttr);
                    if (resultName == null)
                        resultName = GetSpecDocumentNameField(document, fullNames, documentAttr);
                    else if (resultName != currentName)
                        resultName += @"\n(?)" + currentName;

                    // Количество
                    if (document.Count != null)
                    {
                        if (totalCount == null)
                            totalCount = 0.0;
                        totalCount += document.Count;
                    }

                    // Позиционное обозначение на схеме
                    if (document.RefDes != null && document.RefDes.Length > 0)
                    {
                        resultRefdesList.Add(document.RefDes);
                    }

                    // Примечание
                    if (document.Note != null && document.Note.Length > 0)
                    {
                        if (resultNote.Length > 0)
                            resultNote += "\n";
                        resultNote += document.Note;
                    }

                    // Символ первичной применяемости
                    if (resultNoteSuperscriptSymbol == null)
                        resultNoteSuperscriptSymbol = document.PrimaryUsageSymbol;
                    else if (resultNoteSuperscriptSymbol != document.PrimaryUsageSymbol)
                    {
                        // Попытка объединить две записи спецификации с различными символами первичной применяемости
                        resultNoteSuperscriptSymbol = "(?)";
                    }

                    // Единица измерения количества
                    if (document.CountUnit != null && document.CountUnit.Length > 0)
                    {
                        if (resultCountUnit.Length == 0)
                            resultCountUnit = document.CountUnit;
                        else if (resultCountUnit != document.CountUnit)
                            resultCountUnit += "; " + document.CountUnit + "(?)";
                    }

                    // Дополнительные сведения для спецификации
                    if (document.ExtraInfo != null && document.ExtraInfo.Length > 0)
                    {
                        if (extraInfo.Length > 0)
                            extraInfo += @"\n";
                        extraInfo += document.ExtraInfo;
                    }
                }

                // Формат
                string formatNote = null;
                if (resultFormats.Count == 0)
                    format = "";
                else if (resultFormats.Count == 1)
                    format = resultFormats[0];
                else
                {
                    // В соответствие с п. 3.6.3 ОСТ 4Г 0.000.254-84,
                    // на который ссылается ОРУК ГН0.070.003 п. 2.1.8:
                    // "При разработке конструкторской документации
                    // следует руководствоваться ЕСКД, а также ОСТ 4Г 0.000.254-84."
                    format = "*";
                    formatNote = String.Join(", ", resultFormats.ToArray());
                }

                // Зона
                string zoneNote = null;
                if (resultZones.Count == 0)
                    zone = "";
                else if (resultZones.Count == 1)
                    zone = resultZones[0];
                else
                {
                    zone = "*";
                    zoneNote = String.Join(", ", resultZones.ToArray());
                }

                // Номер позиции
                position = resultPosition;

                // Обозначение
                designation = resultDesignation;

                // Наименование
                name = resultName;
                if (extraInfo.Length > 0)
                    name += @"\n" + extraInfo;

                // Количество
                if (totalCount == null)
                    count = "";
                else
                    count = ((double)totalCount).ToString("0.#####");

                // Единица измерения количества
                countUnit = resultCountUnit;

                // Позиционное обозначение на схеме
                refdes = "";
                if (resultRefdesList.Count > 0)
                {
                    refdes = GetRefdesNote(resultRefdesList);
                }

                // Примечание
                StringBuilder noteBuilder = new StringBuilder();
                if (formatNote != null)
                {
                    if (noteBuilder.Length > 0)
                        noteBuilder.Append('\n');
                    noteBuilder.Append("*" + formatNote);
                }
                if (zoneNote != null && zoneNote.Length > 0)
                {
                    if (noteBuilder.Length > 0)
                        noteBuilder.Append('\n');
                    noteBuilder.Append("*" + zoneNote);
                }
                if (resultNote != null && resultNote.Length > 0)
                {
                    if (noteBuilder.Length > 0)
                        noteBuilder.Append('\n');
                    noteBuilder.Append(resultNote);
                }
                note = noteBuilder.ToString();

                // Символ первичной применяемости
                if (resultNoteSuperscriptSymbol == null)
                    noteSuperscriptSymbol = "";
                else
                    noteSuperscriptSymbol = resultNoteSuperscriptSymbol;
            }

            protected virtual string GetSpecDocumentPositionField(SpecDocumentGost2_106 document,
                List<SpecDocumentGost2_106> addedDocuments, bool zeroToDash)
            {
                string position = "";

                SpecDocumentGost2_106 prevDoc = null;
                if (addedDocuments.Count > 0)
                    prevDoc = addedDocuments[addedDocuments.Count - 1];

                if ((document.PositionNumber == null) || (document.PositionNumber == 0))
                {
                    // Позиция не указана
                    if (zeroToDash && ((document.BomSection == TFDBomSection.OtherProducts) ||
                        (document.BomSection == TFDBomSection.StdProducts) ||
                        (document.BomSection == TFDBomSection.Details)))
                        position = "-";
                    else position = "";

                }

                else
                {
                    /*if (document.BomGroup.IsContainedInGroup(BomGroupID.Documentation)
                        || document.BomGroup.IsContainedInGroup(BomGroupID.Complements))*/
                    if ((document.BomSection == TFDBomSection.Documentation) || (document.BomSection == TFDBomSection.Komplekts))
                    {
                        // Для раздела Документация и Комплекты графа Поз. не заполняется
                        position = "";
                    }
                    else if (prevDoc != null && document.PositionNumber == prevDoc.PositionNumber)
                    {
                        // Прочерк вместо повторяющегося номера позиции
                        position = "";
                    }
                    else
                        position = ((int)document.PositionNumber).ToString();
                }
                return position;
            }


            protected virtual string GetSpecDocumentDesignationField(SpecDocumentGost2_106 document,
                List<SpecDocumentGost2_106> addedDocuments)
            {
                return document.Designation;
            }

            protected virtual string GetSpecDocumentNameField(SpecDocumentGost2_106 document, bool fullName, DocumentAttributes documentAttr)
            {
                string name = document.Name;

                //if (document.BomGroup.IsContainedInGroup(BomGroupID.Documentation)
                //    && !fullName)
                if ((document.BomSection == TFDBomSection.Documentation) && !fullName)
                {

                    /*
                     * в разделе "Документация" для документов, входящих в основной комплект
                        документов специфицируемого изделия и составляемых на данное изделие,
                        - только наименование документов, например: "Сборочный чертеж",
                        "Габаритный чертеж", "Технические условия". Для документов на
                        неспецифицированные составные части - наименование изделия и
                        наименование документа;
                     */
                    if (document.Designation.StartsWith(Designation)
                        && document.Name.StartsWith(Name))
                    {
                        // Регулярное выражение для исполнения корневого объекта
                        string nameMake = @"^(?<g0>" + Name + ")-(?<g1>[0-9]{1,})";
                        Regex regexNameMake = new Regex(nameMake);
                        Match matchNameMake;

                        matchNameMake = regexNameMake.Match(document.Name);
                        // Если этот документ для исполнения то 
                        if (regexNameMake.Match(document.Name).Success)
                        {
                            name = document.Name.Substring(matchNameMake.Value.ToString().Length);
                            name = name.Trim();
                        }
                        else
                        {
                            name = document.Name.Substring(Name.Length);
                            name = name.Trim();
                        }
                    }

                }

                return name;
            }

            /**
             * Создание строки примечания с перечнем позиционных обозначений на схеме
             */
            string GetRefdesNote(IList<string> refdesStrList)
            {
                StringBuilder result = new StringBuilder();

                // Получение отсортированного списка обозначений
                SortedDictionary<Refdes, object> sortedRefdeses
                    = new SortedDictionary<Refdes, object>();
                foreach (string refdesStr in refdesStrList)
                {
                    Refdes refdes = new Refdes(refdesStr);
                    if (sortedRefdeses.ContainsKey(refdes))
                    {
                        Refdes errorRefdesStr = new Refdes(refdesStr + "(?)");
                        if (!sortedRefdeses.ContainsKey(errorRefdesStr))
                            sortedRefdeses.Add(errorRefdesStr, null);
                    }
                    else
                    {
                        sortedRefdeses.Add(refdes, null);
                    }
                }

                return GetRefdesNote(sortedRefdeses.Keys);
            }

            /// Получение примечания с элементами вида R1..R3, R4, R5, R9..R11
            string GetRefdesNote(IEnumerable<Refdes> sortedRefdeses)
            {
                StringBuilder result = new StringBuilder();
                List<Refdes> chain = new List<Refdes>();
                IEnumerator<Refdes> refdesEnum = sortedRefdeses.GetEnumerator();
                while (true)
                {
                    // Получение очередного элемента
                    bool next = refdesEnum.MoveNext();
                    Refdes currentRefdes = next ? refdesEnum.Current : null;

                    bool clearChain = false;

                    // Обработка очередного полученного элемента обозначения
                    if (next)
                    {
                        if (chain.Count == 0)
                            chain.Add(currentRefdes);
                        else
                        {
                            Refdes lastChainRefdes = chain[chain.Count - 1];
                            Refdes nextPossibleChainRefdes = lastChainRefdes.IncreaseLastNumber();
                            if (nextPossibleChainRefdes != null)
                            {
                                if (nextPossibleChainRefdes.Equals(currentRefdes))
                                {
                                    // Очередное обозначение является продложением цепочки
                                    chain.Add(currentRefdes);
                                }
                                else
                                {
                                    // Очередное обозначение не является продолжением цепочки.
                                    // Прерывание цепочки и запись результата, создание новой цепочки
                                    clearChain = true;
                                }
                            }
                        }
                    }
                    else // !next
                    {
                        clearChain = true;
                    }

                    if (clearChain)
                    {
                        // Очистка цепочки
                        if (chain.Count > 0)
                        {
                            // Цепочка имеет элемент(ы)
                            if (result.Length > 0)
                                result.Append(", ");

                            if (chain.Count == 1)
                                result.Append(chain[0]);
                            else if (chain.Count == 2)
                                result.Append(String.Format("{0}, {1}", chain[0], chain[1]));
                            else
                                result.Append(String.Format("{0}…{1}", chain[0], chain[chain.Count - 1]));
                            chain.Clear();
                        }

                        // Добавление в цепочку очередного элемента
                        if (next)
                            chain.Add(currentRefdes);
                    }
                    // Выход, если нет очередного элемента
                    if (!next)
                        break;
                }
                return result.ToString();
            }

            void AddBomGroupHeader(SpecForm1 specForm, string title, bool newPage, bool firstHeader)
            {
                if (newPage)
                    specForm.NewPage();

                // переход на новую страницу выполняется, если на странице
                // осталось менее 10 пустых строк
                if (specForm.CurrentPageAvialableRows < 10)
                    specForm.NewPage();

                // Пустая строка перед заголовком раздела
                if (!firstHeader)
                    specForm.AddEmptyRow();

                // Заголовок раздела
                SpecFormCell headerCell = new SpecFormCell(title,
                    SpecFormCell.UnderliningFormat.Underlined, SpecFormCell.AlignFormat.Center);
                SpecRowForm1 header = new SpecRowForm1(SpecFormCell.Empty, headerCell);
                specForm.Add(header);

                // Пустая строка после заголовка раздела
                specForm.AddEmptyRow();
            }

            void NewPagesBomGroupHeader(SpecForm1 specForm, string title, bool Null)
            {
                if (Null)
                {
                    specForm.NewPage();
                    if (specForm.CurrentPageAvialableRows < 28)
                        specForm.NewPage();

                    // Пустая строка перед заголовком радела
                    // specForm.AddEmptyRow();

                    // Заголовок раздела
                    SpecFormCell headerCell = new SpecFormCell(title,
                        SpecFormCell.UnderliningFormat.Underlined, SpecFormCell.AlignFormat.Center);
                    SpecRowForm1 header = new SpecRowForm1(SpecFormCell.Empty, headerCell);
                    specForm.Add(header);

                    // Пустая строка после заголовка раздела
                    specForm.AddEmptyRow();
                }
            }

            /**
             * Добавление наименования комплекта при записи его содержимого непосредственно
             * в спецификацию изделия без составления спецификации комплекта
             */
            /* void AddCompelmentName(SpecDocumentGost2_106 document,
                 List<SpecDocumentGost2_106> addedDocuments,
                 List<SpecDocumentGost2_106> queueDocuments,
                 SpecForm1 specForm)
             {
                 // Обработка только документов из состава раздела Комплекты
                 // if (!document.BomGroup.IsContainedInGroup(BomGroupID.Complements))
                 if (document.BomSection != TFDBomSection.Komplekts)
                     return;

                 // Получение подраздела (непосредственно)
                 string subGroup = document.BomSection;
                 while (subGroup != TFDBomSection.Komplekts)//(int)BomGroupID.Complements)
                 {
                     subGroup = document._parent.BomSection;
                     if (subGroup == null)
                         return;
                 }

                 // Заголовок добавляется перед добавлением первой записи в подразделе
                 foreach (SpecDocumentGost2_106 otherDocument in addedDocuments)
                 {
                     //if (otherDocument.BomGroup.IsContainedInGroup(subGroup))
                     if (otherDocument.BomSection == subGroup)
                         return;
                 }

                 // Заголовок добавляется только в случае, если в подразделе более одной записи
                 bool moreThanOne = false;
                 foreach (SpecDocumentGost2_106 otherDocument in queueDocuments)
                 {
                     if (otherDocument != document
                         && otherDocument.BomSection == subGroup)//otherDocument.BomGroup.IsContainedInGroup(subGroup))
                     {
                         moreThanOne = true;
                         break;
                     }
                 }
                 if (!moreThanOne)
                     return;

                 // Получение наименования комплекта
                 string complementName = subGroup;//subGroup.Name;
                 if (complementName.Length < 1)
                     return;

                 // Первая буква наименования комплекта - заглавная
                 complementName = complementName.Substring(0, 1).ToUpper() + complementName.Substring(1);

                 // Получение строк наименования комплекта
                 SpecRowForm1[] rows = SpecRowForm1.CreateRecord("", "", "", "", complementName, "", "", "", "", "",
                     specForm.TextFormatter);

                 // Добавление строк к форме
                 specForm.AddEmptyRowsOrFinishCurrentPage(0);
                 if (specForm.CurrentPageAvialableRows < rows.Length + 10)
                     specForm.NewPage();
                 specForm.Add(rows);
                 specForm.AddEmptyRowsOrFinishCurrentPage(0);
             }*/

            /**
             *  Получение записей спецификации формы 1, 1а по ГОСт 2.106 для
             *  массива документов
             *  При создании записей СП поддерживается раздел "Устанавливают по БФМИ... МЭ"
             */
            protected virtual void AddSpecSectionsAndRows(SpecForm1 specForm, SpecDocumentGost2_106[] documents, DocumentAttributes documentAttr,
                bool fullNames, string docName, string denotation, bool varPartDoc)
            {
                // Разделение списка документов documents на два:
                // - без МЭ в поле Зоны (идут в СП как обычно по ЕСКД)
                // - с МЭ в поле Зоны (идут в раздел "Устанавливают по БФМИ... МЭ")
                List<SpecDocumentGost2_106> basicDocuments = new List<SpecDocumentGost2_106>();
                List<SpecDocumentGost2_106> meDocuments = new List<SpecDocumentGost2_106>();
                foreach (SpecDocumentGost2_106 document in documents)
                {


                    if (document.Zones.Length >= 1 && document.Zones[0] == "МЭ")
                    {
                        meDocuments.Add(document);
                    }
                    else
                    {
                        basicDocuments.Add(document);
                    }
                }

                if (meDocuments.Count == 0)
                {
                    // Создание спецификации без раздела "Устанавливают по БФМИ... МЭ"
                    AddSpecSectionsAndRowsNoME(specForm, documents, fullNames, documentAttr, true, docName, denotation, varPartDoc);
                }
                else
                {
                    // Разделы по ЕСКД
                    AddSpecSectionsAndRowsNoME(specForm, basicDocuments.ToArray(), fullNames, documentAttr, true, docName, denotation, varPartDoc);

                    // Создание спецификации с разделом "Устанавливают по БФМИ... МЭ"

                    // Заголовок раздела "Устанавливают по БФМИ... МЭ"
                    if (!specForm.AtTopOfPage())
                        specForm.NewPage();

                    specForm.AddEmptyRow();
                    SpecFormCell headerCell1 = new SpecFormCell("Устанавливают по",
                        SpecFormCell.UnderliningFormat.Underlined, SpecFormCell.AlignFormat.Center);
                    SpecRowForm1 header1 = new SpecRowForm1(SpecFormCell.Empty, headerCell1);
                    specForm.Add(header1);
                    SpecFormCell headerCell2 = new SpecFormCell(_productSpecDoc.Designation + " МЭ",
                        SpecFormCell.UnderliningFormat.Underlined, SpecFormCell.AlignFormat.Center);
                    SpecRowForm1 header2 = new SpecRowForm1(SpecFormCell.Empty, headerCell2);
                    specForm.Add(header2);
                    specForm.AddEmptyRow();

                    // Разделы
                    //  ProductComparer pc = new ProductComparer();
                    //   meDocuments.Sort(pc);

                    AddSpecSectionsAndRowsNoME(specForm, meDocuments.ToArray(), fullNames, documentAttr, true, docName, denotation, varPartDoc);
                }

                int i = 1;
                foreach (var replacingObj in documentAttr.ReplacingObjects)
                {
                    specForm.AddEmptyRow();
                    specForm.AddEmptyRow();
                    string remarkStr = "Примечание";
                    if (documentAttr.ReplacingObjects.Count > 1)
                        remarkStr = "Примечание " + i;
                    SpecFormCell remarkCell = new SpecFormCell(remarkStr,
                      SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Left);
                    SpecFormCell emptyCell = new SpecFormCell("",
                          SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Left);
                    SpecRowForm1 firstRow = new SpecRowForm1(remarkCell, emptyCell);
                    specForm.Add(firstRow);
                    SpecFormCell admissibleCell = new SpecFormCell("Допускается замена",
                      SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Left);
                    SpecFormCell replacedObjCell = new SpecFormCell(replacingObj.Links.ToOne[Guids.AdmissibleReplacementsPars.ObjectForReplaceLink].LinkedObject + " на",
                          SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Left);
                    SpecRowForm1 secondRow = new SpecRowForm1(admissibleCell, replacedObjCell);
                    specForm.Add(secondRow);
                    SpecFormCell replacingObjectCell = new SpecFormCell(replacingObj.Links.ToMany[Guids.AdmissibleReplacementsPars.ReplacingObjectLink].Objects.ToList().FirstOrDefault() + "",
                    SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Left);

                    SpecRowForm1 thirddRow = new SpecRowForm1(replacingObjectCell, emptyCell);
                    specForm.Add(thirddRow);

                }
            }

            bool materialConstPart = true;
            /**
             *  Получение записей спецификации формы 1, 1а по ГОСт 2.106 для
             *  массива документов.
             *  При создании записей СП не поддерживается раздел "Устанавливают по БФМИ... МЭ"
             */
            protected virtual void AddSpecSectionsAndRowsNoME(SpecForm1 specForm, SpecDocumentGost2_106[] documents,
                bool fullNames, DocumentAttributes documentAttr, bool RezervPages, string docName, string denotation, bool varPartDoc)
            {
                // Занесенные в спецификацию документы
                List<SpecDocumentGost2_106> addedDocuments = new List<SpecDocumentGost2_106>();

                List<string> otherItemsName = new List<string>();

                otherItemsName.Add("Резистор");
                otherItemsName.Add("Конденсатор");
                otherItemsName.Add("Микросхема");

                string lastNameOtherItems = string.Empty;
                int? lastPosition = -1;
                bool firstObject = false;

                // Документы, подлежащие занесению в спецификацию
                List<SpecDocumentGost2_106> queueDocuments = new List<SpecDocumentGost2_106>(documents);

                bool firstHeader = true;
                // BomGroup[] sections = _tflexdocs.BomGroupsReference.GetChildrenGroups((int)BomGroupID.Specification);

                // Цикл по разделам спецификации
                //foreach (BomGroup section in sections)
                for (int i = 0; i < sectionCategories.Count; i++)
                {
                    bool headerCreated = false;
                    bool SectionNull = true;
                    // Поиск и обработка очередного документа из состава текущего раздела спецификации
                    while (true)
                    {
                        // Поиск очередного документа из текущего раздела
                        SpecDocumentGost2_106 currentDocument = null;
                        foreach (SpecDocumentGost2_106 qdoc in queueDocuments)
                        {

                            //if (qdoc.BomGroup.IsContainedInGroup(section))
                            if (qdoc.BomSection == sectionCategories[i])
                            {
                                currentDocument = qdoc;

                                //         specForm.AddEmptyRows(currentDocument.EmptyLinesAfter);
                                //         specForm.AddEmptyRows(currentDocument.EmptyLinesAfter);
                                SectionNull = false;
                                break;
                            }
                        }
                        if (currentDocument == null)
                            break;

                        // Создание заголовка
                        if (!headerCreated)
                        {
                            bool newPage;
                            if (firstHeader)
                                newPage = false; // первый заголовок
                            else
                                newPage = true;
                            //newPage = documents.Length > 3     // количество документов больше трех
                            //          || specForm.CurrentPageAvialableRows < 10; // осталось меньше 10 строк на  листе
                            if (sectionCategories[i] == TFDBomSection.Komplekts)
                            {
                                AddBomGroupHeader(specForm, sectionCategories[i], newPage, firstHeader);
                                if (documentAttr.addSectionProgramSoft)
                                {
                                    // Заголовок раздела
                                    SpecFormCell textCell = new SpecFormCell("Программные изделия",
                                                                             SpecFormCell.UnderliningFormat.
                                                                                 NotUnderlined,
                                                                             SpecFormCell.AlignFormat.Center);
                                    SpecRowForm1 textHeader = new SpecRowForm1(SpecFormCell.Empty, textCell);
                                    specForm.Add(textHeader);

                                    // Пустая строка после заголовка раздела
                                    specForm.AddEmptyRow();
                                }
                            }
                            else
                                AddBomGroupHeader(specForm, sectionCategories[i], newPage, firstHeader);

                            headerCreated = true;

                        }

                        // Запись наименования комплекта (при необходимости), п. 3.9 ГОСТ 2.106
                        //  AddCompelmentName(currentDocument, addedDocuments, queueDocuments, specForm);

                        int indx = -1;

                        if ((sectionCategories[i] == TFDBomSection.OtherProducts) && (documentAttr.newPageOIAfterKindProducts) && !varPartDoc)
                            foreach (string name in otherItemsName)
                            {
                                indx = currentDocument.Name.IndexOf(name);
                                if (lastNameOtherItems == string.Empty)
                                    firstObject = true;
                                else
                                    firstObject = false;
                                // Вставка новой страницы в прочих изделиях после группы "Резисторы", "Микросхемы" и "Конденсаторы" 
                                if ((lastNameOtherItems == name) && !currentDocument.Name.Contains(name))
                                {
                                    specForm.NewPage(false);
                                    lastNameOtherItems = currentDocument.Name;
                                }

                                if ((currentDocument.PositionNumber != 0) && (lastPosition == 0))
                                    specForm.NewPage(false);

                                // Вставка новой страницы в прочих изделиях до группы "Резисторы", "Микросхемы" и "Конденсаторы"
                                if (((indx == 0) && (name != lastNameOtherItems)) || ((indx == 0) && (currentDocument.PositionNumber == 0) && (lastPosition != 0)))
                                {

                                    if (!firstObject)
                                        specForm.NewPage(false);
                                    lastNameOtherItems = name;
                                    lastPosition = currentDocument.PositionNumber;
                                    break;
                                }

                            }
                        if ((sectionCategories[i] == TFDBomSection.OtherProducts) && (documentAttr.newPageOIAfterKindProductsVariosPart) && varPartDoc)
                            foreach (string name in otherItemsName)
                            {
                                indx = currentDocument.Name.IndexOf(name);
                                if (lastNameOtherItems == string.Empty)
                                    firstObject = true;
                                else
                                    firstObject = false;
                                // Вставка новой страницы в прочих изделиях после группы "Резисторы", "Микросхемы" и "Конденсаторы" 
                                if ((lastNameOtherItems == name) && !currentDocument.Name.Contains(name))
                                {
                                    specForm.NewPage(false);
                                    lastNameOtherItems = currentDocument.Name;
                                }

                                if ((currentDocument.PositionNumber != 0) && (lastPosition == 0))
                                    specForm.NewPage(false);

                                // Вставка новой страницы в прочих изделиях до группы "Резисторы", "Микросхемы" и "Конденсаторы"
                                if (((indx == 0) && (name != lastNameOtherItems)) || ((indx == 0) && (currentDocument.PositionNumber == 0) && (lastPosition != 0)))
                                {

                                    if (!firstObject)
                                        specForm.NewPage(false);
                                    lastNameOtherItems = name;
                                    lastPosition = currentDocument.PositionNumber;
                                    break;
                                }

                            }
                        lastPosition = currentDocument.PositionNumber;

                        // Создание записей. Коллекции addedDocuments, queueDocuments модифицируются
                        SpecRowForm1[] rows = CreateSpecRows(currentDocument, addedDocuments, queueDocuments, specForm,
                            fullNames, sectionCategories[i], documentAttr.ZeroToDash, docName, denotation, documentAttr);

                        // Заполнение таблицы спецификации
                        if (specForm.CurrentPageAvialableRows <= rows.Length + 1)
                            specForm.NewPage();
                        /* if (specForm.AtTopOfPage())
                            specForm.AddEmptyRows(1);*/

                        if ((currentDocument.EmptyLinesBefore > 0) && (!firstHeader))
                            specForm.AddEmptyRows(currentDocument.EmptyLinesBefore);

                        firstHeader = false;

                        specForm.Add(rows);

                        // Сразу после добавления записи выполняем замену повторяющейся базовой части
                        // обозначения в графе Обозначение
                        SpecRowForm1[] currentPageRows = specForm.GetRows(specForm.TotalPageNumber - 1);
                        // ReplaceRepeatedDesignationBasePart(currentPageRows)


                        // Одна пустая строка после записи или переход на новую страницу
                        if (currentDocument.EmptyLinesAfter > 0)
                            specForm.AddEmptyRows(currentDocument.EmptyLinesAfter);
                        else
                            specForm.AddEmptyRowsOrFinishCurrentPage(0);
                    }




                    if (RezervPages && varPartDoc == false)
                    {
                        if (documentAttr.doc && (sectionCategories[i] == TFDBomSection.Documentation))//section.ID == 50)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.DocumentationPages);
                        }
                        if (documentAttr.assem && sectionCategories[i] == TFDBomSection.Assembly)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.AssemblyPages);
                        }
                        if (documentAttr.det && sectionCategories[i] == TFDBomSection.Details)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.DetailsPages);
                        }
                        if (documentAttr.std && sectionCategories[i] == TFDBomSection.StdProducts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.StdProductsPages);
                        }
                        if (documentAttr.oth && sectionCategories[i] == TFDBomSection.OtherProducts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.OtherProductsPages);
                        }
                        if (documentAttr.mat && sectionCategories[i] == TFDBomSection.Materials && materialConstPart)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.MaterialsPages);
                            materialConstPart = false;
                        }
                        if (documentAttr.kom && sectionCategories[i] == TFDBomSection.Komplekts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.KomplektPages);
                        }
                    }

                    if (RezervPages && varPartDoc)
                    {
                        if (documentAttr.docVarPart && (sectionCategories[i] == TFDBomSection.Documentation))//section.ID == 50)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.DocumentationPagesVarPart);
                        }
                        if (documentAttr.assemVarPart && sectionCategories[i] == TFDBomSection.Assembly)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.AssemblyPagesVarPart);
                        }
                        if (documentAttr.detVarPart && sectionCategories[i] == TFDBomSection.Details)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.DetailsPagesVarPart);
                        }
                        if (documentAttr.stdVarPart && sectionCategories[i] == TFDBomSection.StdProducts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.StdProductsPagesVarPart);
                        }
                        if (documentAttr.othVarPart && sectionCategories[i] == TFDBomSection.OtherProducts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.OtherProductsPagesVarPart);
                        }
                        if (documentAttr.matVarPart && sectionCategories[i] == TFDBomSection.Materials && materialConstPart)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.MaterialsPagesVarPart);
                            materialConstPart = false;
                        }
                        if (documentAttr.komVarPart && sectionCategories[i] == TFDBomSection.Komplekts)
                        {
                            NewPagesBomGroupHeader(specForm, sectionCategories[i], SectionNull);
                            specForm.NewPage(documentAttr.KomplektPagesVarPart);
                        }
                    }
                }
                specForm.NewPage(false);
            }

            // Рег. выражение, соответствующее знакам: тире, номер исполнения и код документа
            static readonly Regex _baseAndExtDesignationPartsRegex = new Regex(
                @"(?<base>.+?)(?<ext>|-[0-9]{2,3}(?:[А-Я]+[0-9]*)?)$",
                RegexOptions.Compiled);

            /**
             * Замена повторяющейся базовой части обозначения
             * 
             * @param rows
             *      Строки спецификации, в которых выполняется замена
             */
            void ReplaceRepeatedDesignationBasePart(SpecRowForm1[] rows)
            {
                string lastBaseDesignationPart = null;

                // Цикл по строкам спецификации
                foreach (SpecRowForm1 row in rows)
                {
                    // Получение очередного обозначения из графы Прим.
                    string currentDesignation = row.Designation.Text;
                    if (currentDesignation == null || currentDesignation.Length == 0)
                        continue;

                    // Выделение из текущего обозначения базовой и дополнительной части
                    Match currentDesignationMatch = _baseAndExtDesignationPartsRegex.Match(currentDesignation);
                    if (currentDesignationMatch.Success
                        && currentDesignationMatch.Groups["base"].Success
                        && currentDesignationMatch.Groups["ext"].Success)
                    {
                        // Успешно выделена базовая и дополнительная части текущего обозначения
                        string currentBase = currentDesignationMatch.Groups["base"].Value;

                        if (lastBaseDesignationPart == null)
                        {
                            lastBaseDesignationPart = currentBase;
                        }
                        else if (currentBase == lastBaseDesignationPart)
                        {
                            StringBuilder spaces = new StringBuilder();
                            for (int i = 0; i < lastBaseDesignationPart.Length; ++i)
                                spaces.Append(" "); // неразрывный пробел
                            string newDesignation = spaces.ToString() + currentDesignationMatch.Groups["ext"].Value;
                            row.Designation.Text = newDesignation;
                        }
                        else
                            lastBaseDesignationPart = currentBase;
                    }
                }
            }

            public DocumentAttributes GetAttributes(string name)
            {
                if (_attributes.ContainsKey(name))
                    return _attributes[name];
                else
                    return DocumentAttributes.Empty;
            }

            public string Name
            {
                get { return _productSpecDoc.Name; }
            }

            public string Designation
            {
                get { return _productSpecDoc.Designation; }
            }

            public SpecDocument SpecDocument
            {
                get { return _productSpecDoc; }
            }

            public SpecDocument PrimaryUsage
            {
                get { return _primaryUsage; }
            }

            public SpecDocumentCollection Documents
            {
                get { return _documents; }
            }

            protected SpecDocumentComparer OrderComparer
            {
                get { return _orderComparer; }
            }

            protected SpecDocumentBaseEqualityComparer AuthenticityComparer
            {
                get { return _authenticityComparer; }
            }
            protected SpecDocumentBaseEqualityComparer AuthenticityComparer2
            {
                get { return _authenticityComparer2; }
            }
            protected SpecDocumentBaseEqualityComparer AggregateComparer
            {
                get { return _aggregateComparer; }
            }
        }

        //
        // Сравнение документов для определения их порядка записи в спецификацию.
        //
        internal class SpecDocumentComparerGost2_106 : SpecDocumentComparer
        {
            public SpecDocumentComparerGost2_106()
                : base()
            {

            }

            /**
             * Сравнение двух документов спецификации, определяющей состав одного изделия.
             * Функция учитывает:
             * - раздел спецификцаии;
             * - правила сортировки записей спецификации для каждого из разделов.
             */
            public bool sortByAlphabet = false;
            public override int Compare(SpecDocument doc1, SpecDocument doc2)
            {
                int cmp = 0;

                SpecDocumentGost2_106 d1 = (SpecDocumentGost2_106)doc1;
                SpecDocumentGost2_106 d2 = (SpecDocumentGost2_106)doc2;

                // Сравнение по разделам спецификации

                cmp = CompareByBomGroup(d1, d2);
                if (cmp != 0)
                    return cmp;

                // Сравнение по номеру позиции, обозначению и наименованию

                cmp = CompareByPositionDesignationAndName(d1, d2, sortByAlphabet);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByRefdes(d1, d2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByCount(d1, d2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByNote(d1, d2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByID(d1, d2);
                return cmp;
            }

            /**
             *  Сравнение на основе разделов спецификации
             */
            protected int CompareByBomGroup(SpecDocumentGost2_106 doc1,
                SpecDocumentGost2_106 doc2)
            {
                int cmp = doc1.BomSection.CompareTo(doc2.BomSection);
                return cmp;
            }

            /**
             *  Сравнение документов из одного раздела по графам спецификации:
             *  - порядковый номер позиции;
             *  - обозначение;
             *  - наименование;
             */
            protected int CompareByPositionDesignationAndName(SpecDocumentGost2_106 doc1,
                SpecDocumentGost2_106 doc2, bool sortByAlphabet)
            {
                Debug.Assert(CompareByBomGroup(doc1, doc2) == 0);

                int cmp = 0;

                // Сравнение по номеру позиции

                cmp = CompareByPositionNumber(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                // Сравнение по обозначению

                // if (doc1.BomGroup.IsContainedInGroup(BomGroupID.Documentation)
                //    && doc2.BomGroup.IsContainedInGroup(BomGroupID.Documentation))
                if ((doc1.BomSection == TFDBomSection.Documentation) && (doc2.BomSection == TFDBomSection.Documentation))
                {
                    // Оба документа входят в раздел Документация
                    cmp = CompareByDesignationInDocumentsBomGroup(doc1, doc2, sortByAlphabet);
                }
                else
                {
                    cmp = CompareByDesignation(doc1, doc2);
                }
                if (cmp != 0)
                    return cmp;

                // Сравнение по наименованию

                cmp = CompareByName(doc1, doc2);

                return cmp;
            }

            private static readonly string[] _documentsOrderByAlphabet =
                {
                    #region Сортировка по алфавиту
            "^(?<prefix>.*)(?<![А-Я])В0$", "^(?<prefix>.*)(?<![А-Я])В0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В1$", "^(?<prefix>.*)(?<![А-Я])В1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В2$", "^(?<prefix>.*)(?<![А-Я])В2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В3$", "^(?<prefix>.*)(?<![А-Я])В3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В4$", "^(?<prefix>.*)(?<![А-Я])В4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В5$", "^(?<prefix>.*)(?<![А-Я])В5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В6$", "^(?<prefix>.*)(?<![А-Я])В6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В7$", "^(?<prefix>.*)(?<![А-Я])В7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВД$", "^(?<prefix>.*)(?<![А-Я])ВД-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВИ$", "^(?<prefix>.*)(?<![А-Я])ВИ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВН$","^(?<prefix>.*)(?<![А-Я])УД$",
            "^(?<prefix>.*)(?<![А-Я])ВО$", "^(?<prefix>.*)(?<![А-Я])ВО-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВП$", "^(?<prefix>.*)(?<![А-Я])ВП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВРК$", "^(?<prefix>.*)(?<![А-Я])ВРК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВРС$", "^(?<prefix>.*)(?<![А-Я])ВРС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВС$", "^(?<prefix>.*)(?<![А-Я])ВС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВЭ$", "^(?<prefix>.*)(?<![А-Я])ВЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г0$", "^(?<prefix>.*)(?<![А-Я])Г0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г1$", "^(?<prefix>.*)(?<![А-Я])Г1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г2$", "^(?<prefix>.*)(?<![А-Я])Г2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г3$", "^(?<prefix>.*)(?<![А-Я])Г3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г4$", "^(?<prefix>.*)(?<![А-Я])Г4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г5$", "^(?<prefix>.*)(?<![А-Я])Г5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г6$", "^(?<prefix>.*)(?<![А-Я])Г6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г7$", "^(?<prefix>.*)(?<![А-Я])Г7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ГЧ$", "^(?<prefix>.*)(?<![А-Я])ГЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*)-ЛУ$",
            "(?<prefix>.*)(?<![А-Я])Д54.[0-9]$",
            "^(?<prefix>.*)(?<![А-Я])ДП$", "^(?<prefix>.*)(?<![А-Я])ДП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е0$", "^(?<prefix>.*)(?<![А-Я])Е0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е1$", "^(?<prefix>.*)(?<![А-Я])Е1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е2$", "^(?<prefix>.*)(?<![А-Я])Е2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е3$", "^(?<prefix>.*)(?<![А-Я])Е3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е4$", "^(?<prefix>.*)(?<![А-Я])Е4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е5$", "^(?<prefix>.*)(?<![А-Я])Е5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е6$", "^(?<prefix>.*)(?<![А-Я])Е6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е7$", "^(?<prefix>.*)(?<![А-Я])Е7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ$", "^(?<prefix>.*)(?<![А-Я])ЗИ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ1$", "^(?<prefix>.*)(?<![А-Я])ЗИ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ2$", "^(?<prefix>.*)(?<![А-Я])ЗИ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИК$", "^(?<prefix>.*)(?<![А-Я])ЗИК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИС$", "^(?<prefix>.*)(?<![А-Я])ЗИС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗК$", "^(?<prefix>.*)(?<![А-Я])ЗК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗС$", "^(?<prefix>.*)(?<![А-Я])ЗС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])И(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])И(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])И1.1$", "^(?<prefix>.*)(?<![А-Я])И1.1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ИМ$", "^(?<prefix>.*)(?<![А-Я])ИМ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ИС(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ИС(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К0$", "^(?<prefix>.*)(?<![А-Я])К0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К1$", "^(?<prefix>.*)(?<![А-Я])К1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К2$", "^(?<prefix>.*)(?<![А-Я])К2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К3$", "^(?<prefix>.*)(?<![А-Я])К3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К4$", "^(?<prefix>.*)(?<![А-Я])К4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К5$", "^(?<prefix>.*)(?<![А-Я])К5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К6$", "^(?<prefix>.*)(?<![А-Я])К6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К7$", "^(?<prefix>.*)(?<![А-Я])К7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])КДС$", "^(?<prefix>.*)(?<![А-Я])КДС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л0$", "^(?<prefix>.*)(?<![А-Я])Л0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л1$", "^(?<prefix>.*)(?<![А-Я])Л1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л2$", "^(?<prefix>.*)(?<![А-Я])Л2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л3$", "^(?<prefix>.*)(?<![А-Я])Л3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л4$", "^(?<prefix>.*)(?<![А-Я])Л4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л5$", "^(?<prefix>.*)(?<![А-Я])Л5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л6$", "^(?<prefix>.*)(?<![А-Я])Л6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л7$", "^(?<prefix>.*)(?<![А-Я])Л7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])М$", //Данные программирования
            "^(?<prefix>.*)(?<![А-Я])МК$", "^(?<prefix>.*)(?<![А-Я])МК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МС$", "^(?<prefix>.*)(?<![А-Я])МС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МЧ$", "^(?<prefix>.*)(?<![А-Я])МЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МЭ$", "^(?<prefix>.*)(?<![А-Я])МЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])НЗЧ$", "^(?<prefix>.*)(?<![А-Я])НЗЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])НМ$", "^(?<prefix>.*)(?<![А-Я])НМ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П0$", "^(?<prefix>.*)(?<![А-Я])П0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П1$", "^(?<prefix>.*)(?<![А-Я])П1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П2$", "^(?<prefix>.*)(?<![А-Я])П2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П3$", "^(?<prefix>.*)(?<![А-Я])П3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П4$", "^(?<prefix>.*)(?<![А-Я])П4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П5$", "^(?<prefix>.*)(?<![А-Я])П5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П6$", "^(?<prefix>.*)(?<![А-Я])П6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П7$", "^(?<prefix>.*)(?<![А-Я])П7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ0$", "^(?<prefix>.*)(?<![А-Я])ПВ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ1$", "^(?<prefix>.*)(?<![А-Я])ПВ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ2$", "^(?<prefix>.*)(?<![А-Я])ПВ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ3$", "^(?<prefix>.*)(?<![А-Я])ПВ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ4$", "^(?<prefix>.*)(?<![А-Я])ПВ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ5$", "^(?<prefix>.*)(?<![А-Я])ПВ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ6$", "^(?<prefix>.*)(?<![А-Я])ПВ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ7$", "^(?<prefix>.*)(?<![А-Я])ПВ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ0$", "^(?<prefix>.*)(?<![А-Я])ПГ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ1$", "^(?<prefix>.*)(?<![А-Я])ПГ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ2$", "^(?<prefix>.*)(?<![А-Я])ПГ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ3$", "^(?<prefix>.*)(?<![А-Я])ПГ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ4$", "^(?<prefix>.*)(?<![А-Я])ПГ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ5$", "^(?<prefix>.*)(?<![А-Я])ПГ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ6$", "^(?<prefix>.*)(?<![А-Я])ПГ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ7$", "^(?<prefix>.*)(?<![А-Я])ПГ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ0$", "^(?<prefix>.*)(?<![А-Я])ПЕ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ1$", "^(?<prefix>.*)(?<![А-Я])ПЕ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ2$", "^(?<prefix>.*)(?<![А-Я])ПЕ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ3$", "^(?<prefix>.*)(?<![А-Я])ПЕ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ4$", "^(?<prefix>.*)(?<![А-Я])ПЕ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ5$", "^(?<prefix>.*)(?<![А-Я])ПЕ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ6$", "^(?<prefix>.*)(?<![А-Я])ПЕ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ7$", "^(?<prefix>.*)(?<![А-Я])ПЕ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЗ$", "^(?<prefix>.*)(?<![А-Я])ПЗ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК0$", "^(?<prefix>.*)(?<![А-Я])ПК0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК1$", "^(?<prefix>.*)(?<![А-Я])ПК1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК2$", "^(?<prefix>.*)(?<![А-Я])ПК2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК3$", "^(?<prefix>.*)(?<![А-Я])ПК3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК4$", "^(?<prefix>.*)(?<![А-Я])ПК4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК5$", "^(?<prefix>.*)(?<![А-Я])ПК5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК6$", "^(?<prefix>.*)(?<![А-Я])ПК6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК7$", "^(?<prefix>.*)(?<![А-Я])ПК7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ0$", "^(?<prefix>.*)(?<![А-Я])ПЛ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ1$", "^(?<prefix>.*)(?<![А-Я])ПЛ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ2$", "^(?<prefix>.*)(?<![А-Я])ПЛ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ3$", "^(?<prefix>.*)(?<![А-Я])ПЛ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ4$", "^(?<prefix>.*)(?<![А-Я])ПЛ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ5$", "^(?<prefix>.*)(?<![А-Я])ПЛ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ6$", "^(?<prefix>.*)(?<![А-Я])ПЛ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ7$", "^(?<prefix>.*)(?<![А-Я])ПЛ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПМ(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ПМ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП0$", "^(?<prefix>.*)(?<![А-Я])ПП0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП1$", "^(?<prefix>.*)(?<![А-Я])ПП1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП2$", "^(?<prefix>.*)(?<![А-Я])ПП2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП3$", "^(?<prefix>.*)(?<![А-Я])ПП3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП4$", "^(?<prefix>.*)(?<![А-Я])ПП4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП5$", "^(?<prefix>.*)(?<![А-Я])ПП5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП6$", "^(?<prefix>.*)(?<![А-Я])ПП6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП7$", "^(?<prefix>.*)(?<![А-Я])ПП7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР0$", "^(?<prefix>.*)(?<![А-Я])ПР0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР1$", "^(?<prefix>.*)(?<![А-Я])ПР1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР2$", "^(?<prefix>.*)(?<![А-Я])ПР2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР3$", "^(?<prefix>.*)(?<![А-Я])ПР3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР4$", "^(?<prefix>.*)(?<![А-Я])ПР4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР5$", "^(?<prefix>.*)(?<![А-Я])ПР5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР6$", "^(?<prefix>.*)(?<![А-Я])ПР6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР7$", "^(?<prefix>.*)(?<![А-Я])ПР7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС$", "^(?<prefix>.*)(?<![А-Я])ПС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС0$", "^(?<prefix>.*)(?<![А-Я])ПС0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС1$", "^(?<prefix>.*)(?<![А-Я])ПС1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС2$", "^(?<prefix>.*)(?<![А-Я])ПС2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС3$", "^(?<prefix>.*)(?<![А-Я])ПС3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС4$", "^(?<prefix>.*)(?<![А-Я])ПС4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС5$", "^(?<prefix>.*)(?<![А-Я])ПС5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС6$", "^(?<prefix>.*)(?<![А-Я])ПС6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС7$", "^(?<prefix>.*)(?<![А-Я])ПС7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПТ$", "^(?<prefix>.*)(?<![А-Я])ПТ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ0$", "^(?<prefix>.*)(?<![А-Я])ПХ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ1$", "^(?<prefix>.*)(?<![А-Я])ПХ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ2$", "^(?<prefix>.*)(?<![А-Я])ПХ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ3$", "^(?<prefix>.*)(?<![А-Я])ПХ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ4$", "^(?<prefix>.*)(?<![А-Я])ПХ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ5$", "^(?<prefix>.*)(?<![А-Я])ПХ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ6$", "^(?<prefix>.*)(?<![А-Я])ПХ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ7$", "^(?<prefix>.*)(?<![А-Я])ПХ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ0$", "^(?<prefix>.*)(?<![А-Я])ПЭ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ1$", "^(?<prefix>.*)(?<![А-Я])ПЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ2$", "^(?<prefix>.*)(?<![А-Я])ПЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ3$", "^(?<prefix>.*)(?<![А-Я])ПЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ4$", "^(?<prefix>.*)(?<![А-Я])ПЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ5$", "^(?<prefix>.*)(?<![А-Я])ПЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ6$", "^(?<prefix>.*)(?<![А-Я])ПЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ7$", "^(?<prefix>.*)(?<![А-Я])ПЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р0$", "^(?<prefix>.*)(?<![А-Я])Р0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р1$", "^(?<prefix>.*)(?<![А-Я])Р1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р2$", "^(?<prefix>.*)(?<![А-Я])Р2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р3$", "^(?<prefix>.*)(?<![А-Я])Р3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р4$", "^(?<prefix>.*)(?<![А-Я])Р4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р5$", "^(?<prefix>.*)(?<![А-Я])Р5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р6$", "^(?<prefix>.*)(?<![А-Я])Р6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р7$", "^(?<prefix>.*)(?<![А-Я])Р7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РК$", "^(?<prefix>.*)(?<![А-Я])РК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РР(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])РР(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РС$", "^(?<prefix>.*)(?<![А-Я])РС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ$", "^(?<prefix>.*)(?<![А-Я])РЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ1$", "^(?<prefix>.*)(?<![А-Я])РЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ2$", "^(?<prefix>.*)(?<![А-Я])РЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ3$", "^(?<prefix>.*)(?<![А-Я])РЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ4$", "^(?<prefix>.*)(?<![А-Я])РЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ5$", "^(?<prefix>.*)(?<![А-Я])РЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ6$", "^(?<prefix>.*)(?<![А-Я])РЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ7$", "^(?<prefix>.*)(?<![А-Я])РЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ8$", "^(?<prefix>.*)(?<![А-Я])РЭ8-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ9$", "^(?<prefix>.*)(?<![А-Я])РЭ9-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ10$", "^(?<prefix>.*)(?<![А-Я])РЭ10-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ11$", "^(?<prefix>.*)(?<![А-Я])РЭ11-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ12$", "^(?<prefix>.*)(?<![А-Я])РЭ12-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ13$", "^(?<prefix>.*)(?<![А-Я])РЭ13-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ14$", "^(?<prefix>.*)(?<![А-Я])РЭ14-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ15$", "^(?<prefix>.*)(?<![А-Я])РЭ15-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ16$", "^(?<prefix>.*)(?<![А-Я])РЭ16-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ17$", "^(?<prefix>.*)(?<![А-Я])РЭ17-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ18$", "^(?<prefix>.*)(?<![А-Я])РЭ18-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ19$", "^(?<prefix>.*)(?<![А-Я])РЭ19-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ20$", "^(?<prefix>.*)(?<![А-Я])РЭ20-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ21$", "^(?<prefix>.*)(?<![А-Я])РЭ21-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ22$", "^(?<prefix>.*)(?<![А-Я])РЭ22-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ23$", "^(?<prefix>.*)(?<![А-Я])РЭ23-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ24$", "^(?<prefix>.*)(?<![А-Я])РЭ24-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ25$", "^(?<prefix>.*)(?<![А-Я])РЭ25-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ26$", "^(?<prefix>.*)(?<![А-Я])РЭ26-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ27$", "^(?<prefix>.*)(?<![А-Я])РЭ27-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ28$", "^(?<prefix>.*)(?<![А-Я])РЭ28-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ29$", "^(?<prefix>.*)(?<![А-Я])РЭ29-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ30$", "^(?<prefix>.*)(?<![А-Я])РЭ30-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ31$", "^(?<prefix>.*)(?<![А-Я])РЭ31-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ32$", "^(?<prefix>.*)(?<![А-Я])РЭ32-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ33$", "^(?<prefix>.*)(?<![А-Я])РЭ33-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ34$", "^(?<prefix>.*)(?<![А-Я])РЭ34-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ35$", "^(?<prefix>.*)(?<![А-Я])РЭ35-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ36$", "^(?<prefix>.*)(?<![А-Я])РЭ36-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ37$", "^(?<prefix>.*)(?<![А-Я])РЭ37-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ38$", "^(?<prefix>.*)(?<![А-Я])РЭ38-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ39$", "^(?<prefix>.*)(?<![А-Я])РЭ39-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ40$", "^(?<prefix>.*)(?<![А-Я])РЭ40-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С0$", "^(?<prefix>.*)(?<![А-Я])С0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С1$", "^(?<prefix>.*)(?<![А-Я])С1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С2$", "^(?<prefix>.*)(?<![А-Я])С2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С3$", "^(?<prefix>.*)(?<![А-Я])С3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С4$", "^(?<prefix>.*)(?<![А-Я])С4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С5$", "^(?<prefix>.*)(?<![А-Я])С5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С6$", "^(?<prefix>.*)(?<![А-Я])С6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С7$", "^(?<prefix>.*)(?<![А-Я])С7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])СБ$", "^(?<prefix>.*)(?<![А-Я])СБ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])СТ$", "^(?<prefix>.*)(?<![А-Я])СТ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Т1М$", "^(?<prefix>.*)(?<![А-Я])Т2М$",
            "^(?<prefix>.*)(?<![А-Я])Т3М$","^(?<prefix>.*)(?<![А-Я])Т4М$",
            "^(?<prefix>.*)(?<![А-Я])ТБ(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ТБ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ0$", "^(?<prefix>.*)(?<![А-Я])ТВ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ1$", "^(?<prefix>.*)(?<![А-Я])ТВ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ2$", "^(?<prefix>.*)(?<![А-Я])ТВ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ3$", "^(?<prefix>.*)(?<![А-Я])ТВ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ4$", "^(?<prefix>.*)(?<![А-Я])ТВ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ5$", "^(?<prefix>.*)(?<![А-Я])ТВ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ6$", "^(?<prefix>.*)(?<![А-Я])ТВ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ7$", "^(?<prefix>.*)(?<![А-Я])ТВ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ0$", "^(?<prefix>.*)(?<![А-Я])ТГ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ1$", "^(?<prefix>.*)(?<![А-Я])ТГ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ2$", "^(?<prefix>.*)(?<![А-Я])ТГ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ3$", "^(?<prefix>.*)(?<![А-Я])ТГ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ4$", "^(?<prefix>.*)(?<![А-Я])ТГ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ5$", "^(?<prefix>.*)(?<![А-Я])ТГ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ6$", "^(?<prefix>.*)(?<![А-Я])ТГ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ7$", "^(?<prefix>.*)(?<![А-Я])ТГ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ0$", "^(?<prefix>.*)(?<![А-Я])ТЕ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ1$", "^(?<prefix>.*)(?<![А-Я])ТЕ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ2$", "^(?<prefix>.*)(?<![А-Я])ТЕ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ3$", "^(?<prefix>.*)(?<![А-Я])ТЕ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ4$", "^(?<prefix>.*)(?<![А-Я])ТЕ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ5$", "^(?<prefix>.*)(?<![А-Я])ТЕ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ6$", "^(?<prefix>.*)(?<![А-Я])ТЕ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ7$", "^(?<prefix>.*)(?<![А-Я])ТЕ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК0$", "^(?<prefix>.*)(?<![А-Я])ТК0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК1$", "^(?<prefix>.*)(?<![А-Я])ТК1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК2$", "^(?<prefix>.*)(?<![А-Я])ТК2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК3$", "^(?<prefix>.*)(?<![А-Я])ТК3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК4$", "^(?<prefix>.*)(?<![А-Я])ТК4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК5$", "^(?<prefix>.*)(?<![А-Я])ТК5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК6$", "^(?<prefix>.*)(?<![А-Я])ТК6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК7$", "^(?<prefix>.*)(?<![А-Я])ТК7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ0$", "^(?<prefix>.*)(?<![А-Я])ТЛ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ1$", "^(?<prefix>.*)(?<![А-Я])ТЛ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ2$", "^(?<prefix>.*)(?<![А-Я])ТЛ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ3$", "^(?<prefix>.*)(?<![А-Я])ТЛ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ4$", "^(?<prefix>.*)(?<![А-Я])ТЛ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ5$", "^(?<prefix>.*)(?<![А-Я])ТЛ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ6$", "^(?<prefix>.*)(?<![А-Я])ТЛ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ7$", "^(?<prefix>.*)(?<![А-Я])ТЛ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТО(?<num>[0-9]*)$","^(?<prefix>.*)(?<![А-Я])ТО(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП$", "^(?<prefix>.*)(?<![А-Я])ТП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП0$", "^(?<prefix>.*)(?<![А-Я])ТП0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП1$", "^(?<prefix>.*)(?<![А-Я])ТП1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП2$", "^(?<prefix>.*)(?<![А-Я])ТП2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП3$", "^(?<prefix>.*)(?<![А-Я])ТП3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП4$", "^(?<prefix>.*)(?<![А-Я])ТП4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП5$", "^(?<prefix>.*)(?<![А-Я])ТП5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП6$", "^(?<prefix>.*)(?<![А-Я])ТП6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП7$", "^(?<prefix>.*)(?<![А-Я])ТП7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР0$", "^(?<prefix>.*)(?<![А-Я])ТР0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР1$", "^(?<prefix>.*)(?<![А-Я])ТР1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР2$", "^(?<prefix>.*)(?<![А-Я])ТР2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР3$", "^(?<prefix>.*)(?<![А-Я])ТР3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР4$", "^(?<prefix>.*)(?<![А-Я])ТР4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР5$", "^(?<prefix>.*)(?<![А-Я])ТР5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР6$", "^(?<prefix>.*)(?<![А-Я])ТР6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР7$", "^(?<prefix>.*)(?<![А-Я])ТР7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС0$", "^(?<prefix>.*)(?<![А-Я])ТС0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС1$", "^(?<prefix>.*)(?<![А-Я])ТС1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС2$", "^(?<prefix>.*)(?<![А-Я])ТС2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС3$", "^(?<prefix>.*)(?<![А-Я])ТС3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС4$", "^(?<prefix>.*)(?<![А-Я])ТС4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС5$", "^(?<prefix>.*)(?<![А-Я])ТС5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС6$", "^(?<prefix>.*)(?<![А-Я])ТС6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС7$", "^(?<prefix>.*)(?<![А-Я])ТС7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТУ(?<num>[0-9]*)$",
            "^(?<prefix>.*)(?<![А-Я])ТУ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ0$", "^(?<prefix>.*)(?<![А-Я])ТХ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ1$", "^(?<prefix>.*)(?<![А-Я])ТХ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ2$", "^(?<prefix>.*)(?<![А-Я])ТХ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ3$", "^(?<prefix>.*)(?<![А-Я])ТХ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ4$", "^(?<prefix>.*)(?<![А-Я])ТХ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ5$", "^(?<prefix>.*)(?<![А-Я])ТХ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ6$", "^(?<prefix>.*)(?<![А-Я])ТХ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ7$", "^(?<prefix>.*)(?<![А-Я])ТХ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЧ$", "^(?<prefix>.*)(?<![А-Я])ТЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ0$", "^(?<prefix>.*)(?<![А-Я])ТЭ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ1$", "^(?<prefix>.*)(?<![А-Я])ТЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ2$", "^(?<prefix>.*)(?<![А-Я])ТЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ3$", "^(?<prefix>.*)(?<![А-Я])ТЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ4$", "^(?<prefix>.*)(?<![А-Я])ТЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ5$", "^(?<prefix>.*)(?<![А-Я])ТЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ6$", "^(?<prefix>.*)(?<![А-Я])ТЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ7$", "^(?<prefix>.*)(?<![А-Я])ТЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УД1$","^(?<prefix>.*)(?<![А-Я])УД2$",
            "^(?<prefix>.*)(?<![А-Я])УД3$","^(?<prefix>.*)(?<![А-Я])УД4$",
            "^(?<prefix>.*)(?<![А-Я])УД5$",
            "^(?<prefix>.*)(?<![А-Я])УК$", "^(?<prefix>.*)(?<![А-Я])УК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УП$", "^(?<prefix>.*)(?<![А-Я])УП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УС$", "^(?<prefix>.*)(?<![А-Я])УС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УЧ$", "^(?<prefix>.*)(?<![А-Я])УЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО$", "^(?<prefix>.*)(?<![А-Я])ФО-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО1$", "^(?<prefix>.*)(?<![А-Я])ФО1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО2$", "^(?<prefix>.*)(?<![А-Я])ФО2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО3$", "^(?<prefix>.*)(?<![А-Я])ФО3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х0$", "^(?<prefix>.*)(?<![А-Я])Х0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х1$", "^(?<prefix>.*)(?<![А-Я])Х1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х2$", "^(?<prefix>.*)(?<![А-Я])Х2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х3$", "^(?<prefix>.*)(?<![А-Я])Х3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х4$", "^(?<prefix>.*)(?<![А-Я])Х4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х5$", "^(?<prefix>.*)(?<![А-Я])Х5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х6$", "^(?<prefix>.*)(?<![А-Я])Х6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х7$", "^(?<prefix>.*)(?<![А-Я])Х7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э0$", "^(?<prefix>.*)(?<![А-Я])Э0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э2$", "^(?<prefix>.*)(?<![А-Я])Э2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э3$", "^(?<prefix>.*)(?<![А-Я])Э3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э4$", "^(?<prefix>.*)(?<![А-Я])Э4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э5$", "^(?<prefix>.*)(?<![А-Я])Э5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э6$", "^(?<prefix>.*)(?<![А-Я])Э6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э7$", "^(?<prefix>.*)(?<![А-Я])Э7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЭП$", "^(?<prefix>.*)(?<![А-Я])ЭП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЭТ$", "^(?<prefix>.*)(?<![А-Я])ЭТ-ЛУ$",
                    #endregion
                };

            // Номенклатура КД по ГОСТ 2.102, таблица 3 (порядок соблюден)
            static readonly string[] _documentsOrder = {
#region Массив регулярных выражений, составленных на основе кодов документов
            "^(?<prefix>.*)(?<![А-Я])СБ$", "^(?<prefix>.*)(?<![А-Я])СБ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВО$", "^(?<prefix>.*)(?<![А-Я])ВО-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЧ$", "^(?<prefix>.*)(?<![А-Я])ТЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ГЧ$", "^(?<prefix>.*)(?<![А-Я])ГЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МЭ$", "^(?<prefix>.*)(?<![А-Я])МЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МЧ$", "^(?<prefix>.*)(?<![А-Я])МЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УЧ$", "^(?<prefix>.*)(?<![А-Я])УЧ-ЛУ$",

            // Схемы по ГОСТ 2.701-84
            "^(?<prefix>.*)(?<![А-Я])Э1$", "^(?<prefix>.*)(?<![А-Я])Э1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ1$", "^(?<prefix>.*)(?<![А-Я])ТЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ1$", "^(?<prefix>.*)(?<![А-Я])ПЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э2$", "^(?<prefix>.*)(?<![А-Я])Э2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ2$", "^(?<prefix>.*)(?<![А-Я])ТЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ2$", "^(?<prefix>.*)(?<![А-Я])ПЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э3$", "^(?<prefix>.*)(?<![А-Я])Э3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ3$", "^(?<prefix>.*)(?<![А-Я])ТЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ3$", "^(?<prefix>.*)(?<![А-Я])ПЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э4$", "^(?<prefix>.*)(?<![А-Я])Э4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ4$", "^(?<prefix>.*)(?<![А-Я])ТЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ4$", "^(?<prefix>.*)(?<![А-Я])ПЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э5$", "^(?<prefix>.*)(?<![А-Я])Э5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ5$", "^(?<prefix>.*)(?<![А-Я])ТЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ5$", "^(?<prefix>.*)(?<![А-Я])ПЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э6$", "^(?<prefix>.*)(?<![А-Я])Э6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ6$", "^(?<prefix>.*)(?<![А-Я])ТЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ6$", "^(?<prefix>.*)(?<![А-Я])ПЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э7$", "^(?<prefix>.*)(?<![А-Я])Э7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ7$", "^(?<prefix>.*)(?<![А-Я])ТЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ7$", "^(?<prefix>.*)(?<![А-Я])ПЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Э0$", "^(?<prefix>.*)(?<![А-Я])Э0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЭ0$", "^(?<prefix>.*)(?<![А-Я])ТЭ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЭ0$", "^(?<prefix>.*)(?<![А-Я])ПЭ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г1$", "^(?<prefix>.*)(?<![А-Я])Г1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ1$", "^(?<prefix>.*)(?<![А-Я])ТГ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ1$", "^(?<prefix>.*)(?<![А-Я])ПГ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г2$", "^(?<prefix>.*)(?<![А-Я])Г2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ2$", "^(?<prefix>.*)(?<![А-Я])ТГ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ2$", "^(?<prefix>.*)(?<![А-Я])ПГ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г3$", "^(?<prefix>.*)(?<![А-Я])Г3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ3$", "^(?<prefix>.*)(?<![А-Я])ТГ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ3$", "^(?<prefix>.*)(?<![А-Я])ПГ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г4$", "^(?<prefix>.*)(?<![А-Я])Г4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ4$", "^(?<prefix>.*)(?<![А-Я])ТГ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ4$", "^(?<prefix>.*)(?<![А-Я])ПГ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г5$", "^(?<prefix>.*)(?<![А-Я])Г5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ5$", "^(?<prefix>.*)(?<![А-Я])ТГ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ5$", "^(?<prefix>.*)(?<![А-Я])ПГ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г6$", "^(?<prefix>.*)(?<![А-Я])Г6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ6$", "^(?<prefix>.*)(?<![А-Я])ТГ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ6$", "^(?<prefix>.*)(?<![А-Я])ПГ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г7$", "^(?<prefix>.*)(?<![А-Я])Г7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ7$", "^(?<prefix>.*)(?<![А-Я])ТГ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ7$", "^(?<prefix>.*)(?<![А-Я])ПГ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Г0$", "^(?<prefix>.*)(?<![А-Я])Г0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТГ0$", "^(?<prefix>.*)(?<![А-Я])ТГ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПГ0$", "^(?<prefix>.*)(?<![А-Я])ПГ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П1$", "^(?<prefix>.*)(?<![А-Я])П1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП1$", "^(?<prefix>.*)(?<![А-Я])ТП1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП1$", "^(?<prefix>.*)(?<![А-Я])ПП1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П2$", "^(?<prefix>.*)(?<![А-Я])П2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП2$", "^(?<prefix>.*)(?<![А-Я])ТП2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП2$", "^(?<prefix>.*)(?<![А-Я])ПП2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П3$", "^(?<prefix>.*)(?<![А-Я])П3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП3$", "^(?<prefix>.*)(?<![А-Я])ТП3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП3$", "^(?<prefix>.*)(?<![А-Я])ПП3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П4$", "^(?<prefix>.*)(?<![А-Я])П4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП4$", "^(?<prefix>.*)(?<![А-Я])ТП4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП4$", "^(?<prefix>.*)(?<![А-Я])ПП4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П5$", "^(?<prefix>.*)(?<![А-Я])П5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП5$", "^(?<prefix>.*)(?<![А-Я])ТП5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП5$", "^(?<prefix>.*)(?<![А-Я])ПП5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П6$", "^(?<prefix>.*)(?<![А-Я])П6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП6$", "^(?<prefix>.*)(?<![А-Я])ТП6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП6$", "^(?<prefix>.*)(?<![А-Я])ПП6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П7$", "^(?<prefix>.*)(?<![А-Я])П7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП7$", "^(?<prefix>.*)(?<![А-Я])ТП7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП7$", "^(?<prefix>.*)(?<![А-Я])ПП7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])П0$", "^(?<prefix>.*)(?<![А-Я])П0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП0$", "^(?<prefix>.*)(?<![А-Я])ТП0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПП0$", "^(?<prefix>.*)(?<![А-Я])ПП0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х1$", "^(?<prefix>.*)(?<![А-Я])Х1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ1$", "^(?<prefix>.*)(?<![А-Я])ТХ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ1$", "^(?<prefix>.*)(?<![А-Я])ПХ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х2$", "^(?<prefix>.*)(?<![А-Я])Х2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ2$", "^(?<prefix>.*)(?<![А-Я])ТХ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ2$", "^(?<prefix>.*)(?<![А-Я])ПХ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х3$", "^(?<prefix>.*)(?<![А-Я])Х3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ3$", "^(?<prefix>.*)(?<![А-Я])ТХ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ3$", "^(?<prefix>.*)(?<![А-Я])ПХ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х4$", "^(?<prefix>.*)(?<![А-Я])Х4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ4$", "^(?<prefix>.*)(?<![А-Я])ТХ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ4$", "^(?<prefix>.*)(?<![А-Я])ПХ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х5$", "^(?<prefix>.*)(?<![А-Я])Х5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ5$", "^(?<prefix>.*)(?<![А-Я])ТХ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ5$", "^(?<prefix>.*)(?<![А-Я])ПХ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х6$", "^(?<prefix>.*)(?<![А-Я])Х6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ6$", "^(?<prefix>.*)(?<![А-Я])ТХ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ6$", "^(?<prefix>.*)(?<![А-Я])ПХ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х7$", "^(?<prefix>.*)(?<![А-Я])Х7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ7$", "^(?<prefix>.*)(?<![А-Я])ТХ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ7$", "^(?<prefix>.*)(?<![А-Я])ПХ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Х0$", "^(?<prefix>.*)(?<![А-Я])Х0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТХ0$", "^(?<prefix>.*)(?<![А-Я])ТХ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПХ0$", "^(?<prefix>.*)(?<![А-Я])ПХ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К1$", "^(?<prefix>.*)(?<![А-Я])К1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК1$", "^(?<prefix>.*)(?<![А-Я])ТК1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК1$", "^(?<prefix>.*)(?<![А-Я])ПК1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К2$", "^(?<prefix>.*)(?<![А-Я])К2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК2$", "^(?<prefix>.*)(?<![А-Я])ТК2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК2$", "^(?<prefix>.*)(?<![А-Я])ПК2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К3$", "^(?<prefix>.*)(?<![А-Я])К3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК3$", "^(?<prefix>.*)(?<![А-Я])ТК3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК3$", "^(?<prefix>.*)(?<![А-Я])ПК3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К4$", "^(?<prefix>.*)(?<![А-Я])К4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК4$", "^(?<prefix>.*)(?<![А-Я])ТК4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК4$", "^(?<prefix>.*)(?<![А-Я])ПК4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К5$", "^(?<prefix>.*)(?<![А-Я])К5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК5$", "^(?<prefix>.*)(?<![А-Я])ТК5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК5$", "^(?<prefix>.*)(?<![А-Я])ПК5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К6$", "^(?<prefix>.*)(?<![А-Я])К6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК6$", "^(?<prefix>.*)(?<![А-Я])ТК6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК6$", "^(?<prefix>.*)(?<![А-Я])ПК6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К7$", "^(?<prefix>.*)(?<![А-Я])К7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК7$", "^(?<prefix>.*)(?<![А-Я])ТК7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК7$", "^(?<prefix>.*)(?<![А-Я])ПК7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])К0$", "^(?<prefix>.*)(?<![А-Я])К0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТК0$", "^(?<prefix>.*)(?<![А-Я])ТК0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПК0$", "^(?<prefix>.*)(?<![А-Я])ПК0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В1$", "^(?<prefix>.*)(?<![А-Я])В1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ1$", "^(?<prefix>.*)(?<![А-Я])ТВ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ1$", "^(?<prefix>.*)(?<![А-Я])ПВ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В2$", "^(?<prefix>.*)(?<![А-Я])В2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ2$", "^(?<prefix>.*)(?<![А-Я])ТВ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ2$", "^(?<prefix>.*)(?<![А-Я])ПВ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В3$", "^(?<prefix>.*)(?<![А-Я])В3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ3$", "^(?<prefix>.*)(?<![А-Я])ТВ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ3$", "^(?<prefix>.*)(?<![А-Я])ПВ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В4$", "^(?<prefix>.*)(?<![А-Я])В4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ4$", "^(?<prefix>.*)(?<![А-Я])ТВ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ4$", "^(?<prefix>.*)(?<![А-Я])ПВ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В5$", "^(?<prefix>.*)(?<![А-Я])В5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ5$", "^(?<prefix>.*)(?<![А-Я])ТВ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ5$", "^(?<prefix>.*)(?<![А-Я])ПВ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В6$", "^(?<prefix>.*)(?<![А-Я])В6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ6$", "^(?<prefix>.*)(?<![А-Я])ТВ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ6$", "^(?<prefix>.*)(?<![А-Я])ПВ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В7$", "^(?<prefix>.*)(?<![А-Я])В7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ7$", "^(?<prefix>.*)(?<![А-Я])ТВ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ7$", "^(?<prefix>.*)(?<![А-Я])ПВ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])В0$", "^(?<prefix>.*)(?<![А-Я])В0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТВ0$", "^(?<prefix>.*)(?<![А-Я])ТВ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПВ0$", "^(?<prefix>.*)(?<![А-Я])ПВ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л1$", "^(?<prefix>.*)(?<![А-Я])Л1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ1$", "^(?<prefix>.*)(?<![А-Я])ТЛ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ1$", "^(?<prefix>.*)(?<![А-Я])ПЛ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л2$", "^(?<prefix>.*)(?<![А-Я])Л2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ2$", "^(?<prefix>.*)(?<![А-Я])ТЛ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ2$", "^(?<prefix>.*)(?<![А-Я])ПЛ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л3$", "^(?<prefix>.*)(?<![А-Я])Л3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ3$", "^(?<prefix>.*)(?<![А-Я])ТЛ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ3$", "^(?<prefix>.*)(?<![А-Я])ПЛ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л4$", "^(?<prefix>.*)(?<![А-Я])Л4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ4$", "^(?<prefix>.*)(?<![А-Я])ТЛ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ4$", "^(?<prefix>.*)(?<![А-Я])ПЛ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л5$", "^(?<prefix>.*)(?<![А-Я])Л5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ5$", "^(?<prefix>.*)(?<![А-Я])ТЛ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ5$", "^(?<prefix>.*)(?<![А-Я])ПЛ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л6$", "^(?<prefix>.*)(?<![А-Я])Л6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ6$", "^(?<prefix>.*)(?<![А-Я])ТЛ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ6$", "^(?<prefix>.*)(?<![А-Я])ПЛ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л7$", "^(?<prefix>.*)(?<![А-Я])Л7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ7$", "^(?<prefix>.*)(?<![А-Я])ТЛ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ7$", "^(?<prefix>.*)(?<![А-Я])ПЛ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Л0$", "^(?<prefix>.*)(?<![А-Я])Л0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЛ0$", "^(?<prefix>.*)(?<![А-Я])ТЛ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЛ0$", "^(?<prefix>.*)(?<![А-Я])ПЛ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р1$", "^(?<prefix>.*)(?<![А-Я])Р1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР1$", "^(?<prefix>.*)(?<![А-Я])ТР1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР1$", "^(?<prefix>.*)(?<![А-Я])ПР1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р2$", "^(?<prefix>.*)(?<![А-Я])Р2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР2$", "^(?<prefix>.*)(?<![А-Я])ТР2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР2$", "^(?<prefix>.*)(?<![А-Я])ПР2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р3$", "^(?<prefix>.*)(?<![А-Я])Р3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР3$", "^(?<prefix>.*)(?<![А-Я])ТР3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР3$", "^(?<prefix>.*)(?<![А-Я])ПР3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р4$", "^(?<prefix>.*)(?<![А-Я])Р4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР4$", "^(?<prefix>.*)(?<![А-Я])ТР4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР4$", "^(?<prefix>.*)(?<![А-Я])ПР4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р5$", "^(?<prefix>.*)(?<![А-Я])Р5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР5$", "^(?<prefix>.*)(?<![А-Я])ТР5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР5$", "^(?<prefix>.*)(?<![А-Я])ПР5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р6$", "^(?<prefix>.*)(?<![А-Я])Р6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР6$", "^(?<prefix>.*)(?<![А-Я])ТР6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР6$", "^(?<prefix>.*)(?<![А-Я])ПР6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р7$", "^(?<prefix>.*)(?<![А-Я])Р7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР7$", "^(?<prefix>.*)(?<![А-Я])ТР7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР7$", "^(?<prefix>.*)(?<![А-Я])ПР7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Р0$", "^(?<prefix>.*)(?<![А-Я])Р0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТР0$", "^(?<prefix>.*)(?<![А-Я])ТР0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПР0$", "^(?<prefix>.*)(?<![А-Я])ПР0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е1$", "^(?<prefix>.*)(?<![А-Я])Е1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ1$", "^(?<prefix>.*)(?<![А-Я])ТЕ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ1$", "^(?<prefix>.*)(?<![А-Я])ПЕ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е2$", "^(?<prefix>.*)(?<![А-Я])Е2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ2$", "^(?<prefix>.*)(?<![А-Я])ТЕ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ2$", "^(?<prefix>.*)(?<![А-Я])ПЕ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е3$", "^(?<prefix>.*)(?<![А-Я])Е3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ3$", "^(?<prefix>.*)(?<![А-Я])ТЕ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ3$", "^(?<prefix>.*)(?<![А-Я])ПЕ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е4$", "^(?<prefix>.*)(?<![А-Я])Е4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ4$", "^(?<prefix>.*)(?<![А-Я])ТЕ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ4$", "^(?<prefix>.*)(?<![А-Я])ПЕ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е5$", "^(?<prefix>.*)(?<![А-Я])Е5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ5$", "^(?<prefix>.*)(?<![А-Я])ТЕ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ5$", "^(?<prefix>.*)(?<![А-Я])ПЕ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е6$", "^(?<prefix>.*)(?<![А-Я])Е6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ6$", "^(?<prefix>.*)(?<![А-Я])ТЕ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ6$", "^(?<prefix>.*)(?<![А-Я])ПЕ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е7$", "^(?<prefix>.*)(?<![А-Я])Е7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ7$", "^(?<prefix>.*)(?<![А-Я])ТЕ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ7$", "^(?<prefix>.*)(?<![А-Я])ПЕ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Е0$", "^(?<prefix>.*)(?<![А-Я])Е0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТЕ0$", "^(?<prefix>.*)(?<![А-Я])ТЕ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЕ0$", "^(?<prefix>.*)(?<![А-Я])ПЕ0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С1$", "^(?<prefix>.*)(?<![А-Я])С1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС1$", "^(?<prefix>.*)(?<![А-Я])ТС1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС1$", "^(?<prefix>.*)(?<![А-Я])ПС1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С2$", "^(?<prefix>.*)(?<![А-Я])С2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС2$", "^(?<prefix>.*)(?<![А-Я])ТС2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС2$", "^(?<prefix>.*)(?<![А-Я])ПС2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С3$", "^(?<prefix>.*)(?<![А-Я])С3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС3$", "^(?<prefix>.*)(?<![А-Я])ТС3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС3$", "^(?<prefix>.*)(?<![А-Я])ПС3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С4$", "^(?<prefix>.*)(?<![А-Я])С4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС4$", "^(?<prefix>.*)(?<![А-Я])ТС4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС4$", "^(?<prefix>.*)(?<![А-Я])ПС4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С5$", "^(?<prefix>.*)(?<![А-Я])С5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС5$", "^(?<prefix>.*)(?<![А-Я])ТС5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС5$", "^(?<prefix>.*)(?<![А-Я])ПС5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С6$", "^(?<prefix>.*)(?<![А-Я])С6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС6$", "^(?<prefix>.*)(?<![А-Я])ТС6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС6$", "^(?<prefix>.*)(?<![А-Я])ПС6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С7$", "^(?<prefix>.*)(?<![А-Я])С7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС7$", "^(?<prefix>.*)(?<![А-Я])ТС7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС7$", "^(?<prefix>.*)(?<![А-Я])ПС7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])С0$", "^(?<prefix>.*)(?<![А-Я])С0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТС0$", "^(?<prefix>.*)(?<![А-Я])ТС0-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПС0$", "^(?<prefix>.*)(?<![А-Я])ПС0-ЛУ$",

            // спецификация,
            "^(?<prefix>.*)(?<![А-Я])ВС$", "^(?<prefix>.*)(?<![А-Я])ВС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВД$", "^(?<prefix>.*)(?<![А-Я])ВД-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВП$", "^(?<prefix>.*)(?<![А-Я])ВП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВИ$", "^(?<prefix>.*)(?<![А-Я])ВИ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ДП$", "^(?<prefix>.*)(?<![А-Я])ДП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПТ$", "^(?<prefix>.*)(?<![А-Я])ПТ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЭП$", "^(?<prefix>.*)(?<![А-Я])ЭП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТП$", "^(?<prefix>.*)(?<![А-Я])ТП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПЗ$", "^(?<prefix>.*)(?<![А-Я])ПЗ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТУ(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ТУ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТО(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ТО(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ПМ(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ПМ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ТБ(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ТБ(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РР(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])РР(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])И(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])И(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])И1.1$", "^(?<prefix>.*)(?<![А-Я])И1.1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*).?[0-9]$", "^(?<prefix>.*)(?<![А-Я])Д(?<num>[0-9]*).?[0-9]-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])М$", //Данные программирования

            // Документы эксплуатационные по ГОСТ 2.601-68
             // Документы эксплуатационные по ГОСТ 2.601-68
            "^(?<prefix>.*)(?<![А-Я])РЭ$", "^(?<prefix>.*)(?<![А-Я])РЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ1$", "^(?<prefix>.*)(?<![А-Я])РЭ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ2$", "^(?<prefix>.*)(?<![А-Я])РЭ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ3$", "^(?<prefix>.*)(?<![А-Я])РЭ3-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ4$", "^(?<prefix>.*)(?<![А-Я])РЭ4-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ5$", "^(?<prefix>.*)(?<![А-Я])РЭ5-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ6$", "^(?<prefix>.*)(?<![А-Я])РЭ6-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ7$", "^(?<prefix>.*)(?<![А-Я])РЭ7-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ8$", "^(?<prefix>.*)(?<![А-Я])РЭ8-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ9$", "^(?<prefix>.*)(?<![А-Я])РЭ9-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ10$", "^(?<prefix>.*)(?<![А-Я])РЭ10-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ11$", "^(?<prefix>.*)(?<![А-Я])РЭ11-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ12$", "^(?<prefix>.*)(?<![А-Я])РЭ12-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ13$", "^(?<prefix>.*)(?<![А-Я])РЭ13-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ14$", "^(?<prefix>.*)(?<![А-Я])РЭ14-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ15$", "^(?<prefix>.*)(?<![А-Я])РЭ15-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ16$", "^(?<prefix>.*)(?<![А-Я])РЭ16-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ17$", "^(?<prefix>.*)(?<![А-Я])РЭ17-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ18$", "^(?<prefix>.*)(?<![А-Я])РЭ18-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ19$", "^(?<prefix>.*)(?<![А-Я])РЭ19-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ20$", "^(?<prefix>.*)(?<![А-Я])РЭ20-ЛУ$",
             "^(?<prefix>.*)(?<![А-Я])РЭ21$", "^(?<prefix>.*)(?<![А-Я])РЭ21-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ22$", "^(?<prefix>.*)(?<![А-Я])РЭ22-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ23$", "^(?<prefix>.*)(?<![А-Я])РЭ23-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ24$", "^(?<prefix>.*)(?<![А-Я])РЭ24-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ25$", "^(?<prefix>.*)(?<![А-Я])РЭ25-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ26$", "^(?<prefix>.*)(?<![А-Я])РЭ26-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ27$", "^(?<prefix>.*)(?<![А-Я])РЭ27-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ28$", "^(?<prefix>.*)(?<![А-Я])РЭ28-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ29$", "^(?<prefix>.*)(?<![А-Я])РЭ29-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ30$", "^(?<prefix>.*)(?<![А-Я])РЭ30-ЛУ$",
             "^(?<prefix>.*)(?<![А-Я])РЭ31$", "^(?<prefix>.*)(?<![А-Я])РЭ31-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ32$", "^(?<prefix>.*)(?<![А-Я])РЭ32-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ33$", "^(?<prefix>.*)(?<![А-Я])РЭ33-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ34$", "^(?<prefix>.*)(?<![А-Я])РЭ34-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ35$", "^(?<prefix>.*)(?<![А-Я])РЭ35-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ36$", "^(?<prefix>.*)(?<![А-Я])РЭ36-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ37$", "^(?<prefix>.*)(?<![А-Я])РЭ37-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ38$", "^(?<prefix>.*)(?<![А-Я])РЭ38-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ39$", "^(?<prefix>.*)(?<![А-Я])РЭ39-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РЭ40$", "^(?<prefix>.*)(?<![А-Я])РЭ40-ЛУ$",

            //---------------------------------------------------------------------------
            "^(?<prefix>.*)(?<![А-Я])ИМ$", "^(?<prefix>.*)(?<![А-Я])ИМ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО$", "^(?<prefix>.*)(?<![А-Я])ФО-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО1$", "^(?<prefix>.*)(?<![А-Я])ФО1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО2$", "^(?<prefix>.*)(?<![А-Я])ФО2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ФО3$", "^(?<prefix>.*)(?<![А-Я])ФО3-ЛУ$",

            "^(?<prefix>.*)(?<![А-Я])ПС$", "^(?<prefix>.*)(?<![А-Я])ПС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЭТ$", "^(?<prefix>.*)(?<![А-Я])ЭТ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])КДС$", "^(?<prefix>.*)(?<![А-Я])КДС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])НЗЧ$", "^(?<prefix>.*)(?<![А-Я])НЗЧ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])НМ$", "^(?<prefix>.*)(?<![А-Я])НМ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ$", "^(?<prefix>.*)(?<![А-Я])ЗИ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ1$", "^(?<prefix>.*)(?<![А-Я])ЗИ1-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИ2$", "^(?<prefix>.*)(?<![А-Я])ЗИ2-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УП$", "^(?<prefix>.*)(?<![А-Я])УП-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ИС(?<num>[0-9]*)$", "^(?<prefix>.*)(?<![А-Я])ИС(?<num>[0-9]*)-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВЭ$", "^(?<prefix>.*)(?<![А-Я])ВЭ-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])СТ$", "^(?<prefix>.*)(?<![А-Я])СТ-ЛУ$",

            // Документы ремонтные по ГОСТ 2.602-68
            "^(?<prefix>.*)(?<![А-Я])РК$", "^(?<prefix>.*)(?<![А-Я])РК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])РС$", "^(?<prefix>.*)(?<![А-Я])РС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УК$", "^(?<prefix>.*)(?<![А-Я])УК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])УС$", "^(?<prefix>.*)(?<![А-Я])УС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗК$", "^(?<prefix>.*)(?<![А-Я])ЗК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗС$", "^(?<prefix>.*)(?<![А-Я])ЗС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МК$", "^(?<prefix>.*)(?<![А-Я])МК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])МС$", "^(?<prefix>.*)(?<![А-Я])МС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИК$", "^(?<prefix>.*)(?<![А-Я])ЗИК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ЗИС$", "^(?<prefix>.*)(?<![А-Я])ЗИС-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВРК$", "^(?<prefix>.*)(?<![А-Я])ВРК-ЛУ$",
            "^(?<prefix>.*)(?<![А-Я])ВРС$", "^(?<prefix>.*)(?<![А-Я])ВРС-ЛУ$",

            "^(?<prefix>.*)(?<![А-Я])ВН$","^(?<prefix>.*)(?<![А-Я])УД$",
            "^(?<prefix>.*)(?<![А-Я])УД1$","^(?<prefix>.*)(?<![А-Я])УД2$",
            "^(?<prefix>.*)(?<![А-Я])УД3$","^(?<prefix>.*)(?<![А-Я])УД4$",
            "^(?<prefix>.*)(?<![А-Я])УД5$",
            "^(?<prefix>.*)(?<![А-Я])УЛ$",
            "^(?<prefix>.*)(?<![А-Я])Т1М$", "^(?<prefix>.*)(?<![А-Я])Т2М$",
            "^(?<prefix>.*)(?<![А-Я])Т3М$","^(?<prefix>.*)(?<![А-Я])Т4М$"
            
#endregion
        };

            /**
             *  Сравнение двух строк на основе массива регулярных выражений.
             *
             *  @param s1
             *      Первая строка
             *  @param s2
             *      Вторая строка
             *  @param regexps
             *      Массив регулярных выражений, определяющих порядок соответствующих им строк.
             *      Строка, соответствующая рег. выражению с _меньшим_ индексом будет _меньше_
             *      строки, соответствующей рег. выражению с _большим_ индексом.
             *      Кроме того, регулярное выражение может содержать именованную группу num.
             *      Если обе строки соответствуют одному и тому же регулярному выражению,
             *      содержащему группу num, результат сравнения определяется числовым значением,
             *      которое соответствует группе num каждой из строки.
             */
            protected int CompareStringsByRegexList(string s1, string s2, string[] regexps)
            {
                if (s1 == s2)
                    return 0;

                int i1 = -1;
                int i2 = -1;
                int cmp = 0;
                bool success1 = false;
                bool success2 = false;

                // Цикл по рег. выражениям
                for (int i = 0; i < regexps.Length; ++i)
                {
                    string regex = regexps[i];

                    Match m1 = Regex.Match(s1, regex);
                    Match m2 = Regex.Match(s2, regex);

                    if (m1.Success)
                    {
                        success1 = true;
                    }

                    if (m2.Success)
                    {
                        success2 = true;
                    }

                    if (m1.Success && m2.Success)
                    {
                        if (m1.Groups["prefix"] != null && m1.Groups["prefix"].Success
                            && m2.Groups["prefix"] != null && m2.Groups["prefix"].Success)
                        {
                            cmp = CompareByStringElements(s1, s2);
                            if (cmp != 0)
                                break;
                        }
                        // Одно и то же рег. выражение кода документа подошло для обоих документов.
                        // Возможно, они различаются номером (группа num). Попытка сравнения
                        // по этим номерам
                        if (m1.Groups["num"] != null && m1.Groups["num"].Success
                            && m2.Groups["num"] != null && m2.Groups["num"].Success)
                        {
                            // Группа num имеется в обоих документах
                            int n1 = Convert.ToInt32(m1.Groups["num"].Value);
                            int n2 = Convert.ToInt32(m2.Groups["num"].Value);
                            if (n1 < n2)
                                cmp = -1;
                            else if (n1 > n2)
                                cmp = 1;
                            if (cmp != 0)
                                break;
                        }
                    }
                    if (m1.Success && i1 < 0)
                        i1 = i;
                    if (m2.Success && i2 < 0)
                        i2 = i;

                    if (i1 >= 0 && i2 >= 0)
                    {
                        // найдены рег. выражения, подходящие под оба документа
                        if (i1 < i2)
                            cmp = -1;
                        else if (i1 > i2)
                            cmp = 1;
                        break;
                    }
                }

                if (!success1 && !notCorrectDenotation.Contains(s1))
                {
                    notCorrectDenotation.Add(s1);
                    MessageBox.Show(
                       s1 +
                       " - Неизвестное обозначение документа! Сортировка документов будет выполнена не по ГОСТу. Обратитесь в бюро 911 для добавления этого обозначения в список разрешенных.",
                       "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                if (!success2 && !notCorrectDenotation.Contains(s2))
                {
                    notCorrectDenotation.Add(s2);
                    MessageBox.Show(
                      s2 +
                      " - Неизвестное обозначение документа! Сортировка документов будет выполнена не по ГОСТу. Обратитесь в бюро 911 для добавления этого обозначения в список разрешенных.",
                      "Внимание!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


                return cmp;
            }


            public List<string> notCorrectDenotation = new List<string>();

            // Порядок сборочного чертежа, его листа утверждения и остальных документоы
            static readonly string[] _assemblyDrawingRegexs = {
            "(?<![А-Я])СБ$",       // сборочный чертеж
            "(?<![А-Я])СБ-ЛУ$",    // лист утверждения сборочного чертежа
            ""              // все остальные документы
        };

            /**
             *  Сравнение документов раздела "Документация"
             */
            protected int CompareByDesignationInDocumentsBomGroup(SpecDocumentGost2_106 doc1,
                SpecDocumentGost2_106 doc2, bool sortByAlphabet)
            {
                //Debug.Assert(doc1.BomGroup.IsContainedInGroup(BomGroupID.Documentation));
                // Debug.Assert(doc2.BomGroup.IsContainedInGroup(BomGroupID.Documentation));
                Debug.Assert(doc1.BomSection == TFDBomSection.Documentation);
                Debug.Assert(doc2.BomSection == TFDBomSection.Documentation);

                // Сборочный чертеж должен быть первым
                int cmp = CompareStringsByRegexList(doc1.Designation, doc2.Designation, _assemblyDrawingRegexs);
                if (cmp != 0)
                    return cmp;

                // Сравнение по кодам документов
                if (sortByAlphabet)
                {
                    cmp = CompareStringsByRegexList(doc1.Designation, doc2.Designation, _documentsOrderByAlphabet);
                }
                else
                    cmp = CompareStringsByRegexList(doc1.Designation, doc2.Designation, _documentsOrder);
                if (cmp != 0)
                    return cmp;

                return cmp;
            }

            /**
             * Сравнение по обозначению
             */
            protected int CompareByDesignation(SpecDocument doc1, SpecDocument doc2)
            {
                return CompareByStringElements(doc1.Designation, doc2.Designation);
            }

            protected int CompareByRefdes(SpecDocumentGost2_106 doc1, SpecDocumentGost2_106 doc2)
            {
                return CompareByStringElements(doc1.RefDes, doc2.RefDes);
            }

            /**
             * Сравнение по наименованию
             */
            protected int CompareByName(SpecDocument doc1, SpecDocument doc2)
            {
                return CompareByStringElements(doc1.Name, doc2.Name, true);
            }
            /**
             *  Сравнение по порядковым номерам (графа "Поз.").
             *  Если это сравнение не определило порядок документов, сравнение
             *  по обозначению и наименованию.
             */
            protected int CompareByPositionNumber(SpecDocumentGost2_106 doc1, SpecDocumentGost2_106 doc2)
            {
                if (doc1.PositionNumber != null && doc2.PositionNumber != null)
                {
                    // Указаны оба порядковых номера
                    if (doc1.PositionNumber < doc2.PositionNumber)
                        return -1;
                    else if (doc1.PositionNumber > doc2.PositionNumber)
                        return 1;
                }

                if (doc1.PositionNumber == null && doc2.PositionNumber != null)
                    // Указано поз. обозначение doc1
                    return -1;

                if (doc1.PositionNumber != null && doc2.PositionNumber == null)
                    // Указано поз. обозначение doc2
                    return 1;

                return 0;
            }

            protected int CompareByCount(SpecDocumentGost2_106 doc1, SpecDocumentGost2_106 doc2)
            {
                if (doc1.Count == null && doc2.Count == null)
                    return 0;
                else if (doc1.Count == null && doc2.Count != null)
                    return -1;
                else if (doc1.Count != null && doc2.Count == null)
                    return 1;
                else if (doc1.Count < doc2.Count)
                    return -1;
                else if (doc1.Count > doc2.Count)
                    return 1;
                else
                    return 0;
            }

            protected int CompareByNote(SpecDocumentGost2_106 doc1, SpecDocumentGost2_106 doc2)
            {
                return String.Compare(doc1.Note, doc2.Note);
            }

            protected int CompareByID(SpecDocumentGost2_106 doc1, SpecDocumentGost2_106 doc2)
            {
                if (doc1.ID < doc2.ID)
                    return -1;
                else if (doc1.ID > doc2.ID)
                    return 1;
                else
                    return 0;
            }
        }

        public class SpecDocumentGost2_106 : SpecDocument
        {
            // Параметры для спецификации

            string[] _zones;        // зоны на чертеже
            string[] _formats;      // форматы листов
            int? _positionNumber;   // номер позиции, null если не указана
            double? _count;         // количество вхождений, null если не указано
            string _primaryUsageSymbol;

            string _countUnit;
            string _extraInfo;


            // Пользовательские (дополнительные) параметры

            string _refDes;
            public TFDDocument _parent;// позиционное обозначение компонента

            /**
             *  Constructor
             */
            public SpecDocumentGost2_106(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr /*, SpecDocument parent*/)
                : base(/*tflexdocs,*/ doc, documentAttr /*, parent*/)
            {
                _parent = doc.rootDocument;
                /*MessageBox.Show(_parent.Naimenovanie + " " + _parent.Denotation);  */
            }



            /**
             *  Чтение параметров документа
             */
            protected override void ReadDocumentParameters(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr/*, SpecDocument parent*/)
            {

                // Чтение основных параметров (базовый класс)
                base.ReadDocumentParameters(/*tflexdocs,*/ doc, documentAttr /*, parent*/);

                // Чтение остальных параметров

                // Зона

                _zones = ReadDocumentZone(/*tflexdocs,*/ doc);

                // Формат

                _formats = ReadDocumentFormats(/*tflexdocs,*/ doc);

                // Порядковый номер позиции
                _positionNumber = doc.Position;//ReadPositionNumber(tflexdocs, doc);


                // Количество
                if (doc.BomSection != TFDBomSection.Documentation)//(!BomGroup.IsContainedInGroup(BomGroupID.Documentation))
                {
                    _count = ReadCount(/*tflexdocs*/ doc);
                }
                else
                    _count = null;

                // Обозначение позиции на схеме
                // _refDes = tflexdocs.GetParameterValue(doc, ParametersGUIDs.RefDes);

                // Символ первичной применяемости
                _primaryUsageSymbol = GetPrimaryUsageSymbol(/*tflexdocs,*/ doc /*, parent*/, documentAttr);

                // Единица измерения количества
                _countUnit = doc.MeasureUnit;//tflexdocs.GetParameterValue(doc, ParametersGUIDs.CountUnit);

                // Дополнительные сведения для спецификации
                //_extraInfo = tflexdocs.GetParameterValue(doc, ParametersGUIDs.ExtraSpecInfo);

            }

            protected virtual string GetPrimaryUsageSymbol(/*TFlexDocs tflexdocs,*/ TFDDocument document, DocumentAttributes documentAttributes
                /*, SpecDocument parent*/)
            {
                // if (parent == null)
                //     return null;


                string primaryUsageSymbol = null;
                /*if (BomGroup.IsContainedInGroup(BomGroupID.Assemblies)
                    || BomGroup.IsContainedInGroup(BomGroupID.Parts)
                    || BomGroup.IsContainedInGroup(BomGroupID.Complexes)
                    || BomGroup.IsContainedInGroup(BomGroupID.Complements))*/
                if ((BomSection == TFDBomSection.Assembly) || (BomSection == TFDBomSection.Details)
                    || (BomSection == TFDBomSection.Komplekts) || (BomSection == TFDBomSection.Complex))
                {
                    //TFDDocument primaryUsageDocument = GetPrimaryUsageDocument(tflexdocs, document);
                    if (document.FirstUse != null)
                    {

                        if (document.rootDocument.Denotation.Trim().ToLower() == document.FirstUse.Trim().ToLower() ||
                            (document.rootDocument.Denotation.Remove(document.rootDocument.Denotation.Length - 3) == document.FirstUse && documentAttributes.specificationForPI &&
                             document.rootDocument.Denotation.Remove(0, document.rootDocument.Denotation.Length - 3) == "ПИ"))
                        // Найден объект с указанием первичной применяемости
                        /*if (tflexdocs.GetStringParameterValue(primaryUsageDocument, ParametersGUIDs.Designation)
                            == parent.Designation
                            && tflexdocs.GetStringParameterValue(primaryUsageDocument, ParametersGUIDs.Name)
                            == parent.Name)*/
                        //if ((DocsClass)primaryUsageDocument.Class == DocsClass.Link
                        //        && Math.Abs(primaryUsageDocument.DocRef) == parent.ID)
                        {
                            primaryUsageSymbol = "0";
                        }
                        else
                        {
                            primaryUsageSymbol = "1";
                        }
                    }
                    else
                    {
                        // Первичная применяемость не указана
                        primaryUsageSymbol = "1";
                    }
                    // if ((DocsCategory)document.Category == DocsCategory.ProgramComponents
                    //  || (DocsCategory)document.Category == DocsCategory.ProgramComplexes)
                    if ((document.Class == TFDClass.ComponentProgram) || (document.Class == TFDClass.ComplexProgram))
                    {
                        primaryUsageSymbol = null;
                    }
                }
                return primaryUsageSymbol;
            }

            /**
             *  Чтение зоны
             */
            string[] ReadDocumentZone(/*TFlexDocs tflexdocs,*/ TFDDocument doc)
            {
                // string sheetZoneParamValue = tflexdocs.GetParameterValue(doc, ParametersGUIDs.SheetZone);
                return GetSignificantSubstrings(doc.Zone);
            }

            /**
             *  Чтение форматов листов документа
             */
            string[] ReadDocumentFormats(/*TFlexDocs tflexdocs,*/ TFDDocument doc)
            {
                // string sheetFormatParamValue = tflexdocs.GetParameterValue(doc, ParametersGUIDs.SheetFormat);
                return GetSignificantSubstrings(doc.Format);
            }


            /**
             *  Чтение количества
             */
            double? ReadCount(/*TFlexDocs tflexdocs,*/ TFDDocument doc)
            {
                string count = doc.Amount.ToString();//tflexdocs.GetParameterValue(doc, ParametersGUIDs.RefCount);

                if (count == null || count.Length == 0 || count == "0")
                    return null;
                else
                {
                    string separator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                    if (separator == ".")
                        count = count.Replace(",", separator);
                    else if (separator == ",")
                        count = count.Replace(".", separator);

                    var nom = count.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                    nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);

                    return Convert.ToDouble(nom);
                }
            }

            public string[] Zones
            {
                get { return _zones; }
            }

            public string[] Formats
            {
                get { return _formats; }
            }

            public int? PositionNumber
            {
                get { return _positionNumber; }
            }

            public double? Count
            {
                get { return _count; }
                set { _count = value; }
            }

            public string RefDes
            {
                get { return _refDes; }
            }

            public string PrimaryUsageSymbol
            {
                get { return _primaryUsageSymbol; }
            }

            public string CountUnit
            {
                get { return _countUnit; }
            }

            public string ExtraInfo
            {
                get { return _extraInfo; }
            }
        }

        /*
         * Компаратор аутентичности двух документов объединения (агрегации)
         * записей спецификации по ГОСТ 2.106.
         * Документы считаются аутентичными, если значения их параметров позволяют
         * записать оба документа в виде одной записи спецификации (например, различие
         * только в параметре Количество).
         */
        public class Guids
        {
            public static class References
            {
                public static readonly Guid Nomenclature = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
                public static readonly Guid Documents = new Guid("ac46ca13-6649-4bbb-87d5-e7d570783f26");
                public static readonly Guid AdmissibleReplacements = new Guid("877f5ca3-6816-46bd-bba1-70cdd1154162");
            }
            public static class AdmissibleReplacementsPars
            {
                public static readonly Guid ObjectForReplaceLink = new Guid("a6cea9be-2e21-4522-b5e0-e99f0edfc6a7");
                public static readonly Guid ReplacingObjectLink = new Guid("182eb351-94a2-4710-8077-f4653469c643");
                public static readonly Guid DevicesResolvingReplaceLink = new Guid("be89cfdf-36e4-4834-b4d3-298e64130eb5");
            }
            public static class NomenclatureTypes
            {
                public static readonly Guid OtherItems = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
                public static readonly Guid Assembly = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
                public static readonly Guid StandartItems = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
                public static readonly Guid Materials = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");
                public static readonly Guid Details = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
                public static readonly Guid Documents = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");
            }
            public static class DocumentsType
            {
                public static readonly Guid OtherItems = new Guid("83e1ef55-0658-4e3e-afeb-d8fceee3c86d");
                public static readonly Guid Assembly = new Guid("dd2cb8e8-48fa-4241-8cab-aac3d83034a7");
                public static readonly Guid StandartItems = new Guid("582dad76-1b07-4c4b-b97d-cc89b0149aa6");
                public static readonly Guid Materials = new Guid("65ab3d63-5301-4a30-b62b-f548a9595b98");
                public static readonly Guid Details = new Guid("7c41c277-41f1-44d9-bf0e-056d930cbb14");
                public static readonly Guid Documents = new Guid("f4a65a70-580a-4404-a7f9-c24b2c7dd2bd");
            }
        }
        internal class SpecDocumentGost2_106AggregateComparer : SpecDocumentBaseEqualityComparer
        {
            public override bool Equals(SpecDocument x, SpecDocument y)
            {
                bool equals = base.Equals(x, y);
                if (!equals)
                    return false;
                else
                {
                    Debug.Assert(x is SpecDocumentGost2_106);
                    Debug.Assert(y is SpecDocumentGost2_106);

                    SpecDocumentGost2_106 docx = x as SpecDocumentGost2_106;
                    SpecDocumentGost2_106 docy = y as SpecDocumentGost2_106;

                    if (docx.BomSection != docy.BomSection)
                        return false;
                    if (docx.PositionNumber != docy.PositionNumber)
                        return false;
                    else if (docx.CountUnit != docy.CountUnit)
                        return false;
                    else if (docx.Note != docy.Note)
                        return false;
                    else if (docx.PrimaryUsageSymbol != docy.PrimaryUsageSymbol)
                        return false;
                    else if (docx.ExtraInfo != docy.ExtraInfo)
                        return false;
                    else if (!StringArraysEquals(docx.Zones, docy.Zones))
                        return false;
                    else if (!StringArraysEquals(docx.Formats, docy.Formats))
                        return false;
                    //  else if (!StringArraysEquals(docx.Litera, docy.Litera))
                    //    return false;
                    else
                        return true;
                }
            }
        }

        /**
         *  Компаратор аутентичности двух документов для спецификации по ГОСТ 2.106.
         *  Документы считаются аутентичными, если у них совпадают параметры, на основе
         *  которых выполняется запись в спецификации.
         */
        internal class SpecDocumentGost2_106EqualityComparer : SpecDocumentBaseEqualityComparer
        {
            public override bool Equals(SpecDocument x, SpecDocument y)
            {
                bool equals = base.Equals(x, y);
                if (!equals)
                    return false;
                else
                {
                    Debug.Assert(x is SpecDocumentGost2_106);
                    Debug.Assert(y is SpecDocumentGost2_106);

                    SpecDocumentGost2_106 docx = x as SpecDocumentGost2_106;
                    SpecDocumentGost2_106 docy = y as SpecDocumentGost2_106;

                    if (docx.BomSection != docy.BomSection)
                        return false;
                    if (docx.PositionNumber != docy.PositionNumber)
                        return false;
                    else if (docx.RefDes != docy.RefDes)
                        return false;
                    else if (docx.Count != docy.Count)
                        return false;
                    else if (docx.CountUnit != docy.CountUnit)
                        return false;
                    else if (docx.Note != docy.Note)
                        return false;
                    // else if (docx.PrimaryUsageSymbol != docy.PrimaryUsageSymbol)
                    //     return false;
                    else if (docx.ExtraInfo != docy.ExtraInfo)
                        return false;
                    else if (!StringArraysEquals(docx.Zones, docy.Zones))
                        return false;
                    else if (!StringArraysEquals(docx.Formats, docy.Formats))
                        return false;
                    // else if (!StringArraysEquals(docx.Litera, docy.Litera))
                    //     return false;
                    else
                        return true;
                }
            }
        }
        internal class SpecDocumentGost2_106EqualityComparer2 : SpecDocumentBaseEqualityComparer
        {
            public override bool Equals(SpecDocument x, SpecDocument y)
            {
                bool equals = base.Equals(x, y);
                if (!equals)
                    return false;
                else
                {
                    Debug.Assert(x is SpecDocumentGost2_106);
                    Debug.Assert(y is SpecDocumentGost2_106);

                    SpecDocumentGost2_106 docx = x as SpecDocumentGost2_106;
                    SpecDocumentGost2_106 docy = y as SpecDocumentGost2_106;

                    if (docx.BomSection != docy.BomSection)
                        return false;
                    if (docx.PositionNumber != docy.PositionNumber)
                        return false;
                    else if (docx.RefDes != docy.RefDes)
                        return false;
                    else if (docx.Count != docy.Count)
                        return false;
                    else if (docx.CountUnit != docy.CountUnit)
                        return false;
                    else if (docx.Note != docy.Note)
                        return false;
                    // else if (docx.PrimaryUsageSymbol != docy.PrimaryUsageSymbol)
                    //     return false;
                    else if (docx.ExtraInfo != docy.ExtraInfo)
                        return false;
                    else if (!StringArraysEquals(docx.Zones, docy.Zones))
                        return false;
                    else if (!StringArraysEquals(docx.Formats, docy.Formats))
                        return false;
                    // else if (!StringArraysEquals(docx.Litera, docy.Litera))
                    //     return false;
                    else
                        return true;
                }
            }
        }

        public class SpecForm1
        {
            TextFormatter _textFormatter;

            // Атрибуты основной надписи

            string _name;
            string _designation;
            string _firstUsedDesignation;
            string _documentGuid;
            string[] _litera;

            DocumentAttributes _attributes;

            //string _hash;

            // Строки таблицы
            List<SpecRowForm1> _rows;

            static readonly int _firstPageRows = 25; // Всего 26 строк, но одна пустая в форматке уже имеется
            static readonly int _otherPageRows = 32;

            public SpecForm1(SpecDocument document, /*SpecDocument firstUsedProduct*/string firstUse, DocumentAttributes attributes
                /*string hash*/)
            {
                Font font = GetSpecTableFont();
                _textFormatter = new TextFormatter(font);
                _rows = new List<SpecRowForm1>();

                if (!(document is SpecDocumentGost2_106))
                    throw new ApplicationException("Документ не является документом класса SpecDocumentGost2_106");
                _name = document.Name;
                _designation = document.Designation;

                if (document is SpecDocumentGost2_113)
                    _litera = null;
                // else if (document is SpecDocumentGost2_106)
                // Заполнение графы Лит. только для одиночных документов
                //     _litera = (document as SpecDocumentGost2_106).Litera;
                else
                    _litera = null;

                _attributes = attributes;
                if (/*firstUsedProduct != null*/firstUse != null)
                {
                    _firstUsedDesignation = firstUse;//firstUsedProduct.Designation;
                }
                else
                {
                    _firstUsedDesignation = "";
                }
                //_documentGuid = document.GUID;

                //  _hash = hash;
            }

            Font GetSpecTableFont()
            {
                return new Font("T-FLEX Type B", 3.5f, GraphicsUnit.Millimeter);
            }

            public void Add(SpecRowForm1 row)
            {
                _rows.Add(row);
            }

            public void Add(SpecRowForm1[] rows)
            {
                foreach (SpecRowForm1 row in rows)
                    Add(row);
            }

            public void AddEmptyRow()
            {
                Add(SpecRowForm1.Empty);
            }

            public void AddEmptyRows(int count)
            {
                for (int i = 0; i < count; ++i)
                {
#if DEBUG_REPORT_LINES
                Add(new SpecRowForm1(new SpecFormCell(i.ToString()), new SpecFormCell(count.ToString())));
#else
                    AddEmptyRow();
#endif
                }
            }

            public void AddEmptyRowsOrFinishCurrentPage(int count)
            {
                for (int i = 0; i < count; ++i)
                {
                    if (AtTopOfPage())
                        break;
                    AddEmptyRow();
                }
            }

            public void NewPage()
            {
                NewPage(false);
            }

            /// Добавление пустых строк до перехода на новый лист.
            /// @param force
            ///     Если true, то в случае, если текущий лист пуст, добавляется пустой лист.
            ///     Если false, то если текущий лист пуст, пустые строки не добавляются.
            public void NewPage(bool force)
            {
                if (force && AtTopOfPage())
#if DEBUG_REPORT_LINES
                Add(new SpecRowForm1(new SpecFormCell("a"), new SpecFormCell("")));
#else
                    AddEmptyRow();
#endif
                while (!AtTopOfPage())
#if DEBUG_REPORT_LINES
                Add(new SpecRowForm1(new SpecFormCell("b" + _rows.Count.ToString()), new SpecFormCell(((_rows.Count  - _firstPageRows) % _otherPageRows).ToString())));
#else
                    AddEmptyRow();
#endif
            }

            public void NewPage(decimal NPages)
            {
                if (NPages == 1)
                    for (int i = 0; i < 29; i++)
                        AddEmptyRow();
                else
                {
                    for (int i = 0; i < 29; i++)
                        AddEmptyRow();
                    for (int i = 0; i < 32 * (NPages - 1); i++)
                        AddEmptyRow();
                }
            }


            public SpecRowForm1[] GetRows()
            {
                return _rows.ToArray();
            }

            int GetPageTopRowIndex(int page)
            {
                if (page < 0)
                    throw new ArgumentException("Номер страницы меньше нуля");

                int index;
                if (page == 0)
                    index = 0;
                else
                    index = _firstPageRows + _otherPageRows * (page - 1);
                return index;
            }

            int GetPageRowCapacity(int page)
            {
                if (page == 0)
                    return _firstPageRows;
                else
                    return _otherPageRows;
            }

            public SpecRowForm1[] GetRows(int page)
            {
                if (page < 0)
                    throw new ArgumentException("Номер страницы меньше нуля");
                if (page >= TotalPageNumber)
                    throw new ArgumentException("Номер страницы превышает количество страниц минус 1");

                int topRowIndex = GetPageTopRowIndex(page);
                int capacity = GetPageRowCapacity(page);

                int count;
                if (topRowIndex + capacity > _rows.Count)
                    count = _rows.Count - topRowIndex;
                else
                    count = capacity;

                return _rows.GetRange(topRowIndex, count).ToArray();
            }

            public bool AtTopOfPage()
            {

                if (_rows.Count < _firstPageRows)
                    return _rows.Count == 0;
                else
                    return (_rows.Count - _firstPageRows) % _otherPageRows == 0;
            }

            public TextFormatter TextFormatter
            {
                get { return _textFormatter; }
            }

            public int TotalPageNumber
            {
                get
                {
                    if (_rows.Count <= _firstPageRows)
                        return 1;
                    else
                        return (_rows.Count - _firstPageRows) / _otherPageRows + 1;
                }
            }

            public int CurrentPageAvialableRows
            {
                get
                {
                    if (_rows.Count < _firstPageRows)
                        return _firstPageRows - _rows.Count;
                    else
                        return _otherPageRows - (_rows.Count - _firstPageRows) % _otherPageRows;
                }
            }

            public string Designation
            {
                get { return _designation; }
            }

            public string Name
            {
                get { return _name; }
            }

            public string FirstUsedProductDesignation
            {
                get { return _firstUsedDesignation; }
            }

            public DocumentAttributes Attributes
            {
                get { return _attributes; }
            }

            public string DocumentGuid
            {
                get { return _documentGuid; }
            }

            public string[] Litera
            {
                get { return _litera; }
            }

            /*  public string Hash
              {
                  get { return _hash; }
              }*/
        }

        /**
         *  Строка таблицы спецификации по форме 1, 1а ГОСТ 2.106
         */
        public class SpecRowForm1
        {
            // Графы
            SpecFormCell _format;
            SpecFormCell _zone;
            SpecFormCell _position;
            SpecFormCell _designation;
            SpecFormCell _name;
            SpecFormCell _count;
            SpecFormCell _note;
            string _noteSuperscriptSymbol;

            public SpecRowForm1(SpecFormCell format, SpecFormCell zone, SpecFormCell position,
                SpecFormCell designation, SpecFormCell name, SpecFormCell count, SpecFormCell note)
                : this(format, zone, position, designation, name, count, note, null)
            {
            }

            public SpecRowForm1(SpecFormCell format, SpecFormCell zone, SpecFormCell position,
                SpecFormCell designation, SpecFormCell name, SpecFormCell count, SpecFormCell note,
                string noteSuperscriptSymbol)
            {
                _format = format;
                _zone = zone;
                _position = position;
                _designation = designation;
                _name = name;
                _count = count;
                _note = note;
                _noteSuperscriptSymbol = noteSuperscriptSymbol;
            }

            public SpecRowForm1(SpecFormCell designation, SpecFormCell name)
                : this(SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                designation, name,
                SpecFormCell.Empty, SpecFormCell.Empty)
            {
            }

            public override string ToString()
            {
                return String.Format(
                    "|{0}|{1}|{2}|{3}|{4}|{5}|{6}|",
                    _format.Text, _zone.Text, _position.Text, _designation.Text,
                    _name.Text, _count.Text,
                    _note.Text + (_noteSuperscriptSymbol == null ? "" : _noteSuperscriptSymbol));
            }

            /**
             * Форматирование записи спецификации, заключающееся в разбивке значений полей на строки
             */
            public static SpecRowForm1[] CreateRecord(string format, string zone, string position,
                string designation, string name, string count, string countUnit, string refdes, string note,
                string noteSuperscriptSymbol,
                TextFormatter textFormatter, string sectionName, bool fullName, string docName, string docDenotation, DocumentAttributes documentAttr)
            {
                List<SpecRowForm1> rows = new List<SpecRowForm1>();

                // Разбивка некоторых полей на несколько строк
                string[] designationLines = textFormatter.Wrap(designation, SpecRowForm1.DesignationColumnWidth);
                //string[] nameLines = textFormatter.Wrap(name, SpecRowForm1.NameColumnWidth);
                string[] nameLines = null;
                string[] documentNameLines = null;

                if (sectionName == TFDBomSection.Documentation || sectionName == TFDBomSection.Komplekts)
                    documentNameLines = textFormatter.WrapDocument(name, SpecRowForm1.NameColumnWidth);

                if ((sectionName == TFDBomSection.OtherProducts) || (sectionName == TFDBomSection.StdProducts) || (sectionName == TFDBomSection.Assembly))
                    nameLines = textFormatter.Wrap(name, SpecRowForm1.NameColumnWidth, sectionName);

                else

                    if ((sectionName == TFDBomSection.Documentation || sectionName == TFDBomSection.Komplekts) && fullName)
                {





                    //nameLines = textFormatter.WrapNotFullName(documentNameLines, SpecRowForm1.NameColumnWidth, docName);


                    nameLines = textFormatter.WrapFullName(designation, documentNameLines,
                                                           SpecRowForm1.NameColumnWidth, documentAttr.FullName,
                                                            docDenotation);

                }
                else
                {


                    if (sectionName == TFDBomSection.Materials)
                        nameLines = textFormatter.WrapMaterial(name, note, SpecRowForm1.NameColumnWidth);
                    else
                        nameLines = textFormatter.Wrap(name, SpecRowForm1.NameColumnWidth);

                    if ((sectionName == TFDBomSection.Documentation) && fullName == false)
                        nameLines = textFormatter.WrapNotFullName(documentNameLines, SpecRowForm1.NameColumnWidth, docName);

                    if (sectionName == TFDBomSection.Komplekts)
                        nameLines = textFormatter.WrapNotFullNameKomplekt(documentNameLines, SpecRowForm1.NameColumnWidth, docName);

                }



                string[] noteLines;
                if (sectionName == TFDBomSection.OtherProducts)
                    noteLines = textFormatter.WrapNote(note, 30f); //Для прочих изделий колонка с примечанием шириной 25
                else
                {
                    if (sectionName == TFDBomSection.Materials)
                        noteLines = textFormatter.WrapNoteMaterial(name, note, SpecRowForm1.NoteColumnWidth);
                    else
                        noteLines = textFormatter.Wrap(note, SpecRowForm1.NoteColumnWidth);

                }

                if ((sectionName == TFDBomSection.Details) || (sectionName == TFDBomSection.StdProducts))
                {
                    nameLines = textFormatter.Wrap(nameLines, name, note, NameColumnWidth);

                    if (note.Contains("Заготовка") || note.Contains("заготовка"))
                    {
                        List<string> emptyList = new List<string>();
                        noteLines = emptyList.ToArray();

                    }

                }

                // Положение записи в графе Кол.
                int countLineNumber = (nameLines.Length == 0) ? 0 : (nameLines.Length - 1);

                // Формирование текста примечания на основе единицы измерения,
                // обозначения на схеме и собственно примечания
                if (!String.IsNullOrEmpty(countUnit) || !String.IsNullOrEmpty(refdes))
                {
                    string unitAndRefdes = "";
                    if (!String.IsNullOrEmpty(countUnit))
                        unitAndRefdes += countUnit;
                    if (!String.IsNullOrEmpty(refdes))
                    {
                        if (!String.IsNullOrEmpty(unitAndRefdes))
                            unitAndRefdes += " ";
                        unitAndRefdes += refdes;
                    }
                    // Задана единица измерения или обозначение на схеме
                    List<string> fullNoteLines = new List<string>();
                    /*    if (noteLines.Length <= countLineNumber)
                        {
                            // Строки примечания помещаются выше строки, где должны разместиться
                            // единица измерения количества и обозначение на схеме
                            fullNoteLines.AddRange(noteLines);
                            while (fullNoteLines.Count < countLineNumber)
                                fullNoteLines.Add("");
                            fullNoteLines.Add(unitAndRefdes);
                        }
                        else
                        {
                            // Строки примечания помещаются под строкой, где должны разместиться
                            // единица измерения количества и обозначение на схеме
                            while (fullNoteLines.Count < countLineNumber)
                                fullNoteLines.Add("");
                            fullNoteLines.Add(unitAndRefdes);
                            fullNoteLines.AddRange(noteLines);
                        }
                        noteLines = fullNoteLines.ToArray();
                    }*/

                    // Строки примечания помещаются под строкой, где должны разместиться
                    // единица измерения количества и обозначение на схеме
                    while (fullNoteLines.Count < countLineNumber)
                        fullNoteLines.Add("");
                    fullNoteLines.Add(unitAndRefdes);
                    fullNoteLines.AddRange(noteLines);
                    noteLines = fullNoteLines.ToArray();
                }
                else
                {

                    string unitAndRefdes = "";
                    if (!String.IsNullOrEmpty(countUnit))
                        unitAndRefdes += countUnit;
                    if (!String.IsNullOrEmpty(refdes))
                    {
                        if (!String.IsNullOrEmpty(unitAndRefdes))
                            unitAndRefdes += " ";
                        unitAndRefdes += refdes;
                    }
                    // Задана единица измерения или обозначение на схеме
                    List<string> fullNoteLines = new List<string>();



                    //***
                    if ((sectionName == TFDBomSection.Documentation) || (sectionName == TFDBomSection.Komplekts))
                    {
                        fullNoteLines.AddRange(noteLines);
                        while (fullNoteLines.Count < countLineNumber)
                            fullNoteLines.Add("");
                        //   fullNoteLines.Add(unitAndRefdes);

                    }
                    else
                    {

                        // Строки примечания помещаются под строкой, где должны разместиться
                        // единица измерения количества и обозначение на схеме
                        while (fullNoteLines.Count < countLineNumber)
                            fullNoteLines.Add("");
                        fullNoteLines.AddRange(noteLines);
                    }
                    noteLines = fullNoteLines.ToArray();
                }
                // Создание строк спецификации
                bool countBool = true;
                bool noteSuperscriptSymbolBool = true;

                while (true)
                {
                    bool empty = true;

                    // Ячейка графы "Формат"
                    SpecFormCell formatCell;
                    if (rows.Count == 0)
                    {
                        formatCell = new SpecFormCell(format, SpecFormCell.AlignFormat.Center);
                        empty = false;
                    }
                    else
                        formatCell = SpecFormCell.Empty;

                    // Ячейка графы "Зона"
                    SpecFormCell zoneCell;
                    if (rows.Count == 0)
                    {
                        zoneCell = new SpecFormCell(zone, SpecFormCell.AlignFormat.Center);
                        empty = false;
                    }
                    else
                        zoneCell = SpecFormCell.Empty;

                    // Ячейка графы "Поз."
                    SpecFormCell positionCell;
                    if (rows.Count == 0)
                    {
                        positionCell = new SpecFormCell(position, SpecFormCell.AlignFormat.Center);
                        empty = false;
                    }
                    else
                        positionCell = SpecFormCell.Empty;

                    // Ячейка графы "Обозначение"
                    SpecFormCell designationCell;
                    if (rows.Count < designationLines.Length)
                    {
                        designationCell = new SpecFormCell(designationLines[rows.Count]);
                        empty = false;
                    }
                    else
                        designationCell = SpecFormCell.Empty;

                    // Ячейка графы "Наименование"
                    SpecFormCell nameCell;
                    if (rows.Count < nameLines.Length)
                    {
                        nameCell = new SpecFormCell(nameLines[rows.Count]);
                        empty = false;
                    }
                    else
                        nameCell = SpecFormCell.Empty;

                    // Ячейка графы "Кол." (выравнивание по низу наименования)
                    SpecFormCell countCell;
                    if (sectionName != TFDBomSection.Komplekts)
                    {
                        if (rows.Count == countLineNumber)
                        {

                            countCell = new SpecFormCell(count, SpecFormCell.AlignFormat.Center);
                            empty = false;
                        }
                        else

                            countCell = SpecFormCell.Empty;
                    }
                    else
                        if (countBool == true)
                    {
                        countCell = new SpecFormCell(count, SpecFormCell.AlignFormat.Center);
                        empty = false;
                        countBool = false;
                    }
                    else

                        countCell = SpecFormCell.Empty;



                    // Ячейка графы "Прим."
                    SpecFormCell noteCell;

                    if (rows.Count < noteLines.Length)
                    {
                        noteCell = new SpecFormCell(noteLines[rows.Count]);
                        empty = false;
                    }
                    else
                        noteCell = SpecFormCell.Empty;


                    string rowNoteSuperscriptSymbol;
                    if (sectionName != TFDBomSection.Komplekts)
                    {
                        if (rows.Count == countLineNumber)
                            rowNoteSuperscriptSymbol = noteSuperscriptSymbol;
                        else
                            rowNoteSuperscriptSymbol = null;
                    }
                    else
                        if (noteSuperscriptSymbolBool)
                    {
                        rowNoteSuperscriptSymbol = noteSuperscriptSymbol;
                        noteSuperscriptSymbolBool = false;
                    }
                    else
                        rowNoteSuperscriptSymbol = null;

                    // Прерывание цикла, если очередная строка оказалась пустой
                    if (empty)
                        break;

                    // Создание строки таблицы
                    SpecRowForm1 row = new SpecRowForm1(formatCell, zoneCell, positionCell,
                                                        designationCell, nameCell, countCell, noteCell,
                                                        rowNoteSuperscriptSymbol);
                    rows.Add(row);
                }

                return rows.ToArray();
            }

            static readonly SpecRowForm1 _empty = new SpecRowForm1(
                SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                SpecFormCell.Empty, SpecFormCell.Empty);

            public static SpecRowForm1 Empty
            {
                get { return _empty; }
            }

            public SpecFormCell Format
            {
                get { return _format; }
            }

            public SpecFormCell Zone
            {
                get { return _zone; }
            }

            public SpecFormCell Position
            {
                get { return _position; }
            }

            public SpecFormCell Designation
            {
                get { return _designation; }
            }

            public SpecFormCell Name
            {
                get { return _name; }
            }

            public SpecFormCell Count
            {
                get { return _count; }
            }

            public SpecFormCell Note
            {
                get { return _note; }
            }

            public string NoteSuperscriptSymbol
            {
                get { return _noteSuperscriptSymbol; }
            }

            public static float FormatColumnWidth
            {
                get { return 6f; }
            }

            public static float ZoneColumnWidth
            {
                get { return 6f; }
            }

            public static float PositionColumnWidth
            {
                get { return 8f; }
            }

            public static float DesignationColumnWidth
            {
                get { return 70f; }
            }

            public static float NameColumnWidth
            {
                get { return 65f; }
            }

            public static float CountColumnWidth
            {
                get { return 10f; }
            }

            public static float NoteColumnWidth
            {
                get
                {
                    // В этой графе используется более узкий шрифт T-FLEX Type A
                    return 29f;
                }
            }
        }
    }


    namespace Globus.TFlexDocs.SpecificationReport.GOST_2_113
    {
        using System;
        using System.Collections;
        using System.Collections.Generic;
        using System.Diagnostics;
        using System.Globalization;
        using System.Text;
        using System.Text.RegularExpressions;

        using Globus.TFlexDocs;
        using Globus.TFlexDocs.SpecificationReport;
        using Globus.TFlexDocs.SpecificationReport.GOST_2_106;
        using System.Threading;

        class ProductComparer : SpecDocumentComparer, IComparer<Product>
        {
            public override int Compare(SpecDocument doc1, SpecDocument doc2)
            {
                return CompareByStringElements(doc1.Designation, doc2.Designation);
            }

            public int Compare(Product a1, Product a2)
            {
                return Compare(a1.SpecDocument, a2.SpecDocument);
            }
        }

        public class SpecDocumentWithTU
        {
            public SpecDocumentGost2_106 SpecDocument;
            public string TU;
            public double nominal;
            public string measure;
            public string firstPartName;
        }

        public class ProductsGroup : Product
        {
            bool _readSubproducts;    // флаг необходимости чтения состава объекта
            List<Product> _products;
            SpecDocumentBaseEqualityComparer _baseEqualityComparer;


            bool _makeRead = true;    // флаг необходимости нахождения исполнений объекта

            public ProductsGroup(/*TFlexDocs tflexdocs*/DocumentAttributes docAttr)
                : this(/*tflexdocs*/  docAttr, true)
            {


            }

            private ProductsGroup(/*TFlexDocs tflexdocs,*/DocumentAttributes docAttr, bool readSubproducts)
                : base(/*tflexdocs*/docAttr.SortByAlphabet)
            {

                _products = new List<Product>();

                _readSubproducts = readSubproducts;

                _baseEqualityComparer = new SpecDocumentBaseEqualityComparer();
            }

            private ProductsGroup(/*TFlexDocs tflexdocs,*/ DocumentAttributes docAttr, bool readSubproducts, bool makeRead)
                : base(/*tflexdocs*/docAttr.SortByAlphabet)
            {

                _products = new List<Product>();
                _readSubproducts = readSubproducts;

                _baseEqualityComparer = new SpecDocumentBaseEqualityComparer();
                _makeRead = makeRead;
            }

            protected override SpecDocumentComparer GetSpecDocumentOrderComparer(bool sByAlphabet)
            {
                var specDoc = new SpecDocumentComparerGost2_113();
                specDoc.sortByAlphabet = sByAlphabet;
                return specDoc;
            }

            protected override SpecDocumentBaseEqualityComparer GetSpecDocumentAuthenticityComparer()
            {
                return new SpecDocumentGost2_113EqualityComparer();
            }

            public override void ReadContent(TFDDocument document, DocumentAttributes documentAttr, ProgressForm progressForm)
            {

                base.ReadContent(document, documentAttr, progressForm);
                _products.Add(this);

                ProductComparer pc = new ProductComparer();
                _products.Sort(pc);

            }

            protected override void ReadChildDocument(TFDDocument doc, TFDDocument parent, DocumentAttributes documentAttr)
            {
                // Чтение исполнения, входящего непосредственно в состав основного исполнения
                /*   if (doc.Parent == SpecDocument.ID
                       && (DocsCategory)doc.Category == DocsCategory.Makes
                       && _readSubproducts)
                   {
                       ReadMakeDocument(doc);
                   }
                   else*/
                //                     base.ReadChildDocument(doc, parent);

                SqlConnection connection = ApiDocs.GetConnection(true);
                //********
                ProgressForm progressForm = new ProgressForm();
                if (_makeRead)
                {
                    // Поиск исполнений объекта
                    string findMakesQuery = String.Format(@"SELECT DISTINCT MasterID 
                                                            FROM NomenclatureBaseVersion nbv INNER JOIN Nomenclature n ON (n.s_ObjectID = nbv.MasterID)
                                                            WHERE SlaveID = {0}", parent.ObjectID);
                    SqlCommand findMakesCommand = new SqlCommand(findMakesQuery, connection);
                    findMakesCommand.CommandTimeout = 0;
                    SqlDataReader findMakesReader = findMakesCommand.ExecuteReader();


                    if (findMakesReader.HasRows)
                    {
                        int count = 0;
                        // Чтение объектов состава найденных исполнений (первый уровень)
                        while (findMakesReader.Read())
                        {
                            count++;
                            int makeID = findMakesReader.GetInt32(0);
                            TFDDocument makeDoc = new TFDDocument(makeID);
                            ReadMakeDocument(makeDoc, false, documentAttr, progressForm);
                        }
                        findMakesReader.Close();
                    }
                    _makeRead = false;
                    base.ReadChildDocument(doc, parent, documentAttr);
                }
                else
                    base.ReadChildDocument(doc, parent, documentAttr);
            }

            void ReadMakeDocument(TFDDocument makeDoc, bool makeRead, DocumentAttributes documentAttr, ProgressForm progressForm)
            {

                Product make = new ProductsGroup(/*TFlexDocs,*/ documentAttr, false, false);
                make.ReadContent(makeDoc, documentAttr, progressForm);
                _products.Add(make);
            }

            protected override SpecDocumentGost2_106 CreateSpecDocument(TFDDocument doc, DocumentAttributes documentAttr)
            {
                //MessageBox.Show("CreateSpecDocument ГОСТ 2_113");
                //MessageBox.Show("SpecDocument " + SpecDocument.Name);
                SpecDocumentGost2_113 child = new SpecDocumentGost2_113(/*TFlexDocs,*/ doc, documentAttr /*, SpecDocument*/);
                //MessageBox.Show("child создан");
                return child;
            }

            protected override string GetSpecDocumentPositionField(SpecDocumentGost2_106 document,
                List<SpecDocumentGost2_106> addedDocuments, bool zeroToDash)
            {
                Debug.Assert(document is SpecDocumentGost2_113);

                SpecDocumentGost2_113 d = document as SpecDocumentGost2_113;
                if (d.IsBaseMake)
                {
                    // Прочерк для базового исполнения
                    return "-";
                }
                else
                    return base.GetSpecDocumentPositionField(document, addedDocuments, zeroToDash);
            }


            //*******************
            //Cортировка по единицам измерения	
            public List<SpecDocumentWithTU> SortedByMeasureUnit(List<SpecDocumentWithTU> unSortedDocList, string measureUnit)
            {
                List<SpecDocumentWithTU> SortedDocList = new List<SpecDocumentWithTU>();
                foreach (SpecDocumentWithTU doc in unSortedDocList)
                    if (doc.measure == measureUnit)
                        SortedDocList.Add(doc);
                return SortedByNominal(SortedDocList);
            }

            //Формирование полей размерности, номинала и первой части наименования(до номинала)	
            public List<SpecDocumentWithTU> ModifyDocList(List<SpecDocumentWithTU> DocList)
            {
                string measurePattern = @"(мкФ|пФ|Ом|кОм|МОм)";
                string nominalPattern = @"(?<=(-|\s))((\d+)[.,]?\d*)(?=\s?(мкФ|пФ|Ом|кОм|МОм|/))";
                string firstPartNamePattern = @".*(?=(-|\s)\d+[.,/]?\d*\s?(мкФ|пФ|Ом|кОм|МОм))";
                Regex regexMeasure = new Regex(measurePattern);
                Match matchMeasure;
                Regex regexNominal = new Regex(nominalPattern);
                Match matchNominal;
                Regex regexFirstPartName = new Regex(firstPartNamePattern);
                Match matchFirstPartName;
                string correctNominal;
                string name;

                TextNormalizer norm = new TextNormalizer();

                for (int i = 0; i < DocList.Count; i++)
                {
                    name = DocList[i].SpecDocument.Name;
                    matchMeasure = regexMeasure.Match(name);
                    //Запись в поле единица измерения
                    if (matchMeasure.Value == "")
                    {
                        DocList[i].measure = "Ом";
                    }
                    else
                        DocList[i].measure = matchMeasure.Value;

                    matchNominal = regexNominal.Match(name);
                    correctNominal = matchNominal.Value;
                    correctNominal.Replace(",", ".");
                    if (correctNominal == "")
                        correctNominal = "0";
                    //Запись в поле номинал
                    // MessageBox.Show(i + "   pos " + DocList[i].SpecDocument.PositionNumber+"\nname "+name + "  \n" + correctNominal);
                    try
                    {
                        var nom = correctNominal.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        DocList[i].nominal = Convert.ToDouble(nom);
                    }
                    catch
                    {
                        MessageBox.Show(
                            "Объект:\n" + DocList[i].SpecDocument.Name +
                            "\nзаписан неправильно! Спецификация некорректна!\nОбратитесь к администратору справочника прочие изделия!" + correctNominal,
                            "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        DocList[i].nominal = 0;
                    }

                    matchFirstPartName = regexFirstPartName.Match(name);
                    name = matchFirstPartName.Value;

                    name = norm.GetNormalForm(name);
                    //Запись в поле первая часть наименования(до номинала) 
                    DocList[i].firstPartName = name;

                }
                //Сортровка списка по первой части наименования(до номинала)
                DocList = SortedByFirstPartName(DocList);

                return DocList;
            }


            //Формирование полей размерности, номинала и первой части наименования(до номинала)	
            public SpecDocumentWithTU ModifyDoc(SpecDocumentWithTU doc)
            {
                string measurePattern = @"(мкФ|пФ|Ом|кОм|МОм)";
                string nominalPattern = @"(?<=(-|\s))((\d+)[.,]?\d*)(?=\s?(мкФ|пФ|Ом|кОм|МОм|/))";
                string firstPartNamePattern = @".*(?=(-|\s)\d+[.,/]?\d*\s?(мкФ|пФ|Ом|кОм|МОм))";
                Regex regexMeasure = new Regex(measurePattern);
                Match matchMeasure;
                Regex regexNominal = new Regex(nominalPattern);
                Match matchNominal;
                Regex regexFirstPartName = new Regex(firstPartNamePattern);
                Match matchFirstPartName;
                string correctNominal;
                string name;

                TextNormalizer norm = new TextNormalizer();

                name = doc.SpecDocument.Name;
                matchMeasure = regexMeasure.Match(name);
                //Запись в поле единица измерения
                doc.measure = matchMeasure.Value;

                matchNominal = regexNominal.Match(name);
                correctNominal = matchNominal.Value;
                correctNominal.Replace(",", ".");
                //Запись в поле номинал
                // MessageBox.Show(i + "   pos " + DocList[i].SpecDocument.PositionNumber+"\nname "+name + "  \n" + correctNominal);
                try
                {

                    var nom = correctNominal.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                    nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                    doc.nominal = Convert.ToDouble(nom);
                }
                catch
                {
                    MessageBox.Show(
                       "Объект:\n" + doc.SpecDocument.Name +
                            "\nзаписан неправильно! Спецификация некорректна!\nОбратитесь к администратору справочника прочие изделия!",
                            "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    doc.nominal = 0;
                }

                matchFirstPartName = regexFirstPartName.Match(name);
                name = matchFirstPartName.Value;

                name = norm.GetNormalForm(name);
                //Запись в поле первая часть наименования(до номинала) 
                doc.firstPartName = name;

                return doc;
            }

            //Cортировка по номиналам размерности	
            public List<SpecDocumentWithTU> SortedByNominal(List<SpecDocumentWithTU> unSortedDocList)
            {
                for (int i = 0; i < unSortedDocList.Count; i++)
                {
                    for (int j = 0; j < unSortedDocList.Count - 1 - i; j++)
                    {
                        if (unSortedDocList[j].nominal > unSortedDocList[j + 1].nominal)
                        {
                            SpecDocumentWithTU specDocument = unSortedDocList[j];
                            unSortedDocList[j] = unSortedDocList[j + 1];
                            unSortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                return unSortedDocList;
            }

            //Cортировка по наименованию в порядке убывания	
            public List<SpecDocumentWithTU> SortedByNameDecrease(List<SpecDocumentWithTU> nullPositionDocList)
            {
                for (int i = 0; i < nullPositionDocList.Count; i++)
                {
                    for (int j = 0; j < nullPositionDocList.Count - 1 - i; j++)
                    {
                        if (nullPositionDocList[j + 1].SpecDocument.Name.CompareTo(nullPositionDocList[j].SpecDocument.Name) == 1)
                        {
                            SpecDocumentWithTU specDocument = nullPositionDocList[j];
                            nullPositionDocList[j] = nullPositionDocList[j + 1];
                            nullPositionDocList[j + 1] = specDocument;
                        }
                    }
                }
                return nullPositionDocList;
            }

            //Cортировка по наименованию	
            public List<SpecDocumentWithTU> SortedByName(List<SpecDocumentWithTU> nullPositionDocList)
            {
                for (int i = 0; i < nullPositionDocList.Count; i++)
                {
                    for (int j = 0; j < nullPositionDocList.Count - 1 - i; j++)
                    {
                        if (nullPositionDocList[j].SpecDocument.Name.CompareTo(nullPositionDocList[j + 1].SpecDocument.Name) == 1)
                        {
                            SpecDocumentWithTU specDocument = nullPositionDocList[j];
                            nullPositionDocList[j] = nullPositionDocList[j + 1];
                            nullPositionDocList[j + 1] = specDocument;
                        }
                    }
                }
                return nullPositionDocList;
            }
            //Cортировка по Обозначению	
            public List<SpecDocumentGost2_106> SortedByDenotation(List<SpecDocumentGost2_106> unsortedDocList)
            {
                for (int i = 0; i < unsortedDocList.Count; i++)
                {
                    for (int j = 0; j < unsortedDocList.Count - 1 - i; j++)
                    {
                        if (unsortedDocList[j].Designation.CompareTo(unsortedDocList[j + 1].Designation) == 1)
                        {
                            SpecDocumentGost2_106 specDocument = unsortedDocList[j];
                            unsortedDocList[j] = unsortedDocList[j + 1];
                            unsortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                return unsortedDocList;
            }

            //Cортировка по Обозначению	(обозначения БФМИ.460010.005 и БФМИ.460086.006 ставятся на первое место (замечание А. И. Баскаковой))
            public List<string> SortedByDenotation(List<string> unsortedDocList)
            {
                for (int i = 0; i < unsortedDocList.Count; i++)
                {
                    for (int j = 0; j < unsortedDocList.Count - 1 - i; j++)
                    {
                        if (unsortedDocList[j].CompareTo(unsortedDocList[j + 1]) == 1)
                        {
                            string specDocument = unsortedDocList[j];
                            unsortedDocList[j] = unsortedDocList[j + 1];
                            unsortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                List<string> sortedDocList = new List<string>();
                foreach (string str in unsortedDocList)
                {
                    if (str == "БФМИ.460010.005")
                        sortedDocList.Add(str);

                    if (str == "БФМИ.460086.001" || str == "БФМИ.460086.002" || str == "БФМИ.460086.003" || str == "БФМИ.460086.004" || str == "БФМИ.460086.005" || str == "БФМИ.460086.006")
                    {
                        sortedDocList.Add(str);
                        // break;
                    }


                }

                foreach (string str in unsortedDocList)
                {
                    if (str != "БФМИ.460010.005" && str != "БФМИ.460086.006" && str != "БФМИ.460086.001" && str != "БФМИ.460086.002" && str != "БФМИ.460086.003" && str != "БФМИ.460086.004" && str != "БФМИ.460086.005")
                        sortedDocList.Add(str);
                }
                return sortedDocList;
            }


            //Cортировка по Обозначению	
            public List<SpecDocument> SortedByDenotation(List<SpecDocument> unsortedDocList)
            {
                for (int i = 0; i < unsortedDocList.Count; i++)
                {
                    for (int j = 0; j < unsortedDocList.Count - 1 - i; j++)
                    {
                        if (unsortedDocList[j].Designation.CompareTo(unsortedDocList[j + 1].Designation) == 1)
                        {
                            SpecDocument specDocument = unsortedDocList[j];
                            unsortedDocList[j] = unsortedDocList[j + 1];
                            unsortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                return unsortedDocList;
            }

            //Cортировка по Обозначению	
            public List<SpecDocumentWithTU> SortedByDenotation(List<SpecDocumentWithTU> unsortedDocList)
            {
                for (int i = 0; i < unsortedDocList.Count; i++)
                {
                    for (int j = 0; j < unsortedDocList.Count - 1 - i; j++)
                    {
                        if (unsortedDocList[j].SpecDocument.Designation.CompareTo(unsortedDocList[j + 1].SpecDocument.Designation) == 1)
                        {
                            SpecDocumentWithTU specDocument = unsortedDocList[j];
                            unsortedDocList[j] = unsortedDocList[j + 1];
                            unsortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                return unsortedDocList;
            }

            //Cортировка по первой части наименования (до номинала)
            public List<SpecDocumentWithTU> SortedByFirstPartName(List<SpecDocumentWithTU> unsortedDocList)
            {
                for (int i = 0; i < unsortedDocList.Count; i++)
                {
                    for (int j = 0; j < unsortedDocList.Count - 1 - i; j++)
                    {
                        if (unsortedDocList[j].firstPartName.CompareTo(unsortedDocList[j + 1].firstPartName) == 1)
                        {
                            SpecDocumentWithTU specDocument = unsortedDocList[j];
                            unsortedDocList[j] = unsortedDocList[j + 1];
                            unsortedDocList[j + 1] = specDocument;
                        }
                    }
                }
                return unsortedDocList;
            }

            //Cортировка по номеру позиции
            public List<SpecDocumentWithTU> SortedByPosition(List<SpecDocumentWithTU> withPositionDocList)
            {
                for (int i = 0; i < withPositionDocList.Count; i++)
                {
                    for (int j = 0; j < withPositionDocList.Count - 1 - i; j++)
                    {
                        if (withPositionDocList[j].SpecDocument.PositionNumber > withPositionDocList[j + 1].SpecDocument.PositionNumber)
                        {
                            SpecDocumentWithTU specDocument = withPositionDocList[j];
                            withPositionDocList[j] = withPositionDocList[j + 1];
                            withPositionDocList[j + 1] = specDocument;
                        }
                    }
                }
                return withPositionDocList;
            }

            //Cортировка по номеру позиции
            public List<SpecDocumentGost2_106> SortedByPosition(List<SpecDocumentGost2_106> withPositionDocList)
            {
                for (int i = 0; i < withPositionDocList.Count; i++)
                {
                    for (int j = 0; j < withPositionDocList.Count - 1 - i; j++)
                    {
                        if (withPositionDocList[j].PositionNumber > withPositionDocList[j + 1].PositionNumber)
                        {
                            SpecDocumentGost2_106 specDocument = withPositionDocList[j];
                            withPositionDocList[j] = withPositionDocList[j + 1];
                            withPositionDocList[j + 1] = specDocument;
                        }
                    }
                }
                return withPositionDocList;
            }

            //Метод сортировки подборочных резисторов и конденсаторов по номеру ТУ и наименованию внутри 
            //каждой подгруппы по ТУ
            public List<SpecDocumentWithTU> SortedByTU(List<SpecDocumentWithTU> nullPosCondRezSpecList)
            {
                //Временный список для формирования подгруппы
                List<SpecDocumentWithTU> tempList;
                //Временный список для подгруппы отсортированных по первой части наименования
                List<SpecDocumentWithTU> tempSortedByName;
                //Итоговый отсортированный список
                List<SpecDocumentWithTU> sortedList = new List<SpecDocumentWithTU>();

                //Временный список для формирования подгруппы для размерности Ом
                List<SpecDocumentWithTU> tempListOm;
                //Временный список для формирования подгруппы для размерности кОм
                List<SpecDocumentWithTU> tempListkOm;
                //Временный список для формирования подгруппы для размерности МОм
                List<SpecDocumentWithTU> tempListMOm;
                //Временный список для формирования подгруппы для размерности мкФ
                List<SpecDocumentWithTU> tempListmkF;
                //Временный список для формирования подгруппы для размерности пФ
                List<SpecDocumentWithTU> tempListpF;

                //Cписок удаленных элементов из временного списка
                List<SpecDocumentWithTU> removeTempList;

                string firstPartName;

                while (nullPosCondRezSpecList.Count > 0)
                {
                    tempList = new List<SpecDocumentWithTU>();
                    //Выделяем ТУ первого элемента в списке
                    string tu = nullPosCondRezSpecList[0].TU;

                    //Поиск во входном списке элементов с таким же значением ТУ
                    //и запись во временный список
                    foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                        if (doc.TU == tu)
                            tempList.Add(doc);

                    //Удаление из входного списка элементов записанных во временный список
                    foreach (SpecDocumentWithTU doc in tempList)
                        nullPosCondRezSpecList.Remove(doc);

                    //Формирование поля номинал, поля размерность и поля первая часть наименования (до номинала)
                    tempList = ModifyDocList(tempList);

                    while (tempList.Count > 0)
                    {
                        tempSortedByName = new List<SpecDocumentWithTU>();
                        tempListOm = new List<SpecDocumentWithTU>();
                        tempListkOm = new List<SpecDocumentWithTU>();
                        tempListMOm = new List<SpecDocumentWithTU>();
                        tempListpF = new List<SpecDocumentWithTU>();
                        tempListmkF = new List<SpecDocumentWithTU>();

                        removeTempList = new List<SpecDocumentWithTU>();

                        tempSortedByName.Add(tempList[0]);
                        firstPartName = tempList[0].firstPartName;
                        removeTempList.Add(tempList[0]);

                        //Поиск во временном списке элементов с одинаковыми первыми частями
                        //и формирование списка удаленных элементов
                        for (int i = 1; i < tempList.Count; i++)
                        {
                            if (tempList[i].firstPartName == firstPartName)
                            {
                                tempSortedByName.Add(tempList[i]);
                                removeTempList.Add(tempList[i]);
                            }
                        }
                        //Очистка временного списка от удаленных элементов
                        foreach (SpecDocumentWithTU doc in removeTempList)
                            tempList.Remove(doc);

                        //Сортировка по размерности и номиналу временных списков	 	  
                        tempListOm = SortedByMeasureUnit(tempSortedByName, "Ом");
                        tempListkOm = SortedByMeasureUnit(tempSortedByName, "кОм");
                        tempListMOm = SortedByMeasureUnit(tempSortedByName, "МОм");
                        tempListpF = SortedByMeasureUnit(tempSortedByName, "пФ");
                        tempListmkF = SortedByMeasureUnit(tempSortedByName, "мкФ");

                        //Передача данных из временных списков в итоговый 
                        sortedList.AddRange(tempListOm);
                        sortedList.AddRange(tempListkOm);
                        sortedList.AddRange(tempListMOm);
                        sortedList.AddRange(tempListpF);
                        sortedList.AddRange(tempListmkF);
                    }
                }
                return sortedList;
            }

            //Объединение списка 2 и списка 3 (добавление элементов списка 3 в список 2 к объектам с соотвтствующем значением ТУ)
            public List<SpecDocumentWithTU> UnionGroupList(List<SpecDocumentWithTU> withPositionDocSpecList, List<SpecDocumentWithTU> nullPosCondRezSpecList)
            {
                //Итоговый объединенный список
                List<SpecDocumentWithTU> unionList = new List<SpecDocumentWithTU>();

                //Итоговый объединенный список2 для вставки резисторов и конденсаторов с нулевыми позициями
                List<SpecDocumentWithTU> unionList2 = new List<SpecDocumentWithTU>();

                //Сортировка списка 2 по номеру позиции
                withPositionDocSpecList = SortedByPosition(withPositionDocSpecList);
                //Список записанных элементов списка 3
                List<SpecDocumentWithTU> writedNullPosList = new List<SpecDocumentWithTU>();

                //Значение параметра ТУ текущего объекта
                string tu = String.Empty;

                //Значение параметра марка текущего объекта
                string marka = String.Empty;

                //Значение параметра ТУ следующего объекта
                string nextTu = String.Empty;

                //Значение параметра марка следующего объекта
                string nextMarka = String.Empty;

                if (withPositionDocSpecList.Count >= 2)
                {
                    for (int i = 0; i < withPositionDocSpecList.Count - 1; i++)
                    {
                        tu = withPositionDocSpecList[i].TU;
                        marka = withPositionDocSpecList[i].firstPartName;
                        nextTu = withPositionDocSpecList[i + 1].TU;
                        nextMarka = withPositionDocSpecList[i + 1].firstPartName;

                        //Если параметр ТУ пуст, то записываем в итоговый список
                        if (tu == String.Empty)
                            unionList.Add(withPositionDocSpecList[i]);
                        //Если параметр ТУ не пустой, то объект записывается, а его параметр ТУ
                        //сравнивается с параметром ТУ следующего объекта. 
                        //Если они не совпадают, то обращаемся к списку 3, находим в нем все
                        //объекты с параметром ТУ равным текущему и записываем в итоговый список.
                        else
                        {
                            unionList.Add(withPositionDocSpecList[i]);

                            if (tu != nextTu)
                            {
                                foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                {
                                    if (doc.TU == tu)
                                    {
                                        for (int j = (i + 1); j < withPositionDocSpecList.Count; j++)
                                        {
                                            if (tu == withPositionDocSpecList[j].TU)
                                            {
                                                unionList.Add(withPositionDocSpecList[j]);
                                                withPositionDocSpecList.Remove(withPositionDocSpecList[j]);
                                            }
                                        }
                                        unionList.Add(doc);
                                        writedNullPosList.Add(doc);
                                    }
                                }
                            }
                            else
                            {
                                if (marka != nextMarka)
                                    foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                    {
                                        if (doc.firstPartName == marka)
                                        {
                                            for (int j = (i + 1); j < withPositionDocSpecList.Count; j++)
                                            {
                                                if (marka == withPositionDocSpecList[j].firstPartName)
                                                {
                                                    unionList.Add(withPositionDocSpecList[j]);
                                                    withPositionDocSpecList.Remove(withPositionDocSpecList[j]);
                                                }
                                            }
                                            unionList.Add(doc);
                                            writedNullPosList.Add(doc);
                                        }
                                    }
                            }



                        }
                        //Очистка списка 3 от уже задействованных элементов
                        foreach (SpecDocumentWithTU doc in writedNullPosList)
                        {
                            nullPosCondRezSpecList.Remove(doc);
                        }
                    }


                    writedNullPosList.Clear();

                    //Записываем последний элемент списка 2
                    unionList.Add(withPositionDocSpecList[withPositionDocSpecList.Count - 1]);

                    //Если список 3 не пустой, то в нем находим все элементы с ТУ, равным параметру ТУ
                    //последнего элемента списка 2 и записываем в итоговый список
                    if (nullPosCondRezSpecList.Count > 0)
                    {
                        foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                        {
                            if (withPositionDocSpecList[withPositionDocSpecList.Count - 1].TU == doc.TU)
                            {
                                unionList.Add(doc);
                                writedNullPosList.Add(doc);

                            }
                        }

                        //Очистка списка 3 от уже задействованных элементов
                        foreach (SpecDocumentWithTU doc in writedNullPosList)
                            nullPosCondRezSpecList.Remove(doc);

                        //Если список 3 не пустой, то записываем оставшиеся элементы в итоговый список
                        if (nullPosCondRezSpecList.Count > 0)
                        {

                            foreach (SpecDocumentWithTU doc in unionList)
                                unionList2.Add(doc);

                            nullPosCondRezSpecList = SortedByNominal(nullPosCondRezSpecList);
                            nullPosCondRezSpecList = SortedByNameDecrease(nullPosCondRezSpecList);

                            //Временный список для формирования группы 
                            var tempListNullPos = new List<SpecDocumentWithTU>();
                            //Временный список для формирования подгруппы для размерности Ом
                            List<SpecDocumentWithTU> tempListOm;
                            //Временный список для формирования подгруппы для размерности кОм
                            List<SpecDocumentWithTU> tempListkOm;
                            //Временный список для формирования подгруппы для размерности МОм
                            List<SpecDocumentWithTU> tempListMOm;
                            //Временный список для формирования подгруппы для размерности мкФ
                            List<SpecDocumentWithTU> tempListmkF;
                            //Временный список для формирования подгруппы для размерности пФ
                            List<SpecDocumentWithTU> tempListpF;

                            //Сортировка по размерности и номиналу временных списков
                            tempListmkF = SortedByMeasureUnit(nullPosCondRezSpecList, "мкФ");
                            tempListpF = SortedByMeasureUnit(nullPosCondRezSpecList, "пФ");
                            tempListMOm = SortedByMeasureUnit(nullPosCondRezSpecList, "МОм");
                            tempListkOm = SortedByMeasureUnit(nullPosCondRezSpecList, "кОм");
                            tempListOm = SortedByMeasureUnit(nullPosCondRezSpecList, "Ом");

                            //Передача данных из временных списков в итоговый 
                            tempListNullPos.AddRange(tempListOm);
                            tempListNullPos.AddRange(tempListkOm);
                            tempListNullPos.AddRange(tempListMOm);
                            tempListNullPos.AddRange(tempListpF);
                            tempListNullPos.AddRange(tempListmkF);

                            nullPosCondRezSpecList = new List<SpecDocumentWithTU>();

                            for (int k = tempListNullPos.Count - 1; k >= 0; k--)
                            {
                                nullPosCondRezSpecList.Add(tempListNullPos[k]);
                            }


                            writedNullPosList = new List<SpecDocumentWithTU>();

                            foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                for (int i = 0; i < unionList2.Count - 1; i++)
                                    //Если наименование прочего изделия в алфавитном порядке стоит между i-ым 
                                    if (unionList2[i].firstPartName != null)
                                        if (unionList2[i].firstPartName == doc.firstPartName)
                                        {
                                            if (i == 0)
                                            {
                                                break;
                                                //MessageBox.Show("break");
                                            }
                                            else
                                                unionList.Add(unionList[unionList.Count - 1]);

                                            //Смещаем элементы на одно местоположение вверх
                                            for (int j = unionList.Count - 2; j > i - 1; j--)
                                            {
                                                unionList[j + 1] = unionList[j];
                                            }
                                            //Вставляем в это место объект с нулевой позицией

                                            unionList[i] = doc;
                                            writedNullPosList.Add(doc);
                                            break;
                                        }

                            //Очистка списка 3 от уже задействованных элементов
                            foreach (SpecDocumentWithTU doc in writedNullPosList)
                            {
                                nullPosCondRezSpecList.Remove(doc);
                            }
                            Regex regexFirstPartName = new Regex(@".*(?=-\d+[.,/]?\d*\s?(мкФ|пФ|Ом|кОм|МОм))");

                            foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                for (int i = 0; i < unionList2.Count - 1; i++)
                                {
                                    //Если наименование прочего изделия в алфавитном порядке стоит между i-ым 
                                    if ((i - 1) >= 0 && String.Compare(unionList2[i - 1].SpecDocument.Name, doc.SpecDocument.Name, StringComparison.Ordinal) < 0
                                        && String.Compare(unionList2[i].SpecDocument.Name, doc.SpecDocument.Name, StringComparison.Ordinal) > 0)
                                    // if ((unionList2[i].SpecDocument.Name.Contains("Резистор") && doc.SpecDocument.Name.Contains("Резистор")) ||
                                    //     (unionList2[i].SpecDocument.Name.Contains("Конденсатор") && doc.SpecDocument.Name.Contains("Конденсатор")))
                                    {
                                        unionList.Add(unionList[unionList.Count - 1]);

                                        //Смещаем элементы на одно местоположение вверх
                                        for (int j = unionList.Count - 2; j > i - 1; j--)
                                        {
                                            unionList[j + 1] = unionList[j];
                                        }
                                        //Вставляем в это место объект с нулевой позицией

                                        unionList[i] = doc;
                                        writedNullPosList.Add(doc);
                                        break;
                                    }
                                }

                            //Очистка списка 3 от уже задействованных элементов
                            foreach (SpecDocumentWithTU doc in writedNullPosList)
                            {
                                nullPosCondRezSpecList.Remove(doc);
                            }

                            //foreach (SpecDocumentWithTU doc in unionList)
                            //    unionList2.Add(doc);

                            //Если список 3 не пустой, то записываем оставшиеся элементы в итоговый список
                            if (nullPosCondRezSpecList.Count > 0)
                            {

                                foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                {

                                    unionList.Add(doc);

                                }
                            }

                            //  if (nullPosCondRezSpecList.Count > 0)
                            //    foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                            //      unionList.Add(doc);
                            /*
                            
                            foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                                unionList.Add(doc);*/
                        }

                    }

                }
                else
                {
                    unionList.Add(withPositionDocSpecList[0]);

                    foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                    {

                        if (withPositionDocSpecList[0].TU == doc.TU)
                        {
                            unionList.Add(doc);
                            writedNullPosList.Add(doc);
                        }
                    }

                    //Очистка списка 3 от уже задействованных элементов
                    foreach (SpecDocumentWithTU doc in writedNullPosList)
                        nullPosCondRezSpecList.Remove(doc);

                    //Если список 3 не пустой, то записываем оставшиеся элементы в итоговый список
                    if (nullPosCondRezSpecList.Count > 0)
                        foreach (SpecDocumentWithTU doc in nullPosCondRezSpecList)
                        {
                            unionList.Add(doc);
                            // MessageBox.Show(doc.TU + "\n" + doc.SpecDocument.Name);
                        }
                }
                return unionList;
            }

            public List<SpecDocumentWithTU> SortSpecDocument(List<SpecDocumentGost2_106> specDocList)
            {
                SpecDocumentWithTU record;
                //Список 1 - список объектов, исключая резисторы и конденсаторы, у которых нет номера позиции 
                List<SpecDocumentWithTU> nullPositionDocSpecList = new List<SpecDocumentWithTU>();

                //Список 2 - список объектов, у которых имеется номер позиции
                List<SpecDocumentWithTU> withPositionDocSpecList = new List<SpecDocumentWithTU>();

                //Список 3 - список подборочных резисторов и конденсаторов(без номера позиции)
                List<SpecDocumentWithTU> nullPosCondRezSpecList = new List<SpecDocumentWithTU>();
                TextNormalizer normal = new TextNormalizer();

                //Шаблон, выделяющий из наименования номер ТУ
                string pattern = @"(?:(\S+(\s)*ТУ)|(ТУ(\s)*\S+)|(\S+ТУ\S+))$";
                Regex regex = new Regex(pattern);
                Match m;
                string tu;
                int count1 = 0;
                int count2 = 0;
                int count3 = 0;

                //Разбиение списка прочих изделий на 3 списка
                foreach (SpecDocumentGost2_106 doc in specDocList)
                {
                    doc.Name = doc.Name.Trim();
                    record = new SpecDocumentWithTU();
                    if ((!doc.Name.Contains("Резистор ")) && (!doc.Name.Contains("Конденсатор ")))
                    {
                        if (doc.PositionNumber > 0)
                        {

                            //Добавление объектов с номером позиции в список 2
                            record.SpecDocument = doc;
                            withPositionDocSpecList.Add(record);
                            count2++;
                        }
                        else
                        {
                            //Добавление объектов (не резисторов и не конденсаторов) без номера позиции в список 1
                            record.SpecDocument = doc;
                            nullPositionDocSpecList.Add(record);
                            count1++;
                        }

                    }
                    else
                    {
                        if (doc.PositionNumber > 0)
                        {
                            m = regex.Match(doc.Name);
                            tu = m.ToString();
                            //Добавление резисторов и конденсаторов с номером позиции в список 2
                            record.SpecDocument = doc;
                            tu = normal.GetNormalForm(tu);
                            record.TU = tu;
                            record = ModifyDoc(record);
                            withPositionDocSpecList.Add(record);
                            count2++;
                        }
                        else
                        {
                            m = regex.Match(doc.Name);
                            tu = m.ToString();
                            //Добавление резисторов и конденсаторов без номера позиции в список 3 
                            record.SpecDocument = doc;
                            tu = normal.GetNormalForm(tu);
                            record.TU = tu;
                            nullPosCondRezSpecList.Add(record);
                            count3++;
                        }
                    }
                }


                List<SpecDocumentWithTU> sortedList = new List<SpecDocumentWithTU>();



                List<SpecDocumentWithTU> unitedObjectsList = new List<SpecDocumentWithTU>();
                //Сортировка по наименованию списка 1
                nullPositionDocSpecList = SortedByName(nullPositionDocSpecList);

                //Объединение списка 2 и списка 3
                if (withPositionDocSpecList.Count > 0)
                    unitedObjectsList = UnionGroupList(withPositionDocSpecList, SortedByTU(nullPosCondRezSpecList));

                //Добавление отсортированных списков в итоговый
                sortedList.AddRange(nullPositionDocSpecList);
                sortedList.AddRange(unitedObjectsList);

                return sortedList;
            }

            //Метод сортировки деталей в списке объектов
            public void SortDetail(List<SpecDocument> specDocumentsList)
            {
                //Список прочих изделий для сортировки
                List<SpecDocumentGost2_106> detailList = new List<SpecDocumentGost2_106>();

                int j = 0;
                //Считывание всех прочих изделий
                foreach (SpecDocumentGost2_106 doc in specDocumentsList)

                    if (TFDBomSection.Details == doc.BomSection)

                        detailList.Add(doc);


                //  MessageBox.Show("перед сортировкой " + detailList.Count);
                // detailList = SortedByPosition(detailList);
                ProductComparer pc = new ProductComparer();
                detailList.Sort(pc);
                detailList = SortedByPosition(detailList);
                // MessageBox.Show("после сортировки " + detailList.Count);

                for (int i = 0; i < specDocumentsList.Count; i++)
                {
                    if (TFDBomSection.Details == specDocumentsList[i].BomSection)
                    {
                        //Запись отсортированных прочих изделий
                        foreach (SpecDocumentGost2_106 specDoc in detailList)
                        {
                            specDocumentsList[i + j] = specDoc;
                            j++;
                        }
                        break;
                    }
                }
                //   MessageBox.Show("записал " + j  + " деталей"  );

            }


            //Метод сортировки прочих изделий в списке объектов
            public void SortOtherItems(List<SpecDocument> specDocumentsList)
            {
                //Список прочих изделий для сортировки
                List<SpecDocumentGost2_106> otherItemsList = new List<SpecDocumentGost2_106>();


                //Считывание всех прочих изделий
                foreach (SpecDocumentGost2_106 doc in specDocumentsList)
                    if (TFDBomSection.OtherProducts == doc.BomSection)
                        otherItemsList.Add(doc);

                int correctListCount = otherItemsList.Count;
                //Список отсортированных прочих изделий
                List<SpecDocumentWithTU> sortedSpecDocList = new List<SpecDocumentWithTU>();


                //Сортировка прочих изделий	
                sortedSpecDocList = SortSpecDocument(otherItemsList);
                int j = 0;

                if (correctListCount - sortedSpecDocList.Count != 0)
                    MessageBox.Show("Ошибка некорректная сортировка и запись объектов прочих изделий. Выводимый отчет некорректен. Обратитесь к Администратору для устранения ошибки.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                for (int i = 0; i < specDocumentsList.Count; i++)
                {
                    if (TFDBomSection.OtherProducts == specDocumentsList[i].BomSection)
                    {
                        //Запись отсортированных прочих изделий
                        foreach (SpecDocumentWithTU specDoc in sortedSpecDocList)
                        {
                            specDocumentsList[i + j] = specDoc.SpecDocument;
                            j++;

                        }
                        break;
                    }
                }
            }

            //Метод сортировки стандартных изделий в списке объектов
            public void SortStdProduct(List<SpecDocument> specDocumentsList)
            {
                //Список прочих изделий для сортировки
                List<SpecDocumentGost2_106> stdItemsList = new List<SpecDocumentGost2_106>();

                int j = 0;
                //Считывание всех прочих изделий
                foreach (SpecDocumentGost2_106 doc in specDocumentsList)
                    if (TFDBomSection.StdProducts == doc.BomSection)
                        stdItemsList.Add(doc);

                stdItemsList = SortedByPosition(stdItemsList);

                for (int i = 0; i < specDocumentsList.Count; i++)
                {
                    if (TFDBomSection.StdProducts == specDocumentsList[i].BomSection)
                    {
                        //Запись отсортированных прочих изделий
                        foreach (SpecDocumentGost2_106 specDoc in stdItemsList)
                        {
                            specDocumentsList[i + j] = specDoc;
                            j++;
                        }
                        break;
                    }
                }
            }




            //Метод сортировки документов по обозначению
            public void SortDocuments(List<SpecDocument> specDocumentsList)
            {

                //Список документов для сортировки
                List<SpecDocumentGost2_106> documentsList = new List<SpecDocumentGost2_106>();

                List<SpecDocumentGost2_106> sorteDocumentsList = new List<SpecDocumentGost2_106>();

                // string pattern = @"^(?<prefix>.*)(?<![\sА-Я])(?<![\sА-Я]+\S)(?<![-0-9]{4,})";
                //string pattern = @"^(\S{2,}(\s+(?=\s|$))*)(?<![\sА-Я])(?<![-\w]{4,})";
                string pattern = @"^(\S{2,}(\s+(?=\s|$))*)(?<![\sА-Я])(?<![-\w]{4})(?<![-\w]{4})(?<![И1.1]+\S)";
                string patternProgramSoft = @"^([А-Я]{4})\.([0-9]{5})-([0-9]{2})$";
                string patternProgramSoftForDoc = @"^([А-Я]{4}\.\d{5}-\d{2}\s\d{2}\s\d{2})$";
                string programSoftDenotation = string.Empty;

                // Документы содержащие обозначения программного обеспечения
                List<SpecDocumentGost2_106> docSoftProgramForDoc = new List<SpecDocumentGost2_106>();
                //Комплекты (программы) 
                List<SpecDocumentGost2_106> docSoftProgram = new List<SpecDocumentGost2_106>();

                Match m;
                bool flag = false;
                List<string> denotationList = new List<string>();

                //Считывание всех документов за исключением документов имеющих обозначение программного обеспечения
                foreach (SpecDocumentGost2_106 doc in specDocumentsList)
                {
                    if (TFDBomSection.Documentation != doc.BomSection)
                    {
                        Regex regexProgramSoft = new Regex(patternProgramSoft);
                        m = regexProgramSoft.Match(doc.Designation);
                        programSoftDenotation = m.ToString().Trim();
                        //***** Делаем комплектами объекты с обозначением программного обеспечения? ЗАЧЕМ???????????
                        //if (programSoftDenotation != string.Empty)
                        //{
                        //    doc.BomSection = TFDBomSection.Komplekts;
                        //}
                    }
                    else
                    {
                        Regex regexProgramSoft = new Regex(patternProgramSoftForDoc);
                        m = regexProgramSoft.Match(doc.Designation);
                        programSoftDenotation = m.ToString().Trim();
                        if (programSoftDenotation == string.Empty)
                            documentsList.Add(doc);
                        else docSoftProgramForDoc.Add(doc);
                    }
                }

                foreach (SpecDocumentGost2_106 doc in specDocumentsList)
                    if (TFDBomSection.Komplekts == doc.BomSection)
                    {
                        docSoftProgram.Add((doc));
                    }

                foreach (SpecDocumentGost2_106 doc in docSoftProgram)
                    specDocumentsList.Remove(doc);



                SortKomplekt(specDocumentsList, docSoftProgram);

                //Сортировка документов имеющих обозначение программного обеспечения 
                if (docSoftProgramForDoc.Count > 1)
                    docSoftProgramForDoc = SortedByDenotation(docSoftProgramForDoc);

                Regex regex = new Regex(pattern);
                //Список отсортированных прочих изделий
                List<SpecDocumentWithTU> sortedSpecDocList = new List<SpecDocumentWithTU>();

                SpecDocumentWithTU spec;

                //Сортировка документов	
                //sortedSpecDocList = SortedByDenotation(sortedSpecDocList);


                foreach (SpecDocumentGost2_106 doc in documentsList)
                {
                    spec = new SpecDocumentWithTU();
                    spec.SpecDocument = doc;
                    m = regex.Match(doc.Designation);
                    spec.TU = m.ToString().Trim();
                    sortedSpecDocList.Add(spec);
                    flag = false;
                    foreach (string designation in denotationList)
                        if (designation == m.ToString().Trim())
                            flag = true;

                    if (flag == false)
                        denotationList.Add(m.ToString().Trim());

                }


                // сортровка обозначений по порядку (обозначения БФМИ.460010.005 и БФМИ.460086.006 всегда ставим на первое место)
                denotationList = SortedByDenotation(denotationList);


                foreach (string denotation in denotationList)
                {
                    foreach (SpecDocumentWithTU specDoc in sortedSpecDocList)
                        if (denotation == specDoc.TU)
                            sorteDocumentsList.Add(specDoc.SpecDocument);
                }


                // sorteDocumentsList = sorteDocumentsList.OrderBy(t => t.Designation).ToList();
                // var sortDocumentsListByAlphabet = new List<SpecDocumentGost2_106>();


                int j = 0;

                // Запись в конец отсортированного списка документов, документов с обозначением программного обеспечения


                if (docSoftProgramForDoc.Count > 0)
                    sorteDocumentsList.AddRange(docSoftProgramForDoc);

                for (int i = 0; i < specDocumentsList.Count; i++)
                {
                    if (TFDBomSection.Documentation == specDocumentsList[i].BomSection)
                    {
                        //Запись отсортированных прочих изделий
                        foreach (SpecDocumentGost2_106 specDoc in sorteDocumentsList)
                        {
                            specDocumentsList[i + j] = specDoc;
                            j++;
                        }
                        break;
                    }
                }
            }

            //----------------Сортировка раздела спецификации Комплекты----------------
            private void SortKomplekt(List<SpecDocument> specDocumentsList, List<SpecDocumentGost2_106> docSoftProgram)
            {
                var FirstGroupKomplekt = new List<SpecDocumentGost2_106>();
                var SecondGroupKomplekt = new List<SpecDocumentGost2_106>();
                var ThirdGroupKomplekt = new List<SpecDocumentGost2_106>();
                var FourthGroupKomplekt = new List<SpecDocumentGost2_106>();

                foreach (var doc in docSoftProgram)
                {
                    if (doc.BomSection == TFDBomSection.Komplekts)
                    {
                        if (doc.Name.ToLower().Contains("эксплуатац"))
                        {
                            FirstGroupKomplekt.Add(doc);
                        }
                        else
                        {
                            if (doc.Name.ToLower().Contains("одиночн"))
                            {
                                SecondGroupKomplekt.Add(doc);
                            }
                            else
                            {
                                if (doc.Name.ToLower().Contains("упаковка"))
                                {
                                    ThirdGroupKomplekt.Add(doc);
                                }
                                else
                                {
                                    FourthGroupKomplekt.Add(doc);
                                }
                            }
                        }
                    }
                }

                specDocumentsList.AddRange(SortedByDenotation(FirstGroupKomplekt));
                specDocumentsList.AddRange(SortedByDenotation(SecondGroupKomplekt));
                specDocumentsList.AddRange(SortedByDenotation(ThirdGroupKomplekt));
                specDocumentsList.AddRange(SortedByDenotation(FourthGroupKomplekt));
            }

            //Метод, включающий в постоянную часть объекты со словом "(Заготовка)", если такие объекты с похожим БФМИ уже присутствуют в переменной части
            public void DeletePreparation(List<SpecDocument> specDocumentsList, SpecDocumentCollection varDocs)
            {

                // Регулярное выражение для обозначений заготовок и соответствующих им элементам в переменной части
                string denotationPrep = @"^(?<o>[А-Я]{4})\.(?<g2>[0-9]{6}).";
                Regex regexDenotation = new Regex(denotationPrep);
                Match matchDoc;
                Match matchMakeDoc;

                int countNamePrep = 0;


                foreach (SpecDocument doc in specDocumentsList)
                {
                    if (doc.Name.Contains("(Заготовка)"))
                    {
                        countNamePrep = doc.Name.IndexOf(" ");
                        matchDoc = regexDenotation.Match(doc.Designation);

                        for (int j = 0; j < varDocs.Count - 1; j++)
                        {

                            matchMakeDoc = regexDenotation.Match(varDocs[j].Designation);
                            if (countNamePrep > 0)
                                // Если совпадают первые части обозначений
                                if ((matchDoc.Value == matchMakeDoc.Value) &&
                                // И к этому объекту не было уже приписана заготовка из постоянной части
                                (varDocs[j].AddDocMake == false) &&
                                // И наименование до пробела постоянной части (без слова заготовка) совпадает с наименованием из переменной части 
                                (doc.Name.Substring(0, countNamePrep) == varDocs[j].Name.Substring(0, countNamePrep)))
                                {
                                    // То добавляем объект из постоянной части в переменную
                                    varDocs[j].AddDocMake = true;

                                    SpecDocument spDoc = varDocs[j];
                                    varDocs[j] = doc;

                                    for (int i = j + 1; i < varDocs.Count; i++)
                                    {
                                        SpecDocument spDoc2 = varDocs[i];
                                        varDocs[i] = spDoc;
                                        spDoc = spDoc2;
                                    }
                                    varDocs.Add(spDoc);


                                    break;
                                }
                        }


                    }
                }


            }



            /// Добавление строк спецификации для всей коллекции изделий (для всех исполнений)
            public override void AddSpecRowsForm1(SpecForm1 specForm, bool fullNames, DocumentAttributes documentAttr, string docName, string denotation)
            {
                // Получение документов общего раздела спецификации
                SpecDocumentCollection commonSpecDocuments;
                if (documentAttr.AddMake)
                {
                    commonSpecDocuments = GetCommonSpecPartRows();
                    foreach (var doc in commonSpecDocuments)
                    {
                        if (doc.Name.Contains("###"))
                        {
                            doc.Name.Remove(doc.Name.Length - 3, 3);
                        }
                    }


                }
                else
                {
                    commonSpecDocuments = new SpecDocumentCollection();
                    commonSpecDocuments.AddRange(Documents);
                }


                SpecDocumentCollection materials = new SpecDocumentCollection();
                //***
                //вставка   //Сортировка Прочих изделий в общем списке
                SortOtherItems(commonSpecDocuments);

                int countReplacing = 0;


                foreach (var doc in commonSpecDocuments)
                {
                    if (doc.Note.Contains("См. примечание") && doc.ReplacingObject != null)
                    {
                        countReplacing++;
                        if (commonSpecDocuments.Where(t => t.ReplacingObject != null).ToList().Count > 1)
                            doc.Note = doc.Note.Replace("См. примечание", "См. примечание" + countReplacing);
                        documentAttr.ReplacingObjects.Add(doc.ReplacingObject);
                    }
                }


                /*       foreach (SpecDocumentGost2_113 doc in commonSpecDocuments)
                       {
                           if ((doc.Name.Contains("Микросхема")) && (doc.PositionNumber ==94))
                               MessageBox.Show(doc.Name + " " + doc.PositionNumber + " " + doc.Note);
                       }
                       */
                //Сортировка Стандартных изделий в общем списке
                SortStdProduct(commonSpecDocuments);


                //Сортировка Деталей в общем списке
                SortDetail(commonSpecDocuments);

                SortDocuments(commonSpecDocuments);
                //***            
                //Сортировка комплектов и программых изделий в общем списке
                //SortComplements(commonSpecDocuments);
                // Получение записей для исполнений
                if (_products.Count > 1)
                {

                    // Получение коллекций документов, входящих в переменную часть
                    // для каждого из исполнений
                    Dictionary<Product, SpecDocumentCollection> varParts
                        = new Dictionary<Product, SpecDocumentCollection>();
                    bool haveAnyVarContent = false;
                    foreach (Product make in _products)
                    {
                        // получение переменной части
                        SpecDocumentCollection varDocs
                            = GetMakeSpecDocuments(make, commonSpecDocuments, AuthenticityComparer, materials);
                        if (varDocs.Count > 0)
                            haveAnyVarContent = true;
                        List<SpecDocumentGost2_106> removingCommonDocs = new List<SpecDocumentGost2_106>();

                        if ((make.Designation == rootDocument.Denotation) && (!documentAttr.AddMake))
                        {



                            //Сортировка Прочих изделий из переменной части в общем списке
                            SortOtherItems(varDocs);



                            foreach (var doc in varDocs)
                            {
                                if (doc.Note.Contains("См. примечание") && doc.ReplacingObject != null)
                                {
                                    countReplacing++;
                                    doc.Note = doc.Note.Replace("См. примечание", "См. примечание" + countReplacing);
                                    documentAttr.ReplacingObjects.Add(doc.ReplacingObject);
                                }
                            }

                            varParts.Add(make, varDocs);

                            // Удаление заготовок из постоянной части
                            DeletePreparation(commonSpecDocuments, varDocs);

                            // Удаляем из постоянной части объекты, которые имеют совпадающие обозначение
                            // и наименование с одним из объектов переменной части. Это делается потому, что
                            // в составе объекта может быть несколько объектов, имеющих одинаковое обозначение
                            // и наименование, но разные количество или обозначение на схеме. Если один из таких
                            // объектов включен в переменную часть, то остальные подобные объекты должны быть
                            // перенесены в переменную часть.

                            foreach (SpecDocumentGost2_106 commonDoc in commonSpecDocuments)
                            {
                                if (varDocs.Contains(commonDoc, _baseEqualityComparer))
                                    removingCommonDocs.Add(commonDoc);

                            }

                            foreach (SpecDocumentGost2_106 remove in removingCommonDocs)
                            {
                                commonSpecDocuments.Remove(remove);
                            }
                            break;
                        }


                        // Удаление заготовок из постоянной части
                        DeletePreparation(commonSpecDocuments, varDocs);

                        //Сортировка Деталей в общем списке
                        SortDetail(varDocs);

                        //Сортировка Прочих изделий из переменной части в общем списке
                        SortOtherItems(varDocs);

                        foreach (var doc in varDocs)
                        {
                            if (doc.Note.Contains("См. примечание") && doc.ReplacingObject != null)
                            {
                                countReplacing++;
                                doc.Note = doc.Note.Replace("См. примечание", "См. примечание" + countReplacing);
                                documentAttr.ReplacingObjects.Add(doc.ReplacingObject);
                            }
                        }

                        varParts.Add(make, varDocs);
                        // Удаляем из постоянной части объекты, которые имеют совпадающие обозначение
                        // и наименование с одним из объектов переменной части. Это делается потому, что
                        // в составе объекта может быть несколько объектов, имеющих одинаковое обозначение
                        // и наименование, но разные количество или обозначение на схеме. Если один из таких
                        // объектов включен в переменную часть, то остальные подобные объекты должны быть
                        // перенесены в переменную часть.


                        foreach (SpecDocumentGost2_106 commonDoc in commonSpecDocuments)
                        {
                            if (varDocs.Contains(commonDoc, _baseEqualityComparer) && commonDoc.BomSection != TFDBomSection.Materials)
                                removingCommonDocs.Add(commonDoc);
                        }

                        foreach (SpecDocumentGost2_106 remove in removingCommonDocs)
                        {
                            commonSpecDocuments.Remove(remove);
                        }


                    }

                    //ошибка с материалами
                    // Добавление материалов в постоянную часть
                    // if (materials.Count > 0)
                    // {
                    //     commonSpecDocuments.AddRange(materials);
                    //  MessageBox.Show("count " + materials.Count);
                    // }

                    bool addVarPart = false;
                    var listVarDoc = new List<SpecDocumentGost2_106>();
                    int countDel = 0;
                    string namePos = string.Empty;

                    if (documentAttr.ThroughNumbering)
                    {
                        if (documentAttr.AddMake && haveAnyVarContent)
                        {
                            int? minPosMake = 100000;
                            int? minPos;

                            foreach (var objList in varParts.Values)
                            {

                                minPos = 100000;
                                foreach (SpecDocumentGost2_106 specDocument in objList)
                                {

                                    if (specDocument.PositionNumber != null && specDocument.PositionNumber != 0 &&
                                        specDocument.PositionNumber < minPos)
                                    {
                                        minPos = specDocument.PositionNumber;
                                        namePos = specDocument.Name + " позиция: " + specDocument.PositionNumber;
                                    }
                                }

                                if (minPosMake > minPos)
                                {
                                    minPosMake = minPos;

                                }

                            }

                            //var listVarDoc = new List<SpecDocumentGost2_106>();

                            foreach (SpecDocumentGost2_106 specDocument in commonSpecDocuments)
                            {


                                if (specDocument.PositionNumber > minPosMake)
                                {
                                    listVarDoc.Add(specDocument);

                                }
                            }

                            if (listVarDoc.Count > 0)
                                addVarPart = true;
                            //!!!!!!!!!!!!!!!

                            foreach (var varDoc in listVarDoc)
                            {
                                commonSpecDocuments.Remove(varDoc);
                                countDel++;
                            }


                        }
                    }

                    if (countDel > 0)
                        MessageBox.Show(
              "Спецификация сформирована некорректно. Вы выбрали сквозную нумерацию страниц, всвязи с этим было удалено " + countDel +
              "  объектов. Для корректного отчета либо отожмите флажок 'Сквозная нумерация', либо измените позиции начиная с элемента:\n" + namePos,
              "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    // Добавление постоянной части спецификации
                    SpecDocumentGost2_106[] commonSpecDocumentsArray
                        = new SpecDocumentGost2_106[commonSpecDocuments.Count];

                    commonSpecDocuments.CopyTo(commonSpecDocumentsArray);

                    AddSpecSectionsAndRows(specForm, commonSpecDocumentsArray, documentAttr, fullNames, docName, denotation, false);

                    if (documentAttr.AddMake)
                    {
                        // Заголовок "Переменные данные для исполнений"
                        AddVariablePartHeader(specForm);
                        bool makeHeader = true;

                        if (!haveAnyVarContent)
                        {
                            // Ни одно из исполнений не имеет переменной части.
                            // Вывод заголовока "Различие исполнений по сборочному чертежу"
                            AddDifferenceByAssemblyDrawingHeader(specForm);
                        }
                        else
                        {


                            if (addVarPart)
                                for (int k = 0; k < varParts.Values.Count; k++)
                                {
                                    var listObj = varParts.Values.ToList()[k];


                                    foreach (SpecDocumentGost2_106 specDocumentGost2106 in listVarDoc)
                                    {

                                        int i = 0;
                                        while (i <= listObj.Count)
                                        {

                                            if ((i == 0) && (listObj[i].BomSection == specDocumentGost2106.BomSection) &&
                                                ((SpecDocumentGost2_106)listObj[i]).PositionNumber >
                                                     specDocumentGost2106.PositionNumber)
                                            {

                                                listObj.Insert(i + 1, specDocumentGost2106);
                                                break;
                                            }



                                            if ((i == listObj.Count - 1) && (listObj[i].BomSection == specDocumentGost2106.BomSection) &&
                                                ((SpecDocumentGost2_106)listObj[i]).PositionNumber <
                                                     specDocumentGost2106.PositionNumber)
                                            {
                                                listObj.Insert(i + 1, specDocumentGost2106);
                                                break;
                                            }

                                            if ((i < listObj.Count - 1) && (listObj[i].BomSection == specDocumentGost2106.BomSection)
                                            && (listObj[i + 1].BomSection == specDocumentGost2106.BomSection) &&
                                        (((SpecDocumentGost2_106)listObj[i]).PositionNumber <
                                                 specDocumentGost2106.PositionNumber) &&
                                                (((SpecDocumentGost2_106)listObj[i + 1]).PositionNumber >
                                                 specDocumentGost2106.PositionNumber))
                                            {
                                                listObj.Insert(i + 1, specDocumentGost2106);
                                                break;
                                            }


                                            if ((i < listObj.Count - 1) && (listObj[i].BomSection == specDocumentGost2106.BomSection)
                                                && (listObj[i + 1].BomSection != specDocumentGost2106.BomSection)
                                                && ((SpecDocumentGost2_106)listObj[i]).PositionNumber <
                                                     specDocumentGost2106.PositionNumber)
                                            {
                                                listObj.Insert(i + 1, specDocumentGost2106);
                                                break;
                                            }


                                            if ((i < listObj.Count - 1) && (listObj[i].BomSection != specDocumentGost2106.BomSection)
                                                && (listObj[i + 1].BomSection == specDocumentGost2106.BomSection) &&
                                                ((SpecDocumentGost2_106)listObj[i + 1]).PositionNumber >
                                                     specDocumentGost2106.PositionNumber)
                                            {
                                                listObj.Insert(i + 1, specDocumentGost2106);
                                                break;
                                            }

                                            i++;
                                        }
                                    }



                                }

                            // Одно или несколько (может быть все) исполнений имеют переменную часть
                            // Вывод переменной части спецификации
                            AddVariablePartRows(varParts, commonSpecDocuments, specForm, documentAttr, fullNames,
                                                docName, denotation);

                        }
                    }


                } // _products.Count > 1
                else
                {
                    // Добавление постоянной части спецификации
                    SpecDocumentGost2_106[] commonSpecDocumentsArray
                        = new SpecDocumentGost2_106[commonSpecDocuments.Count];

                    commonSpecDocuments.CopyTo(commonSpecDocumentsArray);

                    AddSpecSectionsAndRows(specForm, commonSpecDocumentsArray, documentAttr, fullNames, docName, denotation, false);
                }


                // }

                if ((SpecDocument as SpecDocumentGost2_113).IsBaseMake)
                {
                    // Если спецификация формируется для документа, представляющего
                    // собой базовое исполнение, добавляем строку
                    // "Остальное - см. спецификацию исполнения"
                    AddOtherLookMakeSpec(specForm);
                }



            }


            /// Добавление строк спецификации для переменной части спецификации
            /// @param varParts
            ///     Словарь, содержащий изделие и его переменную часть. Должен быть
            ///     отсортирован в порядке записи переменных частей.
            /// @param commonSpecDocuments
            ///     Коллекция документов, составляющих постоянную часть спецификации
            void AddVariablePartRows(Dictionary<Product, SpecDocumentCollection> varParts,
                SpecDocumentCollection commonSpecDocuments, SpecForm1 specForm, DocumentAttributes documentAttr, bool fullNames, string docName, string denotation)
            {
                bool firstVarPart = true;
                foreach (KeyValuePair<Product, SpecDocumentCollection> currentMakePair in varParts)
                {
                    Product currentMake = currentMakePair.Key;
                    SpecDocumentCollection currentVarDocs = currentMakePair.Value;

                    if (!firstVarPart)
                        specForm.NewPage();
                    AddMakeHeader(specForm, currentMake);

                    if (currentVarDocs.Count == 0)
                    {
                        // Текущее исполнение не имеет переменной части.
                        // Вывод записи "Отсутствуют"
                        AddNoVariablePartHeader(specForm);
                    }
                    else
                    {
                        // Текущее исполнение имеет переменную часть.
                        // Проверка на совпадение с составом одного из предыдущих исполнений.
                        Product theSameContentProduct = null;
                        foreach (KeyValuePair<Product, SpecDocumentCollection> prevPair in varParts)
                        {
                            Product prevMake = prevPair.Key;
                            SpecDocumentCollection prevVarDocs = prevPair.Value;

                            if (prevMake == currentMake)
                                break;

                            bool equalsContent = currentVarDocs.EqualsContent(prevVarDocs, AuthenticityComparer);
                            if (equalsContent)
                            {
                                theSameContentProduct = prevMake;
                                break;
                            }
                        }



                        if (theSameContentProduct != null)
                        {
                            // Обнаружено выведенное ранее исполнение с идентичной переменной частью.
                            // Вывод записи "Также как для ..."
                            AddTheSameAsForHeader(specForm, theSameContentProduct.Designation);
                        }
                        else
                        {
                            // Занесение переменных данных исполнения
                            AddMakeSpecRows(specForm, currentVarDocs, documentAttr, fullNames, docName, denotation, true);
                        }
                    }
                    firstVarPart = false;
                }
            }

            /// Вывод заголовка "Переменные данные для исполнений"
            void AddVariablePartHeader(SpecForm1 specForm)
            {
                if (!specForm.AtTopOfPage())
                    specForm.NewPage();

                specForm.AddEmptyRow();
                SpecRowForm1 varHeader = new SpecRowForm1(
                    SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                    new SpecFormCell("Переменные данные", SpecFormCell.UnderliningFormat.Underlined,
                        SpecFormCell.AlignFormat.Right),
                    new SpecFormCell("для исполнений", SpecFormCell.UnderliningFormat.Underlined,
                        SpecFormCell.AlignFormat.Left),
                    SpecFormCell.Empty, SpecFormCell.Empty);
                specForm.Add(varHeader);
                specForm.AddEmptyRow();
            }

            /// Вывод заголовока "Различие исполнений по сборочному чертежу"
            void AddDifferenceByAssemblyDrawingHeader(SpecForm1 specForm)
            {
                SpecRowForm1 none = new SpecRowForm1(
                    SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                    new SpecFormCell("Различие исполнений",
                        SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Right),
                    new SpecFormCell("по сборочному чертежу"),
                    SpecFormCell.Empty, SpecFormCell.Empty);
                specForm.Add(none);
                specForm.AddEmptyRow();
            }

            /// Вывод текста "Отсутствуют"
            /// (в случае отсутствия переменной части у исполнения)
            void AddNoVariablePartHeader(SpecForm1 specForm)
            {
                SpecRowForm1 none = new SpecRowForm1(
                    SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                    new SpecFormCell("Отсутствуют"),
                    SpecFormCell.Empty, SpecFormCell.Empty);
                specForm.Add(none);
                specForm.AddEmptyRowsOrFinishCurrentPage(3);
            }

            /// Вывод текста "Также как для..."
            void AddTheSameAsForHeader(SpecForm1 specForm, string designation)
            {
                // Вывод записи "Также как для ..."
                SpecRowForm1 same = new SpecRowForm1(
                    SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty, SpecFormCell.Empty,
                    new SpecFormCell("Также как для " + designation),
                    SpecFormCell.Empty, SpecFormCell.Empty);
                specForm.Add(same);
                specForm.AddEmptyRowsOrFinishCurrentPage(3);
            }

            /// Вывод текста "Остальное - см. спецификацию исполнения"
            void AddOtherLookMakeSpec(SpecForm1 specForm)
            {
                SpecRowForm1 row = new SpecRowForm1(
                    new SpecFormCell("Остальное - см. ", SpecFormCell.UnderliningFormat.Underlined,
                        SpecFormCell.AlignFormat.Right),
                    new SpecFormCell("спецификацию исполнения", SpecFormCell.UnderliningFormat.Underlined,
                        SpecFormCell.AlignFormat.Left)
                    );
                specForm.Add(row);
            }

            /// Получение записей спецификации, составляющих общую часть (общий раздел)
            protected virtual SpecDocumentCollection GetCommonSpecPartRows()
            {
                // Документы, входящие в постоянную часть спецификации
                SpecDocumentCollection constPartDocuments = new SpecDocumentCollection();
                // Документы, проверяемые на возможность включения в постоянную часть СП
                SpecDocumentCollection candidatesQueue = new SpecDocumentCollection();
                candidatesQueue.AddRange(Documents);

                while (candidatesQueue.Count > 0)
                {
                    // Получение набора документов-кандидатов в постоянную часть. В набор
                    // включаются документы с одинаковыми обозначением и наименованием
                    SpecDocument firstCandidate = candidatesQueue[0];
                    SpecDocumentCollection candidates = new SpecDocumentCollection();
                    candidates.Add(firstCandidate);
                    foreach (SpecDocument doc in candidatesQueue)
                    {
                        if (doc == firstCandidate)
                            continue;
                        else if (_baseEqualityComparer.Equals(firstCandidate, doc))
                            candidates.Add(doc);
                    }
                    foreach (var cand in candidates)
                    {


                        // Удаление набора кандидатов из очереди
                        //  foreach (SpecDocument doc in candidates)
                        candidatesQueue.Remove(cand);

                        // Проверка набора кандидатов на возможность включения в постоянную часть СП
                        bool candidatesAreConst = true;
                        // Цикл по исполнениям (кроме основного)

                        foreach (Product product in _products)
                        {
                            if (product == this)
                                continue;
                            // Проверка каждого документв из candidates на наличие в исполнении
                            // foreach (SpecDocument candidate in candidates)
                            {
                                bool containsInProduct = false;
                                // Цикл по документам текущего исполнения
                                foreach (SpecDocument otherDoc in product.Documents)
                                {
                                    if (AuthenticityComparer.Equals(cand, otherDoc))
                                    {
                                        containsInProduct = true;
                                        break;
                                    }
                                }
                                if (!containsInProduct)
                                {
                                    candidatesAreConst = false;
                                    //if (candidate.Name.Contains("Пиломатериал"))
                                    //    MessageBox.Show(candidate.Name + " " + candidate.Position);
                                    break;
                                }
                            }
                            if (!candidatesAreConst)
                                break;
                        }
                        if (candidatesAreConst)
                        {
                            //   foreach (SpecDocument doc in candidates)
                            constPartDocuments.Add(cand);
                        }

                    }
                }

                // Документы, входящие в постоянную часть спецификации
                //SpecDocumentCollection constPart = new SpecDocumentCollection();
                //constPart.AddRange(constPartDocuments);
                //foreach (var product in _products)
                //{
                //    foreach (SpecDocument otherDoc in product.Documents)
                //    {
                //       
                //    }

                //}
                //foreach (var constDoc in constPartDocuments)
                //{


                //        MessageBox.Show(constDoc.Name + " " + constDoc.Position);

                //}
                return constPartDocuments;
            }

            /// Добавление заголовка переменных данных исполнения
            protected virtual void AddMakeHeader(SpecForm1 specForm, Product product)
            {
                // Пустая строка
                specForm.AddEmptyRow();

                // Обозначение исполнения (подчеркнутое)
                string denotation = product.Designation;

                // Вырезаем из обозначения последние два символа (ПИ) по требованию Баскаковой
                string pi = denotation.Trim().Remove(denotation.Length - 2, 2);

                if (product.Designation.Contains("ПИ") && !pi.Contains("ПИ"))
                {
                    denotation = product.Designation.Trim().Remove(product.Designation.Length - 2, 2);
                }
                if (product.Designation.Contains("ПИ") && !(denotation.Trim().Remove(denotation.Length - 5, 2)).Contains("ПИ"))
                {
                    denotation = product.Designation.Trim().Remove(product.Designation.Length - 4, 2);

                }
                if (product.Designation.Contains("ПИ") && !(denotation.Trim().Remove(denotation.Length - 6, 2)).Contains("ПИ"))
                {
                    denotation = product.Designation.Trim().Remove(product.Designation.Length - 5, 2);

                }

                SpecFormCell productHeaderCell
                    = new SpecFormCell(denotation,
                                       SpecFormCell.UnderliningFormat.Underlined,
                                       SpecFormCell.AlignFormat.Center);
                SpecRowForm1 productHeader = new SpecRowForm1(
                    SpecFormCell.Empty, productHeaderCell);
                specForm.Add(productHeader);

                // Литера. Берется последняя литера из списка
                SpecDocumentGost2_106 specDoc = product.SpecDocument as SpecDocumentGost2_106;
                string lit = null;
                // if (specDoc.Litera.Length > 0)
                //     lit = specDoc.Litera[specDoc.Litera.Length - 1];
                // if (lit == null || lit.Length == 0)
                //   lit = "   "; // three spaces
                string litStr = String.Format("Лит. \"{0}\"", lit);
                SpecFormCell litCell
                    = new SpecFormCell(litStr,
                        SpecFormCell.UnderliningFormat.NotUnderlined, SpecFormCell.AlignFormat.Center);
                SpecRowForm1 litRow = new SpecRowForm1(SpecFormCell.Empty, litCell);
                specForm.Add(litRow);

                // Пустая строка
                specForm.AddEmptyRow();
            }

            /// Добавление в таблицу спецификации переменных данных исполнения
            /// @param specForm
            ///     Форма спецификации
            /// @param product
            ///     Изделие (исполнение)
            /// @param exclude
            ///     Исключаемые документы. Например, вошедшие в постоянные данные
            protected virtual void AddMakeSpecRows(SpecForm1 specForm, SpecDocumentCollection documents, DocumentAttributes documentAttr, bool fullNames, string docName, string denotation, bool varPartDoc)
            {
                SpecDocumentGost2_106[] documentsArray = new SpecDocumentGost2_106[documents.Count];
                documents.CopyTo(documentsArray);
                AddSpecSectionsAndRows(specForm, documentsArray, documentAttr, fullNames, docName, denotation, varPartDoc);
            }

            /// Получение коллекции документов изделия (исполнения) за исключением определенного
            /// набора документов
            SpecDocumentCollection GetMakeSpecDocuments(Product product,
               SpecDocumentCollection commonSpecDocuments, SpecDocumentBaseEqualityComparer comparer, SpecDocumentCollection materials)
            {
                foreach (SpecDocumentGost2_113 doc in commonSpecDocuments)
                {
                    doc.Name = doc.Name.Trim();
                }

                SpecDocumentCollection documents = new SpecDocumentCollection();
                foreach (SpecDocumentGost2_113 doc in product.Documents)
                {
                    doc.Name = doc.Name.Trim();
                    if (!commonSpecDocuments.Contains(doc, comparer))
                    // foreach (SpecDocumentGost2_113 comDoc in commonSpecDocuments)
                    //      if ((doc.Name != comDoc.Name) || (doc.Designation != comDoc.Designation))
                    {

                        //ошибка с материалами
                        // if (doc.BomSection == TFDBomSection.Materials)
                        //      materials.Add(doc);
                        //  else
                        //     if (!materials.Contains(doc))
                        //     {
                        //          documents.Add(doc);
                        //
                        //     }
                        /* if ((doc.Name.Contains("Микросхема")) && (doc.PositionNumber == 94))
                            {
                                //   documents.Add(doc);
                                MessageBox.Show(doc.Name + " " + doc.PositionNumber + " " + doc.Note );
                            }*/
                        documents.Add(doc);
                        foreach (SpecDocumentGost2_113 comDoc in commonSpecDocuments)

                            if ((doc.Designation != comDoc.Designation) && (doc.Name != comDoc.Name) &&
                                  (doc.Name.Contains("Микросхема")) && (doc.PositionNumber == 94) && (comDoc.Name.Contains("Микросхема")))
                            {
                                //   documents.Add(doc);
                                //  MessageBox.Show(doc.Name + " " + doc.PositionNumber + " " + doc.Note + "\n" + comDoc.Name + " " + comDoc.PositionNumber + " " + comDoc.Note);
                            }

                    }

                }
                return documents;
            }
        }

        internal class SpecDocumentComparerGost2_113
            : SpecDocumentComparerGost2_106, IComparer<SpecDocumentGost2_113>
        {


            public override int Compare(SpecDocument doc1, SpecDocument doc2)
            {
                Debug.Assert(doc1 is SpecDocumentGost2_113);
                Debug.Assert(doc2 is SpecDocumentGost2_113);

                int cmp = Compare((SpecDocumentGost2_113)doc1, (SpecDocumentGost2_113)doc2);

                return cmp;
            }

            public virtual int Compare(SpecDocumentGost2_113 doc1, SpecDocumentGost2_113 doc2)
            {
                // Сравнение по разделу спецификации

                int cmp = CompareByBomGroup(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                // Сравнение по флагу о базовом исполнении

                cmp = CompareByBaseMakeFlag(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                // Сравнение по позиционному порядковому номеру позиции

                cmp = CompareByPositionDesignationAndName(doc1, doc2, sortByAlphabet);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByRefdes(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByCount(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByNote(doc1, doc2);
                if (cmp != 0)
                    return cmp;

                cmp = CompareByID(doc1, doc2);
                return cmp;
            }

            /**
             *  Сравнение по флагу о базовом исполнении.
             *  Базовое исполнение всегда записывается в начале раздела (предшествует другим записям)
             */
            protected int CompareByBaseMakeFlag(SpecDocumentGost2_113 doc1, SpecDocumentGost2_113 doc2)
            {
                // Базовое исполнение в начале
                int cmp = 0;
                if (doc1.IsBaseMake && !doc2.IsBaseMake)
                    cmp = -1;
                else if (!doc1.IsBaseMake && doc2.IsBaseMake)
                    cmp = 1;
                return cmp;
            }
        }

        internal class SpecDocumentGost2_113 : SpecDocumentGost2_106
        {
            bool _baseMake;             // базовое исполнение

            public SpecDocumentGost2_113(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr /*, SpecDocument parent*/)
                : base(/*tflexdocs,*/ doc, documentAttr /*, parent*/)
            {
            }

            protected override void ReadDocumentParameters(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr/*, SpecDocument parent*/)
            {
                //MessageBox.Show("Read1");
                base.ReadDocumentParameters(/*tflexdocs,*/ doc, documentAttr /*, parent*/);

                // _baseMake = ((DocsCategory)doc.Category) == DocsCategory.BaseMake;
            }

            public bool IsBaseMake
            {
                get { return _baseMake; }
            }

        }

        internal class SpecDocumentGost2_113EqualityComparer : SpecDocumentGost2_106EqualityComparer
        {
            public override bool Equals(SpecDocument x, SpecDocument y)
            {
                bool equals = base.Equals(x, y);
                if (!equals)
                    return false;
                else
                {
                    Debug.Assert(x is SpecDocumentGost2_113);
                    Debug.Assert(y is SpecDocumentGost2_113);

                    SpecDocumentGost2_113 docx = x as SpecDocumentGost2_113;
                    SpecDocumentGost2_113 docy = y as SpecDocumentGost2_113;

                    if (docx.IsBaseMake != docy.IsBaseMake)
                        return false;
                    else
                        return true;
                }
            }
        }

        public partial class ProgressForm : Form
        {
            public ProgressForm()
            {
                InitializeComponent();
            }

            public enum Stage
            {
                None,
                GettingData,
                ProcessingData,
                MakingReport
            };

            public void SetStepProgressBar(int step)
            {
                progressBar.Step = step;
                progressBar.PerformStep();
            }
            public void SetStage(Stage stage)
            {
                Font regular = new Font(gettingDataLabel.Font, System.Drawing.FontStyle.Regular);
                Font bold = new Font(gettingDataLabel.Font, System.Drawing.FontStyle.Bold);
                if (stage == Stage.GettingData)
                {
                    gettingDataLabel.Enabled = true;
                    gettingDataLabel.Font = bold;
                    SetStepProgressBar(10);
                }
                else
                {
                    gettingDataLabel.Enabled = false;
                    gettingDataLabel.Font = regular;
                }

                if (stage == Stage.ProcessingData)
                {
                    processingDataLabel.Enabled = true;
                    processingDataLabel.Font = bold;
                    SetStepProgressBar(10);
                }
                else
                {
                    processingDataLabel.Enabled = false;
                    processingDataLabel.Font = regular;
                }

                if (stage == Stage.MakingReport)
                {
                    makingReportLabel.Enabled = true;
                    makingReportLabel.Font = bold;
                    SetStepProgressBar(10);
                }
                else
                {
                    makingReportLabel.Enabled = false;
                    makingReportLabel.Font = regular;
                    SetStepProgressBar(20);
                }
                UpdateUI();
            }

            void UpdateUI()
            {
                Update();
                System.Windows.Forms.Application.DoEvents();
            }

            private void ProgressForm_FormClosing(object sender, FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                    e.Cancel = true;
            }
        }

        partial class ProgressForm
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
                this.stageGroupBox = new System.Windows.Forms.GroupBox();
                this.makingReportLabel = new System.Windows.Forms.Label();
                this.processingDataLabel = new System.Windows.Forms.Label();
                this.gettingDataLabel = new System.Windows.Forms.Label();
                this.progressBar = new System.Windows.Forms.ProgressBar();
                this.stageGroupBox.SuspendLayout();
                this.SuspendLayout();
                // 
                // stageGroupBox
                // 
                this.stageGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                            | System.Windows.Forms.AnchorStyles.Right)));
                this.stageGroupBox.Controls.Add(this.makingReportLabel);
                this.stageGroupBox.Controls.Add(this.processingDataLabel);
                this.stageGroupBox.Controls.Add(this.gettingDataLabel);
                this.stageGroupBox.Location = new System.Drawing.Point(12, 12);
                this.stageGroupBox.Name = "stageGroupBox";
                this.stageGroupBox.Size = new System.Drawing.Size(414, 100);
                this.stageGroupBox.TabIndex = 0;
                this.stageGroupBox.TabStop = false;
                this.stageGroupBox.Text = "Этап";
                // 
                // makingReportLabel
                // 
                this.makingReportLabel.AutoSize = true;
                this.makingReportLabel.Enabled = false;
                this.makingReportLabel.Location = new System.Drawing.Point(17, 72);
                this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                this.makingReportLabel.Name = "makingReportLabel";
                this.makingReportLabel.Size = new System.Drawing.Size(216, 13);
                this.makingReportLabel.TabIndex = 2;
                this.makingReportLabel.Text = "Формирование документа T-FLEX CAD...";
                // 
                // processingDataLabel
                // 
                this.processingDataLabel.AutoSize = true;
                this.processingDataLabel.Enabled = false;
                this.processingDataLabel.Location = new System.Drawing.Point(17, 50);
                this.processingDataLabel.Name = "processingDataLabel";
                this.processingDataLabel.Size = new System.Drawing.Size(111, 13);
                this.processingDataLabel.TabIndex = 1;
                this.processingDataLabel.Text = "Обработка данных...";
                // 
                // gettingDataLabel
                // 
                this.gettingDataLabel.AutoSize = true;
                this.gettingDataLabel.Enabled = false;
                this.gettingDataLabel.Location = new System.Drawing.Point(17, 27);
                this.gettingDataLabel.Name = "gettingDataLabel";
                this.gettingDataLabel.Size = new System.Drawing.Size(110, 13);
                this.gettingDataLabel.TabIndex = 0;
                this.gettingDataLabel.Text = "Получение данных...";
                // 
                // progressBar
                // 
                this.progressBar.Location = new System.Drawing.Point(12, 124);
                this.progressBar.Name = "progressBar";
                this.progressBar.Size = new System.Drawing.Size(414, 23);

                this.progressBar.TabIndex = 0;
                // 
                // ProgressForm
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(438, 165);
                this.Controls.Add(this.progressBar);
                this.Controls.Add(this.stageGroupBox);
                this.Name = "ProgressForm";
                this.Text = "Формирование спецификации";
                this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProgressForm_FormClosing);
                this.stageGroupBox.ResumeLayout(false);
                this.stageGroupBox.PerformLayout();
                this.ResumeLayout(false);

            }

            #endregion

            private System.Windows.Forms.GroupBox stageGroupBox;
            private System.Windows.Forms.Label gettingDataLabel;
            private System.Windows.Forms.Label makingReportLabel;
            private System.Windows.Forms.Label processingDataLabel;
            private System.Windows.Forms.ProgressBar progressBar;
        }

        public class CadReportVariantA : IDisposable
        {
            static bool _calledFromDocs;


            public static bool CalledFromDocs
            {
                get { return _calledFromDocs; }
            }

            static CadReportVariantA()
            {
                _calledFromDocs = false;
            }

            /**
             *  Точка входа при создании отчета из DOCs
             */
            public static void Make(IReportGenerationContext context, DocumentAttributes documentAttr)
            {
                //Form1Namespace.Form1.NewPagesForm.Dispose();
                _calledFromDocs = true;

                try
                {
                    using (CadReportVariantA report = new CadReportVariantA())
                    {
                        int docID = report.GetDocsDocumentID(context);
                        if (docID == -1)
                        {
                            MessageBox.Show("Выберите объект для формирования отчета");
                            return;
                        }
                        report.Make(docID, context, documentAttr);

                    }
                }
                catch (Exception e)
                {
                    string message = String.Format("ОШИБКА: {0}", e.ToString());
                    System.Windows.Forms.MessageBox.Show(message, "Ошибка",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                // progressForm.Dispose();
            }

            /**
             *  Получение идентификатора документа T-FLEX DOCs
             */
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

            /**
             *  Конструктор
             */
            public CadReportVariantA()
            {
            }

            // Инициализация
            /* void Initialize()
            {
            }*/

            ~CadReportVariantA()
            {
            }

            public ProgressForm progressForm = new ProgressForm();

            public void Dispose()
            {
            }

            // Create report
            public void Make(int documentID, IReportGenerationContext context, DocumentAttributes documentAttr)
            {
                TFDDocument document = new TFDDocument(documentID); //_tflexdocs.GetDocument(documentID);
                Make(document, context, documentAttr);
            }

            // Create report
            public void Make(TFDDocument document, IReportGenerationContext context, DocumentAttributes documentAttr)
            {
                try
                {
                    progressForm.Show();

                    // Чтение данных из DOCs
                    progressForm.SetStage(ProgressForm.Stage.GettingData);
                    ProductsGroup products = new ProductsGroup(/*_tflexdocs*/ documentAttr);
                    products.ReadContent(document, documentAttr, progressForm);


                    // Подготовка данных для CAD
                    progressForm.SetStage(ProgressForm.Stage.ProcessingData);
                    // DocumentAttributes attrs = products.GetAttributes(Settings.AttributesObjectName);
                    //  string hashString = DocumentHashAlgorithm.ConvertToHexString(products.CalcualteHashAlgorithm2());
                    SpecForm1 specForm = new SpecForm1(products.SpecDocument, /*products.PrimaryUsage*/document.FirstUse, documentAttr /*, hashString*/);

                 //   MessageBox.Show("specForm " + string.Join("\r\n", specForm..Select(t => t.Name + " " + t.Designation)));

                    products.AddSpecRowsForm1(specForm, documentAttr.FullNames, documentAttr, document.Naimenovanie, document.Denotation);

                    // Формирование документа CAD
                    progressForm.SetStage(ProgressForm.Stage.MakingReport);
                    //                    APILoader.Initialize();      //Инициализация API CAD
                    //                    context.CopyTemplateFile();  //Создаем копию шаблона
                    TFlex.Model.Document cadDocument = GetCadDocument(context);
                    FillCadDocument(specForm, cadDocument);

                    // Сохранение
                    //SaveCadDocument(cadDocument);
                    //cadDocument.Save();

                    progressForm.SetStage(ProgressForm.Stage.None);
                    // TFlex.Application.ExitSession();
                    //progressForm.Close();
                    progressForm.Dispose();

                    System.Diagnostics.Process.Start(context.ReportFilePath);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.ToString(), "Ошибка",
                                                         System.Windows.Forms.MessageBoxButtons.OK,
                                                         System.Windows.Forms.MessageBoxIcon.Error);
                }
            }


            TFlex.Model.Document GetCadDocument(IReportGenerationContext context)
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
                    return null;
                }
                TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD              
                context.CopyTemplateFile();  //Создаем копию шаблона               
                return TFlex.Application.OpenDocument(context.ReportFilePath);
            }

            void SaveCadDocument(TFlex.Model.Document cadDocument)
            {
                bool saved = cadDocument.Save();
                if (!saved)
                {
                    string message = String.Format("Ошибка сохранения файла отчета");
                    throw new Exception(message);
                }
            }

            void FillCadDocument(SpecForm1 specForm, TFlex.Model.Document cadDocument)
            {
                cadDocument.BeginChanges("Формирование отчета");
                ParagraphText textTable = GetSpecificationTextTable(cadDocument);

                // Начало редактирования чертежа
                textTable.BeginEdit();

                // Заполнение таблицы
                Table table = textTable.GetTableByIndex(0);
                SpecRowForm1[] records = specForm.GetRows();
                foreach (SpecRowForm1 record in records)
                {
                    InsertRow(textTable, table, record);
                }

                // Добавление листа регистрации изменений
                string changeLogPageName = null;
                if (specForm.Attributes.AddChangelogPage)
                    changeLogPageName = AddTitulOrChangeLogPage(specForm, cadDocument);

                // Добавление титульного листа
                string titulPageName = null;
                if (specForm.Attributes.AddTitulList)
                    titulPageName = AddTitulOrChangeLogPage(specForm, cadDocument);

                // Завершение редактирования чертежа
                textTable.EndEdit();

                // Добавление "форматок"
                AddPageForms(specForm, cadDocument, changeLogPageName, titulPageName);

                cadDocument.EndChanges();

                cadDocument.Save();
                cadDocument.Close();
            }

            ParagraphText GetSpecificationTextTable(TFlex.Model.Document document)
            {
                ParagraphText textTable = null;

                foreach (Text text in document.Texts)
                {
                    if (text == null)
                        continue;

                    if (text.Name == "$TextTablePg1")
                        textTable = text as ParagraphText;

                    if (textTable != null)
                        break;
                }

                if (textTable == null)
                    throw new ApplicationException(
                        "На чертеже не найден объект $TextTablePg1");

                return textTable;
            }

            void InsertRow(ParagraphText parText, Table table, SpecRowForm1 row)
            {
                table.InsertRows(1, (uint)(table.CellCount - 1), Table.InsertProperties.After);
                uint newCellsCount = (uint)table.CellCount;
                uint columnsCount = (uint)table.ColumnCount;

                InsertTableText(table, newCellsCount - columnsCount, 0, parText, row.Format);
                InsertTableText(table, newCellsCount - columnsCount + 1, 0, parText, row.Zone);
                InsertTableText(table, newCellsCount - columnsCount + 2, 0, parText, row.Position);
                InsertTableText(table, newCellsCount - columnsCount + 3, 0, parText, row.Designation);
                string name = row.Name.Text.Replace('*', '\x00d7'); // Замена звездочки на крестик
                SpecFormCell nameCell = new SpecFormCell(name, row.Name.Underlining, row.Name.Align);
                InsertTableText(table, newCellsCount - columnsCount + 4, 0, parText, nameCell);
                InsertTableText(table, newCellsCount - columnsCount + 5, 0, parText, row.Count);

                CharFormat noteCharFormat = parText.CharacterFormat;
                //***
                noteCharFormat.FontName = "T-FLEX Type B";
                table.InsertText(newCellsCount - columnsCount + 6, 0, row.Note.Text, noteCharFormat);

                if (row.NoteSuperscriptSymbol != null)
                {
                    CharFormat cf = parText.CharacterFormat;
                    cf.VertOffset = 1.5;
                    cf.FontSize = 2.5;
                    table.InsertText(newCellsCount - columnsCount + 7, 0, row.NoteSuperscriptSymbol, cf);
                }
            }

            void InsertTableText(Table table, uint cell, uint pos, ParagraphText parText, SpecFormCell text)
            {
                InsertTableText(table, cell, pos, parText, text, null);
            }

            void InsertTableText(Table table, uint cell, uint pos, ParagraphText parText, SpecFormCell text,
                string superscriptNote)
            {
                CharFormat ch = parText.CharacterFormat;
                if (text.Underlining == SpecFormCell.UnderliningFormat.Underlined)
                    ch.isUnderline = true;

                table.SetSelection(cell, cell);

                ParaFormat pf = parText.ParagraphFormat;
                if (text.Align == SpecFormCell.AlignFormat.Left)
                {
                    pf.HorJustification = ParaFormat.Just.Left;
                }
                if (text.Align == SpecFormCell.AlignFormat.Center)
                {
                    pf.HorJustification = ParaFormat.Just.Center;
                }
                else if (text.Align == SpecFormCell.AlignFormat.Right)
                {
                    pf.HorJustification = ParaFormat.Just.Right;
                }
                parText.ParagraphFormat = pf;

                table.InsertText(cell, 0, text.Text, ch);
            }

            void FormatParagraphText(ParagraphText text, SpecFormCell cell)
            {
                CharFormat ch = text.CharacterFormat;
                ParaFormat pf = text.ParagraphFormat;

                if (cell.Underlining == SpecFormCell.UnderliningFormat.NotUnderlined)
                    ch.isUnderline = false;
                else if (cell.Underlining == SpecFormCell.UnderliningFormat.Underlined)
                    ch.isUnderline = true;

                if (cell.Align == SpecFormCell.AlignFormat.Left)
                    pf.HorJustification = ParaFormat.Just.Left;
                else if (cell.Align == SpecFormCell.AlignFormat.Center)
                    pf.HorJustification = ParaFormat.Just.Center;
                else if (cell.Align == SpecFormCell.AlignFormat.Right)
                    pf.HorJustification = ParaFormat.Just.Right;

                text.ParagraphFormat = pf;
                text.CharacterFormat = ch;
            }

            /// Добавление листа регистрации изменений или Титульного листа
            /// @return
            ///     Имя страницы с листом регистрации или титульным листом
            string AddTitulOrChangeLogPage(SpecForm1 specForm, TFlex.Model.Document cadDocument)
            {
                Page page = new Page(cadDocument);
                return page.Name;
            }

            /// Добавление форм в виде фгарментов. Заполнение переменных фрагментов.
            void AddPageForms(SpecForm1 specForm, TFlex.Model.Document cadDocument, string changeLogPageName, string titulPageName)
            {
                FileReference fileReference = new FileReference();

                FileLink titulFormLink = null;
                FileLink firstPageFormLink = null;
                FileLink otherPageFormLink = null;
                FileLink regListFormLink = null;

                int pageNumber = specForm.Attributes.FirstPageNumber - 1;

                foreach (Page page in cadDocument.Pages)
                {
                    ++pageNumber;
                    FileLink link = null;
                    if (page.PageType != PageType.Normal)
                        continue;

                    //Вставка титульного листа
                    if (page.Name == titulPageName)
                    {
                        if (titulFormLink == null)
                        {
                            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.TitulGost2_105Name);
                            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                            titulFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                            /* titulFormLink = new FileLink(cadDocument,
                                fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.TitulGost2_105Name).LocalPath,
                                true); */
                            /*titulFormLink = new FileLink(cadDocument,
                                Settings.PathToTemplates + @"\" + Settings.TitulGost2_105Name, true); */
                        }
                        link = titulFormLink;
                    }
                    else if (page.Name == "Страница 1")
                    {
                        // вставка формы 1 в виде фрагмента
                        if (firstPageFormLink == null)
                        {
                            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form1Gost2_106Name);
                            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                            firstPageFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                            /*firstPageFormLink = new FileLink(cadDocument,
                                fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form1Gost2_106Name).LocalPath,
                                true); */
                            /*firstPageFormLink = new FileLink(cadDocument,
                                Settings.PathToTemplates + @"\" + Settings.Form1Gost2_106Name,
                                true); */
                        }
                        link = firstPageFormLink;
                    }
                    else if (page.Name == changeLogPageName)
                    {
                        // вставка листа регистрации изменений
                        if (regListFormLink == null)
                        {
                            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ChangelogGost2_503Name);
                            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                            regListFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                            /*regListFormLink = new FileLink(cadDocument,
                                fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ChangelogGost2_503Name).LocalPath,
                                true); */
                            /*regListFormLink = new FileLink(cadDocument,
                                Settings.PathToTemplates + @"\" + Settings.ChangelogGost2_503Name,
                                true);*/
                        }
                        link = regListFormLink;
                    }
                    else if (page.Name.StartsWith("Страница"))
                    {
                        // вставка формы 1а в виде фрагмента
                        if (otherPageFormLink == null)
                        {
                            var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form1aGost2_106Name);
                            fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                            otherPageFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                            /* otherPageFormLink = new FileLink(cadDocument,
                                fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.Form1aGost2_106Name).LocalPath,
                                true); */
                            /* otherPageFormLink = new FileLink(cadDocument,
                                Settings.PathToTemplates + @"\" + Settings.Form1aGost2_106Name,
                                true); */
                        }
                        link = otherPageFormLink;
                    }
                    else
                    {
                        throw new ApplicationException("Неизвестное наименование страницы документа CAD");
                    }

                    page.Rectangle = new Rectangle(0, 0, 210, 297);

                    Fragment formFragment = new Fragment(link);
                    formFragment.Page = page;

                    // задание значений переменных фрагмента
                    foreach (FragmentVariableValue var in formFragment.VariableValues)
                    {
                        switch (var.Name)
                        {
                            case "$goev":
                                var.TextValue = specForm.Attributes.Goev;
                                break;
                            case "$vpmo":
                                var.TextValue = specForm.Attributes.MilitaryChief;
                                break;
                            case "$vpmochief":
                                if (specForm.Attributes.checkChiefMillitary)
                                    var.TextValue = "СОГЛАСОВАНО\n\nНачальник 454 ВП МО РФ";
                                else var.TextValue = "";
                                break;
                            case "$vpmoVisa":
                                if (specForm.Attributes.checkVPMO)
                                    var.TextValue = "ВП МО РФ";
                                else var.TextValue = "";
                                break;
                            case "$vpmochiefDate":
                                var.TextValue = "";
                                /* if (specForm.Attributes.checkChiefMillitary)
                                     var.TextValue = "";
                                 else var.TextValue = "";*/
                                break;
                            case "$technolog":
                                var.TextValue = specForm.Attributes.Tech;
                                break;
                            case "$otd45":
                                var.TextValue = specForm.Attributes.Otdel45;
                                break;
                            case "$ncontr":
                                var.TextValue = specForm.Attributes.Norm;
                                break;
                            case "$nachbrig":
                                var.TextValue = specForm.Attributes.ChiefTeam;
                                break;
                            case "$year":
                                var.TextValue = specForm.Attributes.YearMake;
                                break;
                            case "$oboznach":
                                var.TextValue = specForm.Designation;
                                break;
                            case "$naimen1":
                                var.TextValue = specForm.Name;
                                break;
                            case "$list":
                                string listVar;
                                if (cadDocument.Pages.Count == 1)
                                    listVar = "-";
                                else
                                    listVar = pageNumber.ToString();
                                var.TextValue = listVar;
                                break;
                            case "$listov":
                                long total = cadDocument.Pages.Count + specForm.Attributes.FirstPageNumber;
                                if (specForm.Attributes.AddTitulList)
                                    var.TextValue = (total - 2).ToString();
                                else var.TextValue = (total - 1).ToString();
                                break;
                            case "$per_prim":
                                var.TextValue = specForm.FirstUsedProductDesignation;
                                break;
                            case "$razrab":
                                var.TextValue = specForm.Attributes.DesignedBy;
                                break;
                            case "$prover":
                                var.TextValue = specForm.Attributes.ControlledBy;
                                break;
                            case "$n_contr":
                                var.TextValue = specForm.Attributes.NControlBy;
                                break;
                            case "$utverd":
                                var.TextValue = specForm.Attributes.ApprovedBy;
                                break;
                            case "$sidebar_text":
                                var.TextValue = GetSidebarText(specForm);
                                break;
                            case "$product_guid":
                                var.TextValue = specForm.DocumentGuid;
                                break;
                            case "$created_date":
                                DateTime now = DateTime.Now;
                                var.TextValue = now.ToString(Settings.DateTimeFormat);
                                break;
                            case "$hash":
                                var.TextValue = string.Empty; //specForm.Hash;
                                break;
                            case "$lit1":
                                string lit1 = "";
                                if (specForm.Litera != null && specForm.Litera.Length > 0)
                                    lit1 = specForm.Litera[0];
                                var.TextValue = lit1;
                                break;
                            case "$lit2":
                                string lit2 = "";
                                if (specForm.Litera != null && specForm.Litera.Length > 1)
                                    lit2 = specForm.Litera[1];
                                var.TextValue = lit2;
                                break;
                            case "$lit3":
                                string lit3 = "";
                                if (specForm.Litera != null && specForm.Litera.Length > 2)
                                    lit3 = specForm.Litera[2];
                                var.TextValue = lit3;
                                break;
                            case "$izm2":
                                if (specForm.Attributes.AddModify)
                                {
                                    foreach (int nPage in specForm.Attributes.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.NumberModify;

                                    foreach (int nPage in specForm.Attributes.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.NumberModify;
                                }
                                break;
                            case "$il2":
                                if (specForm.Attributes.AddModify)
                                {
                                    foreach (int nPage in specForm.Attributes.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.ListZam;

                                    foreach (int nPage in specForm.Attributes.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.ListNov;
                                }
                                break;
                            case "$ido2":
                                if (specForm.Attributes.AddModify)
                                {
                                    foreach (int nPage in specForm.Attributes.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.NumberDocument.ToString();

                                    foreach (int nPage in specForm.Attributes.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.NumberDocument.ToString();
                                }
                                break;
                            case "$idata2":
                                if (specForm.Attributes.AddModify)
                                {
                                    foreach (int nPage in specForm.Attributes.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.DateModify;

                                    foreach (int nPage in specForm.Attributes.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = specForm.Attributes.DateModify;
                                }
                                break;
                        }
                    }
                }

            } // AddPageForms

            string GetSidebarText(SpecForm1 specForm)
            {
                DocumentAttributes attrs = specForm.Attributes;
                string text = String.Format("{0}               {1}     {2}               {3}     {4}               {5}",
                    attrs.SidebarSignHeader1, attrs.SidebarSignName1,
                    attrs.SidebarSignHeader2, attrs.SidebarSignName2,
                    attrs.SidebarSignHeader3, attrs.SidebarSignName3);
                return text.Trim();
            }
        }

        /*
         *  Настройки
         */
        internal sealed class Settings
        {
            public static readonly string PathToTemplates = @"Служебные файлы\Шаблоны отчётов\Спецификация ГОСТ 2.113 (ЕСКД)\";
            //    = @"D:\Шаблоны отчетов\";
            //    = @"D:\Тест - отчет";//Environment.GetEnvironmentVariable("ALLUSERSPROFILE") + @"\Application Data\T-FLEX DOCs 10\Prototypes\";
            // @"S:\Projects\GlobusTFlexSuite\devel\Source\DocsReports\SpecificationReport\Templates";
            // public static readonly string PathToTemplates = @"\\s1.domain.globus\ReportTemplates";

            public static readonly string TitulGost2_105Name = "Титульный лист ГОСТ 2.105-95.grb";

            public static readonly string Form1Gost2_106Name = "Спецификация форма 1 ГОСТ 2.106-96.grb";

            public static readonly string Form1aGost2_106Name = "Спецификация форма 1а ГОСТ 2.106-96.grb";

            public static readonly string ChangelogGost2_503Name = "Лист регистрации изменений ГОСТ 2.503.grb";

            public static readonly string DateTimeFormat = "d.MM.yyyy H:mm:ss";

            public static readonly string AttributesObjectName = "Спецификация вариант А ГОСТ 2.113";
        }
    }

    namespace Globus.TFlexDocs.SpecificationReport
    {
        using System;
        using System.Collections;
        using System.Collections.Generic;
        using System.Diagnostics;
        using System.Globalization;
        using System.Text;
        using System.Text.RegularExpressions;

        using Globus.TFlexDocs;
        using System.Runtime.InteropServices;
        using System.Windows.Forms;
        using System.Threading;

        /**
         *  Документ спецификации.
         */
        public class SpecDocument
        {
            // Параметры объекта (документа) T-FLEX DOCs

            int _id;                // ID объекта
            // string _guid;           // GUID объекта
            string _docInfoString;  // Строка для метода ToString()

            // Основные параметры документа

            //  SpecDocument _parent;   // изделие, в состав которого входит данное изделие
            string _designation;    // обозначение
            string _name;           // наименование документа

            int? _position;           // наименование документа

            string _bomSection;     // раздел спецификации
            // string[] _litera;       // литера
            string _note;           // примечание

            int _IDSection;   // ID разднла спецификации  

            int _emptyLinesBefore;
            int _emptyLinesAfter;
            ReferenceObject _replacingObject;



            bool _addDocMake = false; // флаг добавления заготовки из постоянной части в переменную

            public bool AddDocMake
            {
                get { return _addDocMake; }
                set { _addDocMake = value; }
            }

            /**
             *  Constructor
             */


            public SpecDocument(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr /*, SpecDocument parent*/)
            {
                //_parent = parent;
                // MessageBox.Show("SpecDocument parent"+_parent.Name+" "+_parent.Designation);
                ReadDocumentParameters(/*tflexdocs,*/ doc, documentAttr /*, parent*/);
                //MessageBox.Show(doc.Naimenovanie + " parent " + parent.Name);
            }

            /**
             *  Чтение параметров документа TFlexDocs
             */
            protected virtual void ReadDocumentParameters(/*TFlexDocs tflexdocs,*/ TFDDocument doc, DocumentAttributes documentAttr /*, SpecDocument parent*/)
            {
                //MessageBox.Show("Last Read");
                //if (tflexdocs == null)
                //    throw new ArgumentException("Parameter is null", "tflexdocs");
                if (doc == null)
                    throw new ArgumentException("Parameter is null", "doc");

                //tflexdocs.LoadDocumentParameters(doc, LOAD_PARAMETERS.ALL);

                //_id = Math.Abs(doc.id);
                // _guid = doc.GUID;
                // TFlexDocs tflexdocs = new TFlexDocs();
                _docInfoString = String.Format(@"[{0}] {1} {2}", doc.ObjectID.ToString(), doc.Naimenovanie, doc.Denotation);//tflexdocs.GetDocumentInfoString(doc);
                // отбрасываем пробел и символы в круглых скобках в конце обозначения
                string docsDesignation = doc.Denotation.Trim();//tflexdocs.GetParameterValue(doc, ParametersGUIDs.Designation);
                Match match = Regex.Match(docsDesignation, "^(?<p1>.*?)(?: \\(.*\\))?$");
                _designation = match.Groups["p1"].Value;

                _name = doc.Naimenovanie;//tflexdocs.GetParameterValue(doc, ParametersGUIDs.Name);
                _replacingObject = ReplacingObject;
                // Раздел спецификации
                _bomSection = doc.BomSection;//ReadBomGroup(tflexdocs, doc);

                // Литера
                // _litera = ReadDocumentLitera(tflexdocs, doc);

                // Примечание
                _note = doc.Remarks;//tflexdocs.GetParameterValue(doc, ParametersGUIDs.BomNote);

                // tflexdocs.GetBomEmptyLineCounts(doc.ObjectID, out _emptyLinesBefore, out _emptyLinesAfter);
                // _emptyLinesBefore = 0;
                // _emptyLinesAfter = 0;
                _emptyLinesBefore = documentAttr.CountEmptyLinesBefore;
                _emptyLinesAfter = documentAttr.CountEmptyLinesAfter;
                // MessageBox.Show(_docInfoString);
            }

            // Рег. выражение для выделения значимого элемента строки (не
            // пробельный символ и не запятая)
            protected static readonly string SignificantCharsPattern = @"[^\s,]+";

            static readonly Regex _significantCharsRegex = new Regex(SignificantCharsPattern, RegexOptions.Compiled);

            /**
             *  Получение массива значимых элементов строки.
             *  Каждый элемент определяется рег. выражением.
             */
            protected string[] GetSignificantSubstrings(string str, Regex item)
            {
                List<string> items = new List<string>();

                if (str != null)
                {
                    MatchCollection mc = item.Matches(str);
                    foreach (Match m in mc)
                        items.Add(m.Value);
                }

                return items.ToArray();
            }

            /**
             *  Получение массива значимых элементов строки.
             *  Элементы строки определяются шаблоном SignificantCharsPattern.
             */
            protected string[] GetSignificantSubstrings(string str)
            {
                return GetSignificantSubstrings(str, _significantCharsRegex);
            }

            /**
             *  Чтение литеры документа
             */
            /* string[] ReadDocumentLitera(TFlexDocs tflexdocs, TFDDocument doc)
            {
                string litera = tflexdocs.GetParameterValue(doc, ParametersGUIDs.TechLitera);
                return GetSignificantSubstrings(litera, _significantCharsRegex);
            } */

            public override string ToString()
            {
                return _docInfoString;
            }

            /* public override int GetHashCode()
             {
                 return _guid.GetHashCode();
             }*/

            public int ID
            {
                get { return _id; }
                set { _id = value; }
            }

            public int? Position
            {
                get { return _position; }
                set { _position = value; }
            }

            public ReferenceObject ReplacingObject
            {
                get { return _replacingObject; }
            }

            /* public string GUID
            {
                get { return _guid; }
            } */

            public string Designation
            {
                get { return _designation; }
            }

            public string Name
            {
                get { return _name; }
                set { _name = value; }
            }

            /*   protected SpecDocument Parent
               {
                   get { return _parent; }
               }*/

            public string BomSection
            {
                get { return _bomSection; }
                set { _bomSection = value; }
            }

            /* public string[] Litera
             {
                 get { return _litera; }
             }*/

            public string Note
            {
                get { return _note; }
                set { _name = value; }
            }

            public int EmptyLinesBefore
            {
                get { return _emptyLinesBefore; }
            }

            public int EmptyLinesAfter
            {
                get { return _emptyLinesAfter; }
            }

            public int IDSection
            {
                set { _IDSection = IDSection; }
                get { return _IDSection; }
            }
        }

        /// Компаратор аутентичности двух документов спецификации.
        /// Документы считаются аутентичными, если у них совпадают параметры,
        /// на основе которых делается запись в спецификации.
        public class SpecDocumentBaseEqualityComparer : IEqualityComparer<SpecDocument>
        {
            // Сравнение документов по обозначению и наименованию
            public virtual bool Equals(SpecDocument x, SpecDocument y)
            {
                Debug.Assert(x != null);
                Debug.Assert(y != null);

                if (x.Designation != y.Designation)
                    return false;
                else if (x.Name != y.Name)
                    return false;
                else
                    return true;
            }

            public virtual int GetHashCode(SpecDocument document)
            {
                return document.GetHashCode();
            }

            /// Сравнение массивов двух строк. Массивы строк равны, если они содержат
            /// равное количество элементов и каждый элемент одного массива
            /// равен элементу второго массива с тем же индексом
            protected bool StringArraysEquals(string[] a1, string[] a2)
            {
                if (a1.Length != a2.Length)
                    return false;
                else
                {
                    for (int i = 0; i < a1.Length; ++i)
                        if (a1[i] != a2[i])
                            return false;
                }
                return true;
            }
        }

        /**
         *  Компаратор документов спецификации
         */
        public abstract class SpecDocumentComparer : IComparer<SpecDocument>
        {
            /**
             *  Сравнение двух документов
             */
            public abstract int Compare(SpecDocument doc1, SpecDocument doc2);

            /**
             *  Рег. выражения для выделения элементов строки.
             *  Элементом строки может быть:
             *  - последовательность буквенных символов;
             *  - последовательность цифровых символов или символа точка;
             */

            static readonly Regex _elementRegex = new Regex(
                "(?<alpha>[A-Za-zА-Яа-я]+)"             // последовательность букв
                + "|(?<number>[0-9]+(?:,[0-9]+)?)"      // число
            );

            /**
             * Получение массива элементов строки.
             */
            protected object[] Slice(string s)
            {
                if (s == null)
                    return new object[] { };

                ArrayList elements = new ArrayList();
                MatchCollection mc = _elementRegex.Matches(s);

                foreach (Match m in mc)
                {
                    Group alphaGroup = m.Groups["alpha"];
                    if (alphaGroup.Success)
                    {
                        elements.Add(alphaGroup.Value);
                        continue;
                    }
                    Group numberGroup = m.Groups["number"];
                    if (numberGroup.Success)
                    {

                        var nom = m.Value.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);

                        double element = Convert.ToDouble(nom);
                        elements.Add(element);
                        continue;
                    }
                }
                return elements.ToArray();
            }


            /**
             *  Множители кратных единиц измерения
             */
            static readonly Dictionary<string, double> _unitMultiplies;

            static SpecDocumentComparer()
            {
                _unitMultiplies = new Dictionary<string, double>();
                // Метры
                _unitMultiplies.Add("мкм", 1e-6);
                _unitMultiplies.Add("мм", 1e-3);
                _unitMultiplies.Add("см", 1e-2);
                _unitMultiplies.Add("м", 1);
                _unitMultiplies.Add("км", 1e3);
                // Фарады
                _unitMultiplies.Add("пФ", 1e-12);
                _unitMultiplies.Add("нФ", 1e-9);
                _unitMultiplies.Add("мкФ", 1e-6);
                _unitMultiplies.Add("мФ", 1e-3);
                _unitMultiplies.Add("Ф", 1);
                // Омы
                _unitMultiplies.Add("мкОм", 1e-6);
                _unitMultiplies.Add("мОм", 1e-3);
                _unitMultiplies.Add("Ом", 1);
                _unitMultiplies.Add("кОм", 1e3);
                _unitMultiplies.Add("МОм", 1e6);
                _unitMultiplies.Add("ГОм", 1e9);
                // Генри
                _unitMultiplies.Add("пГн", 1e-12);
                _unitMultiplies.Add("нГн", 1e-9);
                _unitMultiplies.Add("мкГн", 1e-6);
                _unitMultiplies.Add("мГн", 1e-3);
                _unitMultiplies.Add("Гн", 1);
                _unitMultiplies.Add("кГн", 1e3);
                // Вольты
                _unitMultiplies.Add("мкВ", 1e-6);
                _unitMultiplies.Add("мВ", 1e-3);
                _unitMultiplies.Add("В", 1);
                _unitMultiplies.Add("кВ", 1e3);
                _unitMultiplies.Add("МВ", 1e6);
                _unitMultiplies.Add("ГВ", 1e9);
                // Амперы
                _unitMultiplies.Add("мкА", 1e-6);
                _unitMultiplies.Add("мА", 1e-3);
                _unitMultiplies.Add("А", 1);
                _unitMultiplies.Add("кА", 1e3);
                // Ватты
                _unitMultiplies.Add("мкВт", 1e-6);
                _unitMultiplies.Add("мВт", 1e-3);
                _unitMultiplies.Add("Вт", 1);
                _unitMultiplies.Add("кВт", 1e3);
                _unitMultiplies.Add("МВт", 1e6);
                _unitMultiplies.Add("ГВт", 1e9);
            }

            /**
             *  Умножение числа на соответствующий множитель единицы измерения.
             * 
             *  @param ref double number
             *      Число. Параметр модифицируется.
             *  @param string unit
             *      Единица измерения
             *  @return
             *      true, если единица измерения определена и число модифицировано.
             */
            bool MultiplyByUnitFactor(ref double number, string unit)
            {
                if (_unitMultiplies.ContainsKey(unit))
                {
                    number = _unitMultiplies[unit] * number;
                    return true;
                }
                else
                    return false;
            }

            protected int CompareByStringElements(string s1, string s2)
            {
                return CompareByStringElements(s1, s2, false);
            }

            protected int CompareByStringElements(string s1, string s2, bool withUnits)
            {
                if (s1 == s2)
                    return 0;

                object[] elements1 = Slice(s1);
                object[] elements2 = Slice(s2);

                IEnumerator e1 = elements1.GetEnumerator();
                bool next1 = e1.MoveNext();
                IEnumerator e2 = elements2.GetEnumerator();
                bool next2 = e2.MoveNext();
                while (true)
                {
                    if (!next1 && next2)
                        // s1 имеет меньше элементов, чем s2
                        return -1;
                    else if (next1 && !next2)
                        // s1 имеет больше элементов, чем s2
                        return 1;
                    else if (!next1 && !next2)
                        // s1 и s2 имеют равное количество элементов
                        return 0;
                    else
                    {
                        // сравнение очередной пары элементов
                        object element1 = e1.Current;
                        object element2 = e2.Current;
                        if (element1 is string && element2 is string)
                        {
                            // Сравнение двух строк
                            int c = String.Compare((string)element1, (string)element2);
                            if (c != 0)
                                return c;
                        }
                        else if (element1 is double && element2 is string)
                        {
                            return 1;
                        }
                        else if (element1 is string && element2 is double)
                        {
                            return -1;
                        }
                        else if (element1 is double && element2 is double)
                        {
                            // Сравнение двух чисел
                            double num1 = (double)element1;
                            double num2 = (double)element2;
                            if (!withUnits)
                            {
                                // Сравнение без учета возможных единиц измерения
                                if (num1 < num2)
                                    return -1;
                                else if (num1 > num2)
                                    return 1;
                            }
                            else
                            {
                                // Сравнение с учетом указания единицы измерения в следующем элементе
                                next1 = e1.MoveNext();
                                if (next1 && e1.Current is string)
                                {
                                    bool multiplied = MultiplyByUnitFactor(ref num1, (string)e1.Current);
                                    if (multiplied)
                                        next1 = e1.MoveNext();
                                }
                                next2 = e2.MoveNext();
                                if (next2 && e2.Current is string)
                                {
                                    bool multiplied = MultiplyByUnitFactor(ref num2, (string)e2.Current);
                                    if (multiplied)
                                        next2 = e2.MoveNext();
                                }
                                if (num1 < num2)
                                    return -1;
                                else if (num1 > num2)
                                    return 1;
                                else
                                    // в начало цикла, следующие элементы строк уже прочитаны
                                    continue;
                            }
                        }
                        else
                            Debug.Assert(false);
                    }
                    next1 = e1.MoveNext();
                    next2 = e2.MoveNext();
                }
            }
        }

        /**
        *  Коллекция документов спецификации
        */
        public class SpecDocumentCollection : List<SpecDocument>
        {
            /**
             * Проверка, содержит ли коллекция документ document
             * 
             * @param document
             *      Документ, наличие которого интересует
             * 
             * @param ecomparer
             *      Компаратор, определяющий аутентичность двух документов
             */
            public bool Contains(SpecDocument document, IEqualityComparer<SpecDocument> ecomparer)
            {
                foreach (SpecDocument thisDoc in this)
                {
                    if (ecomparer.Equals(document, thisDoc))
                        return true;
                }
                return false;
            }


            /**
             * Проверка аутентичности коллекции другой коллекции. Коллекции аутентичны, если
             * для каждого документа коллекции this существует аутентичный документ
             * коллекции other, и наоборот. Аутентичность документов определяется компаратором.
             * 
             * @param other
             *      Коллекция, на аутентичность которой проверяется коллекция this
             * 
             * @param ecomparer
             *      Компаратор, определяющий аутентичность двух документов
             */
            public bool EqualsContent(SpecDocumentCollection other, IEqualityComparer<SpecDocument> ecomparer)
            {
                if (Count != other.Count)
                    return false;

                foreach (SpecDocument thisDoc in this)
                {
                    if (!other.Contains(thisDoc, ecomparer))
                        return false;
                }

                foreach (SpecDocument otherDoc in other)
                {
                    if (!this.Contains(otherDoc, ecomparer))
                        return false;
                }

                return true;
            }
        }

        /**
         *  Ячейка формы спецификации
         */
        public class SpecFormCell
        {
            string _text;
            UnderliningFormat _underlining;
            AlignFormat _align;

            /**
             *  Параметр подчеркивания
             */
            public enum UnderliningFormat
            {
                NotUnderlined,
                Underlined
            }

            /**
             *  Параметр выравнивания
             */
            public enum AlignFormat
            {
                Left,
                Center,
                Right
            }

            public SpecFormCell(string text, UnderliningFormat underlining, AlignFormat align)
            {
                _text = text;
                _underlining = underlining;
                _align = align;
            }

            public SpecFormCell(string text, AlignFormat align)
                : this(text, UnderliningFormat.NotUnderlined, align)
            {
            }

            public SpecFormCell(string text)
                : this(text, UnderliningFormat.NotUnderlined, AlignFormat.Left)
            {
            }

            public override string ToString()
            {
                return _text;
            }

            public string Text
            {
                get
                {
                    return _text;
                }

                set
                {
                    _text = value;
                }
            }

            public UnderliningFormat Underlining
            {
                get { return _underlining; }
            }

            public AlignFormat Align
            {
                get { return _align; }
            }

            public static SpecFormCell Empty
            {
                get
                {
                    return new SpecFormCell("");
                }
            }
        }

        public class TextFormatter
        {
            Font _font;

            int _displayResolutionX;

            public TextFormatter(Font font)
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
                double width = 25.4f * widthFactor * (double)size.Width / (double)_displayResolutionX;
                return width;
            }

            public string[] WrapMaterial(string text, string note, double maxWidth)
            {
                List<string> wrappedLines = new List<string>();
                string patternGost = @"(?:((ГОСТ|ОСТ|ТУ)+[\s,\S]+))$";
                // строка с регулярным выражением для выделения размера лакоткани

                string patternSize = @"^(\d\d?[,\.]?\d?\s?([мм]?)+\s?(\*|х|x)\s?\d\d?[,\.]?\d?\s?[мм]+[,\.]*)";
                string patternWordSize = @"(лента шириной \d\s?мм *)";
                //string patternSize = @"(?:(\d[,\.]?\d?\d?\s?()+(\*|х|x)\d\d?[,\.]?\d?\d?))"; для поиска размеров в наименовании
                string sizeLKM;
                string sizeWordLKM = string.Empty;
                Regex regex;
                Match match;

                regex = new Regex(patternGost);
                match = regex.Match(text);
                string gost = match.Value;
                gost = gost.Trim();
                if (gost != String.Empty)
                {
                    int indx = text.IndexOf(gost);
                    text = text.Remove(indx);
                    regex = new Regex(patternSize);
                    match = regex.Match(note);
                    sizeLKM = match.Value;

                    if ((text.Contains("ЛКМ")) && (sizeLKM != string.Empty))
                    {
                        sizeLKM = match.Value;
                        int indx2 = note.IndexOf(sizeLKM);
                        note = note.Remove(indx2, sizeLKM.Length);
                    }
                    else
                    {

                        regex = new Regex(patternWordSize);
                        match = regex.Match(note);
                        sizeWordLKM = match.Value;

                        if ((text.Contains("ЛКМ")) && (sizeWordLKM != string.Empty))
                        {
                            sizeWordLKM = match.Value;
                            int indx2 = note.IndexOf(sizeWordLKM);
                            note = note.Remove(indx2, sizeLKM.Length);
                        }
                    }
                    wrappedLines.AddRange(Wrap(text, _font, maxWidth));



                    wrappedLines.Add(gost);

                    if ((text.Contains("ЛКМ")) && (sizeLKM != string.Empty))
                    {

                        wrappedLines.Add(sizeLKM);
                    }
                    else
                    {
                        if ((text.Contains("ЛКМ")) && (sizeWordLKM != string.Empty))
                        {

                            wrappedLines.Add(sizeWordLKM);
                        }
                    }

                    return wrappedLines.ToArray();
                }

                else return Wrap(text, _font, maxWidth);
            }

            public string[] Wrap(string text, double maxWidth)
            {
                return Wrap(text, _font, maxWidth);
            }

            public string[] WrapNoteMaterial(string name, string text, double maxWidth)
            {
                // строка с регулярным выражением для выделения размера лакоткани
                // string patternSize = @"^(\d\d?[,\.]?\d?\s?([мм]?)+(\*|х|x)\d\d?[,\.]?\d?\s?[мм]+[,\.]*)";
                string patternSize = @"^(\d\d?[,\.]?\d?\s?([мм]?)+\s?(\*|х|x)\s?\d\d?[,\.]?\d?\s?[мм]+[,\.]*)";
                string patternWordSize = @"(лента шириной \d\s?мм *)";
                string sizeLKM;
                string sizeWordLKM;

                Regex regex;
                Match match;

                regex = new Regex(patternSize);
                match = regex.Match(text);
                sizeLKM = match.Value;

                //Вырезаем из примечания Лакоткани строку с размерами, если она есть
                if ((name.Contains("ЛКМ")) && (sizeLKM != string.Empty))
                    text = text.Remove(0, sizeLKM.Length).Trim();
                else
                {
                    regex = new Regex(patternWordSize);
                    match = regex.Match(text);
                    sizeWordLKM = match.Value;

                    //Вырезаем из примечания Лакоткани строку с размерами, если она есть
                    if ((name.Contains("ЛКМ")) && (sizeWordLKM != string.Empty))
                        text = text.Remove(0, sizeWordLKM.Length).Trim();
                }

                return Wrap(text, _font, maxWidth);
            }

            public string[] WrapFullName(string textDesignation, string[] text, double maxWidth, string docName, string docDenotation)
            {
                List<string> wrappedLines = new List<string>();
                int indx = text[0].IndexOf(docName, StringComparison.InvariantCultureIgnoreCase);
                // на тот случай, если имя уже разделено на несколько строк
                if (indx == -1 && text.Length > 1)
                {
                    for (int i = 1; i < text.Length; i++)
                    {
                        // "" потому, что есть пробелы в text
                        var firstLine = string.Join("", text.Take(i + 1).ToArray());
                        indx = firstLine.IndexOf(docName, StringComparison.InvariantCultureIgnoreCase);
                        if (indx > -1)
                        {
                            var tmp = text.Skip(i + 1).ToList();
                            tmp.Insert(0, firstLine);
                            text = tmp.ToArray();
                            break;
                        }
                    }
                }
                int indxDesignation = textDesignation.IndexOf(docDenotation);

                if (indx == 0)
                {

                    // Для полных наименований наименований исполнений берем полное наименование основного исполнения, добавляя к нему номер, все остальное переносим
                    if ((text[0].Contains(docName + "-1")) || (text[0].Contains(docName + "-2")) || (text[0].Contains(docName + "-3"))
                        || (text[0].Contains(docName + "-4")) || (text[0].Contains(docName + "-5")) || (text[0].Contains(docName + "-6"))
                        || (text[0].Contains(docName + "-7")) || (text[0].Contains(docName + "-8")) || (text[0].Contains(docName + "-9")))
                    {

                        wrappedLines.AddRange(Wrap(text[0].Substring(indx, docName.Length + 2), _font, maxWidth));
                        text[0] = text[0].Remove(indx, docName.Length + 2);
                        foreach (string t in text)
                            wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                        return wrappedLines.ToArray();
                    }

                    text[0] = text[0].Remove(indx, docName.Length);
                    wrappedLines.AddRange(Wrap(docName, _font, maxWidth));
                    foreach (string t in text)
                        wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                    return wrappedLines.ToArray();

                }
                else
                {
                    if (indxDesignation == 0)
                    {
                        wrappedLines.AddRange(Wrap(docName, _font, maxWidth));
                        foreach (string t in text)
                            wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                        return wrappedLines.ToArray();
                    }
                    else
                    {
                        wrappedLines.AddRange(Wrap(docName, _font, maxWidth));

                        foreach (string t in text)
                            wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                    }
                }
                return wrappedLines.ToArray();

            }

            public string[] WrapNotFullName(string[] text, double maxWidth, string docName)
            {
                List<string> wrappedLines = new List<string>();
                int indx = text[0].IndexOf(docName, StringComparison.InvariantCultureIgnoreCase);
                // на тот случай, если имя уже разделено на несколько строк
                if (indx == -1 && text.Length > 1)
                {
                    for (int i = 1; i < text.Length; i++)
                    {
                        // "" потому, что есть пробелы в text
                        var firstLine = string.Join("", text.Take(i + 1).ToArray());
                        indx = firstLine.IndexOf(docName, StringComparison.InvariantCultureIgnoreCase);
                        if (indx > -1)
                        {
                            var tmp = text.Skip(i + 1).ToList();
                            tmp.Insert(0, firstLine);
                            text = tmp.ToArray();
                            break;
                        }
                    }
                }

                if (indx == 0)
                {
                    text[0] = text[0].Remove(indx, docName.Length);
                    foreach (string t in text)
                        wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                    return wrappedLines.ToArray();
                }


                foreach (string t in text)
                {
                    wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                }
                return wrappedLines.ToArray();
            }

            public string[] WrapNotFullNameKomplekt(string[] text, double maxWidth, string docName)
            {
                List<string> wrappedLines = new List<string>();
                int indx = text[0].IndexOf(docName, StringComparison.InvariantCultureIgnoreCase);
                if (indx == 0)
                {
                    text[0] = text[0].Remove(indx, docName.Length);


                }


             //   wrappedLines.AddRange(Wrap(docName, _font, maxWidth));

                foreach (string t in text)
                {

                    if (t.Trim() != string.Empty)
                        wrappedLines.AddRange(Wrap(t, _font, maxWidth));
                }
                // MessageBox.Show(wrappedLines.Count + "");
                return wrappedLines.ToArray();
            }

            // Разбиваем название документов на строки в соостветствии с "клюквой"
            public string[] WrapDocument(string text, double maxWidth)
            {
                List<string> wrappedLines = new List<string>();
                string wrapStr = string.Empty;
                var allElemetsForSplit = new List<string>();
                allElemetsForSplit.AddRange(elementsForSplit.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
                foreach (var element in allElemetsForSplit)
                    if (text.Contains(element))
                    {
                        int indx = text.IndexOf(element);
                        if (indx != 0)
                            text = text.Insert(indx, "*#*");
                    }

                text = text + "*#*";
                while (text.Contains("*#*"))
                {
                    int indx = text.IndexOf("*#*");
                    wrapStr = text.Remove(indx, text.Length - indx);
                    wrappedLines.Add(wrapStr);
                    text = text.Remove(0, indx + 3);

                }

                return wrappedLines.ToArray();
            }

            private static string elementsForSplit = @"Альбом
Базовая
Базовый
Ведомость
Габаритный
Генератор
Данные
Заключение
Инструкция
Комплект
Комплекс
Лист
Монитор
Монтажный
Общее
Общие
Операционная
Описание
Опись
Пакет
Паспорт
Перечень
Пояснительная
Приложение
Программа
Руководство
Сборочный
Система
Спецификация
Схема
Таблица
Текст
Технические
Удостоверяющий
Формуляр
Химмотологическая
Чертеж
Электромонтажный
Этикетка";


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

            // Вставка в наименование для типов Детали и Стандартные изделия примечание с наименованием заготовки(если такое примечание есть
            public string[] Wrap(string[] nameLines, string text, string note, double maxWidth)
            {
                List<string> wrappedLines = new List<string>();

                if (note.Contains("(Заготовка") || note.Contains("(заготовка"))
                    wrappedLines.AddRange(Wrap(note, maxWidth));
                else
                    if (note.Contains("Заготовка") || note.Contains("(заготовка"))
                    wrappedLines.AddRange(Wrap("(" + note + ")", maxWidth));
                else return nameLines;


                List<string> summaryLines = new List<string>();
                if (nameLines[0].Contains("(Заготовка)"))
                {
                    nameLines[0] = nameLines[0].Remove(nameLines[0].IndexOf("(Заготовка)"), 11);

                }
                if (nameLines[0].Contains("(заготовка)"))
                {
                    nameLines[0] = nameLines[0].Remove(nameLines[0].IndexOf("(заготовка)"), 11);

                }
                summaryLines.AddRange(nameLines.ToList());

                summaryLines.AddRange(wrappedLines);
                return summaryLines.ToArray();


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
                        {
                            wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                        }
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
                        if (((GetTextWidth(wrappedLines[wrappedLines.Count - 1] + wrappedLines[wrappedLines.Count - 2], _font)) < maxWidth + maxWidth / 5)
                            && ((wrappedLines[wrappedLines.Count - 2] != "Резистор") && (wrappedLines[wrappedLines.Count - 2] != "Конденсатор")
                            && (wrappedLines[wrappedLines.Count - 2] != "Транзисторная матрица") && (wrappedLines[wrappedLines.Count - 2] != "Оптопара транзисторная")
                            && (wrappedLines[wrappedLines.Count - 2] != "Блок трансформаторов") && (wrappedLines[wrappedLines.Count - 2] != "Катушка индуктивности")))
                        {
                            wrappedLines[wrappedLines.Count - 2] = wrappedLines[wrappedLines.Count - 2] + " " + wrappedLines[wrappedLines.Count - 1];
                            wrappedLines.Remove(wrappedLines[wrappedLines.Count - 1]);
                            if (remPart.Trim() != String.Empty)
                                wrappedLines.Add(endPart);
                            else
                            {
                                if (!wrappedLines.Contains(endPart) && endPart.Trim() != string.Empty)
                                    wrappedLines.Add(endPart);
                            }

                            return wrappedLines.ToArray();
                        }
                        else if (remPart.Trim() != String.Empty)
                            wrappedLines.Add(endPart);
                        else
                        {
                            if (!wrappedLines.Contains(endPart) && endPart.Trim() != string.Empty)
                                wrappedLines.Add(endPart);
                        }
                        return wrappedLines.ToArray();
                    }


                    else if (remPart.Trim() != String.Empty)
                        wrappedLines.Add(endPart);
                    else
                    {
                        if (!wrappedLines.Contains(endPart) && endPart.Trim() != string.Empty)
                            wrappedLines.Add(endPart);
                    }

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
                    else
                        if ((sectionName == TFDBomSection.Assembly) && (text.Contains("Катушка индуктивности")))
                    {
                        regex = new Regex(patternBeginPart);
                        match = regex.Match(text);
                        beginPart = match.Value;
                        beginPart = beginPart.Trim();

                        if (GetTextWidth(remPart, _font) > maxWidth)
                            wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                        else wrappedLines.Add(beginPart);

                        if (beginPart != text)
                        {
                            text = text.Remove(0, beginPart.Length);
                            wrappedLines.AddRange(Wrap(text, maxWidth));
                        }

                        return wrappedLines.ToArray();

                    }
                    else
                    {
                        wrappedLines.AddRange(Wrap(text, maxWidth));

                    }


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
    }
    //--------Форма для Спецификации---------------


    namespace Form1Namespace
    {
        using System;
        using System.Collections.Generic;
        using System.ComponentModel;
        using System.Data;
        using System.Drawing;
        using System.Text;
        using System.Windows.Forms;
        using System.Diagnostics;

        public partial class MainForm
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

            public void progressBarMax(int val)
            {
                progressBar1.Value = 1;
                progressBar1.Minimum = 1;
                progressBar1.Step = 1;
                progressBar1.Maximum = val;

            }

            public void progressBarParam(int step)
            {
                progressBar1.Step = step;
                progressBar1.PerformStep();
            }
            public void setStage(Stages stage)
            {
                progressBarParam(4);
                switch (stage)
                {
                    case Stages.Initialization:
                        ProcessingLabel.Text = "Инициализация...";
                        break;
                    case Stages.DataAcquisition:
                        ProcessingLabel.Text = "Получение данных";
                        break;
                    case Stages.DataProcessing:
                        ProcessingLabel.Text = "Обработка данных";
                        break;
                    case Stages.ReportGenerating:
                        ProcessingLabel.Text = "Создание отчёта";
                        break;
                    case Stages.Done:
                        break;
                    default:
                        Debug.Assert(false);
                        break;
                }
                this.Update();
            }


        }
        //--------------Код Дизайн----------------

        public partial class MainForm : System.Windows.Forms.Form
        {

            private System.Windows.Forms.ProgressBar progressBar1;

            private System.Windows.Forms.Label ProcessingLabel;

            private System.Windows.Forms.Label LogLabel;

            private System.Windows.Forms.TextBox LogTextBox;

            private void InitializeComponent()
            {
                this.progressBar1 = new System.Windows.Forms.ProgressBar();
                this.ProcessingLabel = new System.Windows.Forms.Label();
                this.LogLabel = new System.Windows.Forms.Label();
                this.LogTextBox = new System.Windows.Forms.TextBox();
                this.SuspendLayout();
                // 
                // progressBar1
                // 
                this.progressBar1.Location = new System.Drawing.Point(21, 184);
                this.progressBar1.Name = "progressBar1";
                this.progressBar1.Size = new System.Drawing.Size(433, 23);
                this.progressBar1.TabIndex = 3;
                // 
                // ProcessingLabel
                // 
                this.ProcessingLabel.Location = new System.Drawing.Point(21, 165);
                this.ProcessingLabel.Name = "ProcessingLabel";
                this.ProcessingLabel.Size = new System.Drawing.Size(136, 16);
                this.ProcessingLabel.TabIndex = 2;
                this.ProcessingLabel.Text = "Инициализация...";
                this.ProcessingLabel.UseCompatibleTextRendering = true;
                // 
                // LogLabel
                // 
                this.LogLabel.Location = new System.Drawing.Point(21, 9);
                this.LogLabel.Name = "LogLabel";
                this.LogLabel.Size = new System.Drawing.Size(100, 18);
                this.LogLabel.TabIndex = 1;
                this.LogLabel.Text = "Журнал операций";
                this.LogLabel.UseCompatibleTextRendering = true;
                // 
                // LogTextBox
                // 
                this.LogTextBox.BackColor = System.Drawing.SystemColors.ActiveBorder;
                this.LogTextBox.Location = new System.Drawing.Point(21, 30);
                this.LogTextBox.Multiline = true;
                this.LogTextBox.Name = "LogTextBox";
                this.LogTextBox.Size = new System.Drawing.Size(433, 121);
                this.LogTextBox.TabIndex = 0;
                // 
                // MainForm
                // 
                this.ClientSize = new System.Drawing.Size(477, 228);
                this.Controls.Add(this.progressBar1);
                this.Controls.Add(this.ProcessingLabel);
                this.Controls.Add(this.LogLabel);
                this.Controls.Add(this.LogTextBox);
                this.DoubleBuffered = false;
                this.Text = "Спецификация ГОСТ 19.202";
                this.ResumeLayout(false);
                this.PerformLayout();
            }
        }
    }



    namespace Globus.TFlexDocs
    {
        using System;
        using System.Collections.Generic;
        using System.Diagnostics;
        using System.Text;
        using System.Text.RegularExpressions;


        /** Нормализация строк.
         *
         *  Нормализация представляет собой приведение строки к виду, который минимизирует
         *  количество возможных опечаток. Например, несколко последовательно идущих пробелов
         *  приводятся к одному пробелу, буква О приводится к цифре 0 и т.п.
         *
         *  Используется при сравнении строк. Пусть S1 и S2 - строки, а S1' и S2' - их
         *  нормелизованные формы. Тогда если S1' = S2', но S1 != S2, строки S1  и S2
         *  различны в результате опечатки (например, вместо цифры 0 имеется буква О)
         */
        public class TextNormalizer
        {
            // Все замены нормализации, кроме смены регистра, реализованы в виде
            // регулярных выражений. Приведение сроки к нижнему регистру выполняется
            // перед применение регулярных выражений.
            //
            // ВНИМАНИЕ! После выполнения импорта править эти правила нельзя.
            static readonly string[,] _normalizationReplaces = {

            // Замена латинских букв на сходные по начертанию русские.
            // При этом учитывается сходство не только строчных букв, но и заглавных.
            {@"a(?=.*[а-я])|(?<=[а-я].*)a", "а"},
            {@"b(?=.*[а-я])|(?<=[а-я].*)b", "в"}, // 'B', 'В'
            {@"c(?=.*[а-я])|(?<=[а-я].*)c", "с"},
            {@"e(?=.*[а-я])|(?<=[а-я].*)e", "е"},
            {@"h(?=.*[а-я])|(?<=[а-я].*)h", "н"}, // 'H', 'Н'
            {@"k(?=.*[а-я])|(?<=[а-я].*)k", "к"},
            {@"m(?=.*[а-я])|(?<=[а-я].*)m", "м"}, // 'M', 'М'
            {@"n(?=.*[а-я])|(?<=[а-я].*)n", "п"},
            {@"о(?=.*[а-я])|(?<=[а-я].*)о", "о"},
            {@"p(?=.*[а-я])|(?<=[а-я].*)p", "р"},
            {@"t(?=.*[а-я])|(?<=[а-я].*)t", "т"}, // 'T', 'Т'
            {@"x(?=.*[а-я])|(?<=[а-я].*)x", "х"},

            // Замена примыкающей к цифре справа латинской или кириллической буквы о на цифру 0
            // Пример: 1О -> 10
            {@"(?<=[0-9][o\u043e]*)[o\u043e]", "0"},

            // Замена примыкающей к цифре слева латинской или кириллической буквы о на цифру 0
            // Пример: O1 -> 01
            {@"[o\u043e](?=[o\u043e]*[0-9])", "0"},

            // Замена примыкающей находящихся между буквами нуля или латинской О на кириллическую О
            // Пример: лмн0пр -> лмнопр
            {@"(?<=\p{Ll}[0o]*)[0o](?=[0o]*\p{Ll})", "\u043e"}, // Cyrillic Small Letter O

            // Замена первой в строке 0 или латинской О, примыкающей к букве, на кириллическую О
            // Пример: 0пр -> опр
            {@"^[0o](?=[0o]*\p{Ll})", "\u043e"}, // Cyrillic Small Letter O

            // Замена последней в строке 0 или латинской О, примыкающей к букве, на кириллическую О
            // Пример: лмн0 -> лмно
            {@"(?<=\p{Ll}[0o]*)[0o]$", "\u043e"}, // Cyrillic Small Letter O

            // Замена одиноко стоящих последовательностей o на 0
            // Пример: А.ОО.Б -> А.00.Б
            {@"(?<=\W[o\u043e]*)[o\u043e](?=[o\u043e]*\W)", "0"},
            {@"^[o\u043e](?=[o\u043e]*\W)", "0"},
            {@"(?<=\W[o\u043e]*)[o\u043e]$", "0"},

            // Замена латинской I на 1
            {@"i", "1"},

            // Замена кириллической П на 11 (частто вместо римской II пишут П)
            {@"п", "11"},

            // Замена кириллической Ш на 111 (частто вместо римской III пишут Ш)
            {@"ш", "111"},

            // Замена латинской х, кириллической х или звездочки *, находящихся между цифрами и, возможно,
            // отбитых пробелами на знак подчеркивания _
            // Пример: 10 x 16 -> 10_16
            {@"(?<=[0-9])\s*[xх\*×]\s*(?=[0-9])", "_"},

            // Удаление специальных символов в начале строки
            {@"^\W+", ""},

            // Удаление специальных символов в конце строки
            {@"\W+$", ""},

            // Удаление специальных символов между буквами
            {@"(?<=\p{Ll})\W+(?=\p{Ll})", ""},

            // Удаление специальных символов между буквой и цифрой
            {@"(?<=\p{Ll})\W+(?=[0-9])", ""},

            // Удаление специальных символов между цифрой и буквой
            {@"(?<=[0-9])\W+(?=\p{Ll})", ""},

            // Удаление специальных символов между цифрами
            {@"(?<=[0-9])\W+(?=[0-9])", "_"},

            // Удаление остальных неизвестных символов
            {@"[^а-яa-z_0-9]", ""},
        };

            static readonly Dictionary<Regex, string> _replaces;

            static TextNormalizer()
            {
                _replaces = new Dictionary<Regex, string>();
                int lowerBound = _normalizationReplaces.GetLowerBound(0);
                int upperBound = _normalizationReplaces.GetUpperBound(0);
                for (int i = lowerBound; i <= upperBound; ++i)
                {
                    Regex re = new Regex(_normalizationReplaces[i, 0], RegexOptions.Compiled);
                    _replaces.Add(re, _normalizationReplaces[i, 1]);
                }
            }

            //* Получение нормальной формы строки
            public string GetNormalForm(string text)
            {
                // Преобразование к нижнему регистру
                string normal = text.ToLower();

                // Применение замен нормализации
                foreach (KeyValuePair<Regex, string> pair in _replaces)
                {
                    Regex regex = pair.Key;
                    string replaced = regex.Replace(normal, pair.Value);
                    normal = replaced;
                }

                return normal;
            }

        }
    }
}