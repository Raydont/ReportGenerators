using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using TFlex;
//using TFlex.DOCs.Common;
using TFlex.Drawing;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Reporting;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
//using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
//using Globus.TFlexDocs;
//using FormAttributes;
using DocumentAttributes;
using Rectangle = TFlex.Drawing.Rectangle;
using Size = System.Drawing.Size;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;

namespace SpecificationESPD
{
    public class CADGenerator : IReportGenerator
    {
        //static CADGenerator()
        //{
        //    //Добавляем ссылку на директории с API CAD 11            
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
                //Справочник Группы и пользователи
                ReferenceInfo info = ReferenceCatalog.FindReference(GroupsAndUsersReferenceGuid);
                Reference reference = info.CreateReference();
                var chiefDesigner = reference.Find(PostChiefDesignerGuid).Children.Where(t=>t.Class.IsInherit(UserTypeGuid)).OrderBy(t=>t.ToString()).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).ToArray();
                var chiefVPMO = reference.Find(PostChiefVPMOGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).OrderBy(t => t.ToString()).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).ToArray();
                var engineers05 = reference.Find(PostEngineerDepartmentEDGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList();
                engineers05.AddRange(reference.Find(PostEngineerDepartment51Guid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList());
                engineers05.AddRange(reference.Find(PostEngineerDepartment52Guid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList());
                engineers05.AddRange(reference.Find(PostEngineerDepartment54Guid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList());
                engineers05.AddRange(reference.Find(PostEngineerDepartment55Guid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList());
                
                //057 был удален из справочника Группы и пользователи
                //List<string> children = new List<string>();
                //children = reference.Find(PostEngineerDepartment57Guid)?.Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList();
                //if (children != null) engineers05.AddRange(children);

                engineers05.AddRange(reference.Find(PostEngineerDepartment58Guid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToList());
                var rateChecker = reference.Find(PostRateCheckerGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t=> GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToArray();
                var chiefTeam = reference.Find(PostChiefTeamGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToArray();
                var representativeVPM = reference.Find(RepresentativeVPMOGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToArray();
                var zgkz = reference.Find(PostZGKZGuid).Children.Where(t => t.Class.IsInherit(UserTypeGuid)).Select(t => GetBeginStringInitialUser(t[ShortUserNameGuid].Value.ToString())).OrderBy(t => t.ToString()).ToArray();

                attr.tBoxYearMake.Text = DateTime.Now.Year.ToString();
                //Главный конструктор
                attr.comboBoxChiefDesigner.Items.AddRange(chiefDesigner);
                if (chiefDesigner.Count() > 0)
                    attr.comboBoxChiefDesigner.Text = chiefDesigner.FirstOrDefault().ToString();

                //Начальник ВПМО
                attr.comboBoxChiefVPMO.Items.AddRange(chiefVPMO);
                if (chiefVPMO.Count() > 0)
                    attr.comboBoxChiefVPMO.Text = chiefVPMO.FirstOrDefault().ToString();

                //Исполнитель
                attr.comboBoxMaker.Items.AddRange(engineers05.OrderBy(t=>t).ToArray());
                if (engineers05.Where(t => t.ToString().Contains(context.AuthorInfo.AuthorLastName)).Count() > 0)
                    attr.comboBoxMaker.Text = engineers05.Where(t=>t.ToString().Contains(context.AuthorInfo.AuthorLastName)).FirstOrDefault().ToString();

                //Нормоконтроль
                attr.comboBoxRateChecker.Items.AddRange(rateChecker);
                if (rateChecker.Count() > 0)
                    attr.comboBoxRateChecker.Text = rateChecker.FirstOrDefault().ToString();

                //Начальник бригады
                attr.comboBoxChiefTeam.Items.AddRange(chiefTeam);
                if (chiefTeam.Count() > 0)
                   attr.comboBoxChiefTeam.Text = chiefTeam.FirstOrDefault().ToString();

                //Представитель ВПМО
                attr.comboBoxRepresentative.Items.AddRange(representativeVPM);
                if (representativeVPM.Count() > 0)
                    attr.comboBoxRepresentative.Text = representativeVPM.FirstOrDefault().ToString();

                //ЗГКЗ
                attr.comboBoxZGKZ.Items.AddRange(zgkz);
                if (zgkz.Count() > 0)
                    attr.comboBoxZGKZ.Text = zgkz.FirstOrDefault().ToString();

                // Вызов Диалога настройки параметров отчета
                attr.DotsEntry(context);


            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (document != null)
                    document.Close();
            }
        }

        public string GetBeginStringInitialUser (string shortUserName)
        {
            var regexInitial = new Regex(@"[А-Я]\.\s?[А-Я]\.");

            if (regexInitial.Match(shortUserName).Success)
            {
                var newShortUserName = regexInitial.Match(shortUserName).Value.ToString();
                var lastName = shortUserName.Replace(newShortUserName,"").Trim();
                newShortUserName += " " + lastName;
                return newShortUserName;
            }
            else
                return shortUserName;

        }

        public static readonly Guid GroupsAndUsersReferenceGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea"); // Справочник Группы и пользователи
        public static readonly Guid UserTypeGuid = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");                // Тип сотрудник
        public static readonly Guid ShortUserNameGuid = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");           // Короткое имя пользователя
        public static readonly Guid PostChiefDesignerGuid = new Guid("60934222-b8cd-475e-83c9-e8cf9e5c2894");       // Главный конструктор
        public static readonly Guid PostChiefVPMOGuid = new Guid("8313ab27-e676-4396-9363-fad865656d1e");           // Начальник ВПМО    
        public static readonly Guid PostEngineerDepartmentEDGuid = new Guid("f16dadfb-801f-4c27-8d91-e8912630884d");// Инженер группы ЭД
        public static readonly Guid PostEngineerDepartment51Guid = new Guid("1a0d6a7c-dc61-4144-96e9-db0d889d5b07");// Инженер группы 051
        public static readonly Guid PostEngineerDepartment52Guid = new Guid("72786a8d-2e5e-454e-a518-eaeedbec100b");// Инженер группы 052
        public static readonly Guid PostEngineerDepartment54Guid = new Guid("4ae15a03-8b97-4e31-916a-cf5ac077800a");// Инженер группы 054
        public static readonly Guid PostEngineerDepartment55Guid = new Guid("e6fd6a8c-a0ca-459e-beba-a7f20a446d3f");// Инженер группы 055
        public static readonly Guid PostEngineerDepartment57Guid = new Guid("2608e20d-d5d1-4851-b48c-65fded962121");// Инженер группы 057
        public static readonly Guid PostEngineerDepartment58Guid = new Guid("ce430512-3b61-4a68-8451-f70e068a81ce");// Инженер группы 058
        public static readonly Guid PostRateCheckerGuid = new Guid("fe977b49-a953-4747-9c28-4d3e97e93a93");         // Группа 451
        public static readonly Guid PostChiefTeamGuid = new Guid("ae69dcfa-de56-4baa-8310-3a6e7057dd16");           // Начальники РБ отдела 05
        public static readonly Guid RepresentativeVPMOGuid = new Guid("75a90351-b689-41fe-a0f1-2e6a00bc3147");      // Подразделение ВП МО
        public static readonly Guid PostZGKZGuid = new Guid("cc2e06c2-a38e-47f7-b9c5-987cda478488");                // Начальник 05 отдела

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


 //=================МАКРОС СП ЕСПД ================================   
}

namespace DocumentAttributes
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Windows.Forms;
    using SpecificationReport.Espd;
    using Form1Namespace;
    //using FormAttributes;

    public class DocumentAttributes
    {
        // TFlexDocs _tflexdocs;
        int _id;
        public string NameProgFullName = ""; 
        string _designedBy = "";
        string _controlledBy = "";
        string _nControlBy = "";
        string _approvedBy = "";
        int _firstPageNumber = 1;
        bool _addChangelogPage = true;
        string _sidebarSignHeader1 = "";
        string _sidebarSignName1 = "";
        string _sidebarSignHeader2 = "";
        string _sidebarSignName2 = "";
        string _sidebarSignHeader3 = "";
        string _sidebarSignName3 = "";
        bool _fullNames = false;

        int _countEmptyLinesBefore;
        int _countEmptyLinesAfter;

        string _zakaz = string.Empty;
        string _nameProgram = string.Empty;
        string _yearMake = string.Empty;

        string _goev = string.Empty;
        string _militaryChief = string.Empty;
        string _designer = string.Empty;
        string _norm = string.Empty;
        string _chiefTeam = string.Empty;
        string _vPMO = string.Empty;
        private string _chiefZgKZTeam = string.Empty;
        private string _surnameZGKZ = string.Empty;

        bool _checkMilitaryChief = true;
        bool _checkChiefTeam = true;
        private bool _zgkz = true;
        bool _checkVPMO = true;

        bool _titulList = false;
        bool _lu = false;

        string _litera = string.Empty;

      
        static readonly DocumentAttributes _empty = new DocumentAttributes();

        public DocumentAttributes()
        {
        }

        public DocumentAttributes(TFDDocument document)
        {
            _id = document.ObjectID;
            //ReadAttributes(document,documentAttr);
        }

        /* private void ReadAttributes(TFDDocument document,DocumentAttributes documentAttr)
        {
            // tflexdocs.LoadDocumentParameters(document, LOAD_PARAMETERS.ALL);

            string designedBy = documentAttr.DesignedBy;//tflexdocs.GetParameterValue(document, ParametersGUIDs.DesignedBy);
            _designedBy = (designedBy == null) ? "" : designedBy;

            string controlledBy = documentAttr.ControlledBy;//tflexdocs.GetParameterValue(document, ParametersGUIDs.ControlledBy);
            _controlledBy = (controlledBy == null) ? "" : controlledBy;

            string nControlBy = documentAttr.NControlBy;//tflexdocs.GetParameterValue(document, ParametersGUIDs.NControlBy);
            _nControlBy = (controlledBy == null) ? "" : nControlBy;

            string approvedBy = documentAttr.ApprovedBy;//tflexdocs.GetParameterValue(document, ParametersGUIDs.ApprovedBy);
            _approvedBy = (approvedBy == null) ? "" : approvedBy;

            _firstPageNumber = documentAttr.FirstPageNumber;//tflexdocs.GetIntegerParameterValue(document, ParametersGUIDs.FirstPageNumber);

            _addChangelogPage = documentAttr.AddChangelogPage;//tflexdocs.GetBooleanParameterValue(document, ParametersGUIDs.AddChangelogPage);

            string sidebarSignHeader1 = documentAttr.SidebarSignHeader1;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader1);
            _sidebarSignHeader1 = (sidebarSignHeader1 == null) ? "" : sidebarSignHeader1;

            string sidebarSignName1 = documentAttr.SidebarSignName1;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName1);
            _sidebarSignName1 = (sidebarSignName1 == null) ? "" : sidebarSignName1;

            string sidebarSignHeader2 = documentAttr.SidebarSignHeader2;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader2);
            _sidebarSignHeader2 = (sidebarSignHeader2 == null) ? "" : sidebarSignHeader2;

            string sidebarSignName2 = documentAttr.SidebarSignName2;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName2);
            _sidebarSignName2 = (sidebarSignName2 == null) ? "" : sidebarSignName2;

            string sidebarSignHeader3 = documentAttr.SidebarSignHeader3;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignHeader3);
            _sidebarSignHeader3 = (sidebarSignHeader3 == null) ? "" : sidebarSignHeader3;

            string sidebarSignName3 = documentAttr.SidebarSignName3;//tflexdocs.GetParameterValue(document, ParametersGUIDs.SidebarSignName3);
            _sidebarSignName3 = (sidebarSignName3 == null) ? "" : sidebarSignName3;

            _fullNames = documentAttr.FullNames;//tflexdocs.GetBooleanParameterValue(document, ParametersGUIDs.FullNames);
        } */

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
            set { _controlledBy = value;  }
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
            set { _firstPageNumber = value;  }
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

        public bool Lu
        {
            get { return _lu; }
            set { _lu = value; }
        }

        public bool TitulList
        {
            get { return _titulList; }
            set { _titulList = value; }
        }

        public bool CheckVPMO
        {
            get { return _checkVPMO; }
            set { _checkVPMO = value; }
        }

        public bool CheckChiefTeam
        {
            get { return _checkChiefTeam; }
            set { _checkChiefTeam = value; }
        }

        public bool CheckMilitaryChief
        {
            get { return _checkMilitaryChief; }
            set { _checkMilitaryChief = value; }
        }

        public string VPMO
        {
            get { return _vPMO; }
            set { _vPMO = value; }
        }

        public string ChiefTeam
        {
            get { return _chiefTeam; }
            set { _chiefTeam = value; }
        }


        public string Norm
        {
            get { return _norm; }
            set { _norm = value; }
        }

        public string Designer
        {
            get { return _designer; }
            set { _designer = value; }
        }

        public string MilitaryChief
        {
            get { return _militaryChief; }
            set { _militaryChief = value; }
        }

        public string Goev
        {
            get { return _goev; }
            set { _goev = value; }
        }

        public string YearMake
        {
            get { return _yearMake; }
            set { _yearMake = value; }
        }

        public string NameProgram
        {
            get { return _nameProgram; }
            set { _nameProgram = value; }
        }

        public string Zakaz
        {
            get { return _zakaz; }
            set { _zakaz = value; }
        }

        public string Litera
        {
            get { return _litera; }
            set { _litera = value; }
        }

        public bool Zgkz
        {
            get { return _zgkz; }
            set { _zgkz = value; }
        }

        public string ChiefZGKZ
        {
            get { return _chiefZgKZTeam; }
            set { _chiefZgKZTeam = value; }
        }

        public string SurnameZgkz
        {
            get { return _surnameZGKZ; }
            set { _surnameZGKZ = value; }
        }
    }
}

namespace Globus.TFlexDocs
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.IO.IsolatedStorage;
    using System.Text;
    using System.Text.RegularExpressions;
    //using tfdAPI;
    using SpecificationReport.Espd;

    public class TFlexDocs : IDisposable
    {
        public TFlexDocs()
        {
        }

        ~TFlexDocs()
        {
            Dispose();            
        } 

        /// Освобаждает ресурсы DOCs API
        public void Dispose()
        {
        }

        // public void GetBomEmptyLineCounts(int docID, out int countBefore, out int countAfter)
        // {
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
            // countBefore = 10;
            // countAfter = 10;
        // }

        public string GetDocumentInfoString(TFDDocument document)
        {
           // Debug.Assert(document != null);
            //LoadDocumentParameters(document, LOAD_PARAMETERS.ALL);

            StringBuilder info = new StringBuilder();

            // Наименование
            string namePar = document.Naimenovanie;//document.Params.GetItem(ParametersGUIDs.Name) as TFDParameter;
            if (namePar != null && namePar.Length > 0)
                info.Append(namePar);

            // Обозначение
            string designationPar = document.Denotation; //document.Params.GetItem(ParametersGUIDs.Designation) as TFDParameter;
            if (designationPar != null && designationPar.Length > 0)
            {
                if (info.Length > 0)
                    info.Append(" ");
                info.Append(designationPar);
            }

            // Идентификатор
            if (info.Length > 0)
                info.Append(" ");
            info.AppendFormat("[{0}]", document.ObjectID.ToString());

            return info.ToString();
        }
    }

    public class DocumentParameters : IDictionary<int, string>, IEquatable<DocumentParameters>
    {
        int _id;
        int _parentId;
        string _name;
        string _designation;
        string _litera;
        SortedDictionary<int, string> _values;

        /* public DocumentParameters(TFlexDocs tflexdocs, TFDDocument document, int[] parameters)
        {
            if (tflexdocs == null)
                throw new ArgumentException("Нулевой указатель на экземпляр TFlexDocs");
            Load(tflexdocs, document, parameters);
        } */

        /* void Load(TFlexDocs tflexdocs, TFDDocument document, int[] parameters)
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

            _values = new SortedDictionary<int,string>();
            Dictionary<int, string> all = tflexdocs.GetDocumentParameters(document);
            foreach (int id in parameters)
            {
                if (all.ContainsKey(id))
                    _values.Add(id, all[id]);
            }
        } */

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

        public string Litera { get { return _litera; } }

        public string GetValue(int id)
        {
            return _values[id];
        }

        /* public string[] GetReport(TFlexDocs tflexdocs)
        {
            return GetReport(tflexdocs, false, false, false);
        } */

        /* public string[] GetReport(TFlexDocs tflexdocs, bool zeroAsNull, bool noAsNull, bool excludeEmpty)
        {
            List<string> text = new List<string>();
            foreach(KeyValuePair<int, string> pair in _values)
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
        } */

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
        //public static readonly string OtherDocuments = "";

    }
}

namespace SpecificationReport.Espd
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.Text;
    using System.Text.RegularExpressions;
	using System.Windows.Forms;
    using Form1Namespace;    
    using SpecificationESPD;
    using Globus.TFlexDocs;    
    using DocumentAttributes;    

    /// Программа, на которую оформляется спецификация.
    public class Program
    {
	
//        TFlexDocs _tflexdocs;

        // Само изделие
        SpecDocument _programSpecDoc;

        // Состав (документы спецификации)
        SpecDocumentCollection _documents;

        // Хэши дочерних объектов
//        SortedDictionary<int, byte[]> _childrenHashes;

/*        protected SortedDictionary<int, byte[]> ChildrenHashes
        {
            get { return _childrenHashes; }
        } */

        SpecDocumentComparer _orderComparer;

        // Атрибуты документа
        Dictionary<string, DocumentAttributes> _attributes;

/*        byte[] _hash;

        internal byte[] Hash
        {
            get { return _hash; }
        }*/


    //    DocumentHashAlgorithm _hashAlgorithm;

//        HierarchyReport _report;

        public Program(/*TFlexDocs tflexdocs*/)
        {
            // _tflexdocs = tflexdocs;
            _documents = new SpecDocumentCollection();
            _attributes = new Dictionary<string, DocumentAttributes>();
            _orderComparer = new SpecDocumentComparerEspd();
             // _hashAlgorithm = new DocumentHashAlgorithm(_tflexdocs, TreeHashAlgorithmSettingsA.Parameters);
            // _childrenHashes = new SortedDictionary<int, byte[]>();
        }
 	 

        public void ReadContent(TFDDocument document, DocumentAttributes documentAttr)
        {            
//            _report = new HierarchyReport(_tflexdocs, HierarchyReportConsts.HierarchyReportParameters);
//            _report.AddDocument(document, null);

            using (SqlConnection connection = ApiDocs.GetConnection(true))
            {                
                string countQuery = String.Format(
                      @"SELECT COUNT(n.[Name])
                        FROM Nomenclature n JOIN NomenclatureHierarchy nh ON n.s_ObjectID = nh.s_ObjectID
                        WHERE nh.s_ActualVersion = 1
                          AND nh.s_Deleted = 0
                          AND n.s_ClassID IN ({1},{2},{3},{4})
                          AND nh.s_ParentID = {0}", document.ObjectID, TFDClass.ComplexProgram, TFDClass.ComponentProgram,
                                                                       TFDClass.Document, TFDClass.OtherDocument);
                SqlCommand countCommand = new SqlCommand(countQuery, connection);
                countCommand.CommandTimeout = 0;
                int count = (int)countCommand.ExecuteScalar() ;
               
                EspdSpecificationCadReport.m_form.progressBarMax(count);

                string objectListQuery = String.Format(
                     @"SELECT n.[Name],
                               n.Denotation,
                               nh.Remarks,
                               nh.BOMSection
                        FROM Nomenclature n JOIN NomenclatureHierarchy nh ON n.s_ObjectID = nh.s_ObjectID
                        WHERE nh.s_ActualVersion = 1 AND nh.s_Deleted = 0
                          AND n.s_ActualVersion = 1 AND n.s_Deleted = 0
                          AND n.s_ClassID IN ({1},{2},{3},{4})
                          AND nh.s_ParentID = {0} 
                          -- 790 - Комплекс (программы)
                          -- 789 - Компонент (программы)
                          -- 691 - Документ 
                          -- 729 - Другие документы", document.ObjectID, TFDClass.ComplexProgram, TFDClass.ComponentProgram,
                                                                         TFDClass.Document, TFDClass.OtherDocument);

                SqlCommand objectListCommand = new SqlCommand(objectListQuery, connection);
                objectListCommand.CommandTimeout = 0;
                SqlDataReader objectListReader = objectListCommand.ExecuteReader();

                /*DocumentAttributes attributes = new DocumentAttributes(document);
                string name = document.Naimenovanie;//_tflexdocs.GetParameterValue(doc, ParametersGUIDs.Name);
                if (!_attributes.ContainsKey(name))
                    _attributes.Add(name, attributes);
                SpecDocumentEspd baseDoc = new SpecDocumentEspd(document, _programSpecDoc);
                _documents.Add(baseDoc);*/
                while (objectListReader.Read())
                {
                    TFDDocument childDocument = new TFDDocument();
                    if (documentAttr.FullNames && TFDBomSection.Documentation == objectListReader.GetString(3))
                    {
                        childDocument.Naimenovanie = documentAttr.NameProgFullName + " " + objectListReader.GetString(0);
                    }
                    else
                    {
                        childDocument.Naimenovanie = objectListReader.GetString(0);
                    }

                    childDocument.Denotation = objectListReader.GetString(1);
                    childDocument.Remarks = objectListReader.GetString(2);
                    childDocument.BomSection = objectListReader.GetString(3);

                    SpecDocumentEspd child = new SpecDocumentEspd(childDocument, _programSpecDoc, documentAttr);                    
                    _documents.Add(child);
                    EspdSpecificationCadReport.m_writeToLog(childDocument.Naimenovanie + " " + childDocument.Denotation);
                }                
            }
          /*  int result = document.LoadComposition(LOAD_TYPE.DOCUMENT_CONTENT, LOAD_MODE.ITEMS,
                LOAD_PARAMETERS.ALL);
            if (result != 0)
            {
                throw new ApplicationException(String.Format(
                    "Ошибка при загрузке состава объекта {0}", _tflexdocs.GetDocumentInfoString(document)));
            }*/

            _programSpecDoc = new SpecDocumentEspd(document, _programSpecDoc, documentAttr);
            //_hash = HashAlgorithm.ComputeHash(document);
            
           // ITFDEnumerator childDocEnum = document.Documents.GetEnumerator();
           // childDocEnum.Reset();
            
	//		int count=document.Documents.Count;
			
           // result = childDocEnum.MoveNext();
 			
          //  while (result > 0)
           // {
//				
                /*TFDDocument childDocument = (TFDDocument)(childDocEnum.Current);
                if (childDocument == null)
                {
                    throw new ApplicationException(String.Format(
                        "Ошибка при загрузке вложенного объекта из состава {0}",
                        _tflexdocs.GetDocumentInfoString(document)));
                }*/

                //ReadChildDocument(childDocument, document);
 			    
              //  result = childDocEnum.MoveNext();
         //   }
 			
            // Сортировка документов
            _documents.Sort(_orderComparer);
        }

/*        public byte[] CalcualteHashAlgorithm1()
        {
            SortedDictionary<int, byte[]> hashes = new SortedDictionary<int, byte[]>();
            hashes.Add(SpecDocument.ID, _hash);
            foreach (KeyValuePair<int, byte[]> pair in _childrenHashes)
                hashes.Add(pair.Key, pair.Value);
            byte[] hash = HashAlgorithm.ComputeHash(hashes.Values, "1");
            return hash;
        } */

        /// Список категорий, объекты которых включаются в спецификацию
       /* static readonly List<DocsCategory> _specificationCategories = new List<DocsCategory>(new DocsCategory[] {
            DocsCategory.Documentation,
            DocsCategory.ProgramComplexes,
            DocsCategory.ProgramComponents
        }); */

        /**
         *  ProcessDocument child document
         */
        protected virtual void ReadChildDocument(TFDDocument doc, TFDDocument parent)
        {
           // DocsCategory documentCategory = (DocsCategory)doc.Category;

            // Пропускаем категории Отчеты
           // if (documentCategory == DocsCategory.Reports)
           //     return;

           // byte[] hash = HashAlgorithm.ComputeHash(doc);
           // _childrenHashes.Add(doc.id, hash);

//            _report.AddDocument(doc, parent);

           /* if (doc.Class == (int)DocsClass.DocumentAttributes
                || doc.Class == (int)DocsClass.Link
                   && doc.ClassRef == (int)DocsClass.DocumentAttributes)
            {
                // Атрибуты документа
                DocumentAttributes attributes = new DocumentAttributes(doc);
                string name = doc.Naimenovanie;//_tflexdocs.GetParameterValue(doc, ParametersGUIDs.Name);
                if (!_attributes.ContainsKey(name))
                    _attributes.Add(name, attributes);
            }
            else if (_specificationCategories.Contains((DocsCategory)doc.Category))
            {
                // Условие включения документа в спецификацию.
                // T-FLEX DOCs имеет для этого параметр "Включать в спецификацию".
                // Но он не используется по следующим причинам:
                // - лишняя сущность: включение объекта в дерево уже означает включение в спецификацию
                // - при создании ссылок категории "Первичная применяемость" логично было бы снимать флаг
                //   "Включать в спецификацию". Но на это уходит много времени.
                // Поэтому было решено включать в спецификацию все объекты определенной категории.
                SpecDocumentEspd child = new SpecDocumentEspd(doc, _programSpecDoc);
 				//EspdSpecificationCadReport.m_writeToLog("Добавлен документ: "+ child.ToString());
                _documents.Add(child);
            } */
        }

        /**
         *  Получение записей спецификации по ГОСТ 19.202-78
         */
        public void AddSpecRowsFormEspd(SpecFormEspd specForm, DocumentAttributes documentAttributes)
        {
            // Создание списка документов
            List<SpecDocumentEspd> documents = new List<SpecDocumentEspd>();
            foreach (SpecDocument d in _documents)
                documents.Add(d as SpecDocumentEspd);

            // Добавление строк таблицы спецификации
            AddSpecSectionsAndRows(specForm, documents.ToArray(),documentAttributes);
        }

        /**
         * Получение строк таблицы спецификации для одного документа
         * с учетом уже занесенных в таблицу документов, а также
         * документов, находящихся в очереди.
         */

        protected SpecRowFormEspd[] CreateSpecRows(SpecDocumentEspd document,
                                                   List<SpecDocumentEspd> addedDocuments,
                                                   List<SpecDocumentEspd> queueDocuments,
                                                   SpecFormEspd specForm, DocumentAttributes documentAttributes)
        {
            Debug.Assert(queueDocuments.Contains(document));
            Debug.Assert(!addedDocuments.Contains(document));
            string name = string.Empty;

            if (documentAttributes.FullNames)
                name = document.Name;
            else
                name = GetSpecDocumentNameField(document);

            SpecRowFormEspd[] records = SpecRowFormEspd.CreateRecord(document.Designation,
                                                                     name,
                                                                     document.Note,
                                                                     specForm.TextFormatter);
            queueDocuments.Remove(document);
            addedDocuments.Add(document);

            return records;
        }

        protected string GetSpecDocumentNameField(SpecDocumentEspd document)
        {
            string name = document.Name;

            if (document.BomGroup == TFDBomSection.Documentation)//.IsContainedInGroup(BomGroupID.ProgramDocuments))
            {
                int index = Name.IndexOf("Спецификация");
				string withoutSpecName = Name.Substring(0,index-1);
				withoutSpecName = withoutSpecName.Trim();
				/*
                 * в разделе "Документация" для документов, входящих в основной комплект
                    документов специфицируемого изделия и составляемых на данное изделие,
                    - только наименование документов, например: "Сборочный чертеж",
                    "Габаритный чертеж", "Технические условия". Для документов на
                    неспецифицированные составные части - наименование изделия и
                    наименование документа;

                 */
                if (document.Designation.StartsWith(Designation)
                    && document.Name.StartsWith(withoutSpecName))
                {
                    name = document.Name.Substring(withoutSpecName.Length);
                    name = name.Trim();					
                }
            }
            else if (document.BomGroup == TFDBomSection.Components 
                     || document.BomGroup == TFDBomSection.Complex)
            {
				//   name += " Спецификация";
            }
            return name;
        }

        void AddBomGroupHeader(SpecFormEspd specForm, string section, bool newPage)
        {
            if (newPage)
                specForm.NewPage();

            // переход на новую страницу выполняется, если на странице
            // осталось менее 10 пустых строк
            if (specForm.CurrentPageAvialableRows < 10)
                specForm.NewPage();

            // Пустая строка перед заголовком раздела
            specForm.AddEmptyRow();

            // Заголовок раздела
            SpecFormCell headerCell = new SpecFormCell(section,
                SpecFormCell.UnderliningFormat.Underlined, SpecFormCell.AlignFormat.Center);
            SpecRowFormEspd header = new SpecRowFormEspd(SpecFormCell.Empty, headerCell, SpecFormCell.Empty);
            specForm.Add(header);

            // Пустая строка после заголовка раздела
            // specForm.AddEmptyRow();
        }

        /**
         *  Получение записей спецификации формы массива документов
         */
        protected virtual void AddSpecSectionsAndRows(SpecFormEspd specForm, SpecDocumentEspd[] documents, DocumentAttributes documentAttributes)
        {
            // Занесенные в спецификацию документы
            List<SpecDocumentEspd> addedDocuments = new List<SpecDocumentEspd>();

            // Документы, подлежащие занесению в спецификацию
            List<SpecDocumentEspd> queueDocuments = new List<SpecDocumentEspd>(documents);

            bool firstHeader = true;

            /*BomGroup[] sections = new BomGroup[] {
                _tflexdocs.BomGroupsReference.GetBomGroup((int)BomGroupID.ProgramDocuments),
                _tflexdocs.BomGroupsReference.GetBomGroup((int)BomGroupID.ProgramComplexes),
                _tflexdocs.BomGroupsReference.GetBomGroup((int)BomGroupID.ProgramComponents) */
            string[] sections = new string[] {
                TFDBomSection.Documentation, TFDBomSection.Complex, TFDBomSection.Components};

            // Цикл по разделам спецификации
            for (int i=0; i< sections.Length; i++)
            {
                bool headerCreated = false;

                // Поиск и обработка очередного документа из состава текущего раздела спецификации
                while (true)
                {
                    // Поиск очередного документа из текущего раздела
                    SpecDocumentEspd currentDocument = null;
                    foreach (SpecDocumentEspd qdoc in queueDocuments)
                    {
                        if (qdoc.BomGroup == sections[i])
                        {
                            currentDocument = qdoc;
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
                            //          || specForm.CurrentPageAvialableRows < 10; // осталось меньше 10 строк на листе
                        AddBomGroupHeader(specForm, sections[i], newPage);
                        headerCreated = true;
                        firstHeader = false;
                    }

                    // Создание записей. Коллекции addedDocuments, queueDocuments модифицируются
                    SpecRowFormEspd[] rows = CreateSpecRows(currentDocument,
                                                            addedDocuments,
                                                            queueDocuments,
                                                            specForm,  documentAttributes);

                    // Заполнение таблицы спецификации
                    if (specForm.CurrentPageAvialableRows < rows.Length + 3)
                        specForm.NewPage();
                    if (specForm.AtTopOfPage())
                        specForm.AddEmptyRows(2);

                    if (currentDocument.EmptyLinesBefore > 0)
                        specForm.AddEmptyRows(currentDocument.EmptyLinesBefore);
                    
                    specForm.Add(rows);

                    // /*Две пустые строки*/ 0 пустых строк после записи или переход на новую страницу
                    if (currentDocument.EmptyLinesAfter > 0)
                        specForm.AddEmptyRows(currentDocument.EmptyLinesAfter);
                    else
                        specForm.AddEmptyRowsOrFinishCurrentPage(0);
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

/*        protected TFlexDocs TFlexDocs
        {
            get { return _tflexdocs; }
        } */

/*       public virtual string[] GetHierarchyReport()
        {
            return _report.GetReport(HierarchyReport.ReportOptions.ExcludeIDs);
        } */

        public string Name
        {
            get { return _programSpecDoc.Name; }
        }

        public string Designation
        {
            get { return _programSpecDoc.Designation; }
        }

        public string Litera
        {
            get { return _programSpecDoc.Litera; }
        }

        public SpecDocument SpecDocument
        {
            get { return _programSpecDoc; }
        }

        public SpecDocumentCollection Documents
        {
            get { return _documents; }
        }

        protected SpecDocumentComparer OrderComparer
        {
            get { return _orderComparer; }
        }

/*        protected DocumentHashAlgorithm HashAlgorithm
        {
            get { return _hashAlgorithm; }
        } */
    }  
}

namespace SpecificationReport
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

   // using tfdAPI;
    using Globus.TFlexDocs;
    using SpecificationReport.Espd;
    using DocumentAttributes;
        
    //
    //  Документ спецификации.
    //
    public class SpecDocument
    {
        // Параметры объекта (документа) T-FLEX DOCs

        int _id;                // ID объекта
        // string _guid;           // GUID объекта
        string _docInfoString;  // Строка для метода ToString()

        // Основные параметры документа

        SpecDocument _parent;   // изделие, в состав которого входит данное изделие
        string _designation;    // обозначение
        string _name;           // наименование документа

        string _bomGroup;     // раздел спецификации
        string _litera;       // литера
        string _note;           // примечание

        int _emptyLinesBefore;
        int _emptyLinesAfter;

        /**
         *  Constructor
         */
        public SpecDocument(TFDDocument doc, SpecDocument parent)
        {
            _parent = parent;
            ReadDocumentParameters(doc, parent);
        }

        public SpecDocument(TFDDocument doc, SpecDocument parent, DocumentAttributes documentAttr)
        {
            _parent = parent;
            ReadDocumentParameters(doc, parent, documentAttr);
        }

        //
        //  Чтение параметров документа TFlexDocs
        //
        protected virtual void ReadDocumentParameters(TFDDocument doc, SpecDocument parent, DocumentAttributes documentAttr)
        {
            _emptyLinesBefore = documentAttr.CountEmptyLinesBefore;
            _emptyLinesAfter = documentAttr.CountEmptyLinesAfter;
            ReadDocumentParameters(doc, parent);
        }

        protected virtual void ReadDocumentParameters(TFDDocument doc, SpecDocument parent)
        {
            if (doc == null)
                throw new ArgumentException("Parameter is null", "doc");

            _id = doc.ObjectID;//Math.Abs(doc.id);
            TFlexDocs tflexdocs = new TFlexDocs();
            _docInfoString = tflexdocs.GetDocumentInfoString(doc);

            // отбрасываем пробел и символы в круглых скобках в конце обозначения
            string docsDesignation = doc.Denotation.Trim();//tflexdocs.GetParameterValue(doc, ParametersGUIDs.Designation);
            Match match = Regex.Match(docsDesignation, "^(?<p1>.*?)(?: \\(.*\\))?$");
            _designation = match.Groups["p1"].Value;

            _name = doc.Naimenovanie;//tflexdocs.GetParameterValue(doc, ParametersGUIDs.Name);

            // Раздел спецификации
            _bomGroup = doc.BomSection;//ReadBomGroup(tflexdocs, doc);

            // Литера
            //_litera = // ReadDocumentLitera(tflexdocs, doc);
            _litera = doc.Litera;

            // Примечание
            _note = doc.Remarks;//tflexdocs.GetParameterValue(doc, ParametersGUIDs.BomNote);

            // tflexdocs.GetBomEmptyLineCounts(doc.ObjectID, out _emptyLinesBefore, out _emptyLinesAfter);
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

        public override string ToString()
        {
            return _docInfoString;
        }

        public int ID
        {
            get { return _id; }
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
        }

        protected SpecDocument Parent
        {
            get { return _parent; }
        }

        public string BomGroup
        {
            get { return _bomGroup; }
        }

        public string Litera
        {
            get { return _litera; }
        }

        public string Note
        {
            get { return _note; }
        }

        public int EmptyLinesBefore
        {
            get { return _emptyLinesBefore; }
        }

        public int EmptyLinesAfter
        {
            get { return _emptyLinesAfter; }
        }
    }

    //
    //  Компаратор документов спецификации
    //
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
                    double element = Convert.ToDouble(m.Value);
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

    //
    //  Коллекция документов спецификации
    //
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

    //
    //  Ячейка формы спецификации
    //
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

        public string[] Wrap(string text, double maxWidth)
        {
            return Wrap(text, _font, maxWidth);
        }

        public string[] WrapName(string text, double maxWidth)
        {
            string patternWrapWord = "(База данных|Спецификация|Описание применения|Текст программы|Лист утверждения|Программа|Библиотека|Операционная система|Модуль|Информационно-удостоверяющий лист)";
            Regex regex = new Regex(patternWrapWord);
            MatchCollection match;
            match = regex.Matches(text);

            int indx;
            List<string> rowList = new List<string>();
            List<string> wrappedLines = new List<string>();
            string beginPart;


            for (int i = 0; i < match.Count; i++)
            {
                if (match[i].Value.Trim() != string.Empty)
                    rowList.Add(match[i].Value);
            }

            if (rowList.Count > 0)
            {
                foreach (string row in rowList)
                {
                    indx = text.IndexOf(row);
                    beginPart = text.Remove(indx);
                    if (beginPart.Trim() != string.Empty)
                    {
                        text = text.Remove(0, beginPart.Length);
                        text = text.Trim();
                        wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                    }
                 
                    if (text.Contains(row + "."))
                    {
                        if (row != "Информационно-удостоверяющий лист")
                            wrappedLines.Add(row+".");
                        else
                        {
                            wrappedLines.Add("Информационно-удостоверяющий");
                            wrappedLines.Add("лист");
                        }
                        text = text.Remove(0, row.Length + 2);
                    }
                    else
                    {
                        if (row != "Информационно-удостоверяющий лист")
                            wrappedLines.Add(row);
                        else
                        {
                            wrappedLines.Add("Информационно-удостоверяющий");
                            wrappedLines.Add("лист");
                        }
                        text = text.Remove(0, row.Length);
                    }
                }
                return wrappedLines.Select(t=>t + " ").ToArray();
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

        /// Выделение слога - последовательности символов с завершающим пробелом
        /// или символом мягкого переноса. Символ мягкого переноса (если присутствует)
        /// выносится в отдельную группу
        static readonly Regex _syllableRegex = new Regex(
            @"(?<syllable>[^\x20\u00AD]+)(?<softHyphen>\u00AD+)?",
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
            @"[\p{Ll}\p{Lu}]+[^\p{Ll}\p{Lu}]*", RegexOptions.Compiled);

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

namespace SpecificationReport.Espd
{

    using System;
    using System.Collections.Generic;
    using System.Collections;
    using System.Data;
    using System.Reflection;
    using System.IO;
    using System.Xml;
    using TFlex;
    using TFlex.Model;
    using TFlex.Model.Model2D;
    using System.Diagnostics;
    using Form1Namespace;
    //using FormAttributes;

    using Globus.TFlexDocs;
    using DocumentAttributes;
    using SpecificationReport;    

    public delegate void LogDelegate(string line);

    //--------------
    public class TFDDocument
    {
        // Порядок полей в запросе чтения документа

        public int ObjectID;
        public int ParentID;
        public string Class;
        public double Amount;
        public string Naimenovanie;
        public string Denotation;
        public string Remarks;
        public string BomSection = string.Empty;
        public string Litera;

        //private List<TFDDocument> _childDocuments;

        public TFDDocument()
        {
        }

        public TFDDocument(int id, int parent, string docClass)
        {
            ObjectID = id;
            ParentID = parent;
            Class = docClass;

        }

        public TFDDocument(int documentID)
        {
            NomenclatureReference refNomenclature = new NomenclatureReference();

            TFlex.DOCs.Model.Structure.ParameterInfo classInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
            TFlex.DOCs.Model.Structure.ParameterInfo nameInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
            TFlex.DOCs.Model.Structure.ParameterInfo denotationInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);
            //TFlex.DOCs.Model.Structure.ParameterInfo bomSectionInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.);

            ParentID = 0;
            ObjectID = documentID;
            NomenclatureObject nomenclatureObject = (NomenclatureObject)refNomenclature.Find(documentID);            
            Class = nomenclatureObject[classInfo].Value.ToString();            
            Amount = 1;
            Naimenovanie = (string)nomenclatureObject[nameInfo].Value;
            Denotation = (string)nomenclatureObject[denotationInfo].Value;
            // Литера в Номенклатуре
            Litera = (string)nomenclatureObject[993].Value;
            Remarks = String.Empty;            
        }
    }

    internal static class ApiDocs
    {
        public static SqlConnection GetConnection(bool open)
        {
            // string connectionString = "Persist Security Info=False;Integrated Security=true;database=TFlexDOCs;server=S2"; //SRV1";

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
	
    public class EspdSpecificationCadReport : IDisposable
    {
 	
        static bool _calledFromDocs;

        public static bool CalledFromDocs
        {
            get { return _calledFromDocs; }
        }

        static EspdSpecificationCadReport()
        {
            _calledFromDocs = false;
        }

        /**
         *  Точка входа при создании отчета из DOCs
         */
 		public static MainForm m_form;       
        public static LogDelegate m_writeToLog;
		 
        public static void Make(IReportGenerationContext context, DocumentAttributes documentAttr)
        {
            m_form = new MainForm();
            m_form.Show();
            m_writeToLog = new LogDelegate(m_form.writeToLog);
        
            m_form.setStage(MainForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");
            m_form.progressBarParam(2);
            m_writeToLog("Получение идентификатора корневого документа");	
 			
            _calledFromDocs = true;

            try
            {
                using (EspdSpecificationCadReport report = new EspdSpecificationCadReport())
                {

                    int docID = report.GetDocsDocumentID(context);
                    if (docID == -1) return;
                    report.Make(docID,context,documentAttr);                    
                }
            }
            catch (Exception e)
            {
                string message = String.Format("ОШИБКА: {0}", e.ToString());
                System.Windows.Forms.MessageBox.Show(message, "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

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

            NomenclatureReference refNomenclature = new NomenclatureReference();
            int objectClassID = refNomenclature.Find(documentID).Class.Id;
            if ((objectClassID != TFDClass.ComplexProgram) && (objectClassID != TFDClass.ComponentProgram))
            {
                MessageBox.Show("Неверный тип объекта.\nДля формирования спецификации ЕСПД выберите объект типа 'Комплекс(программы)' или 'Компонент (программы)'");
                return -1;
            }
           
          m_writeToLog("Идентификатор корневого документа: "+documentID.ToString());	
 		  m_form.progressBarParam(2);	
            
            // УДАЛИТЬ ПОСЛЕ ТЕСТОВ
            // documentID = 3436;

            return documentID;
        }

      //  TFlexDocs _tflexdocs;

        /**
         *  Конструктор
         */
        public EspdSpecificationCadReport()
        {
        }

        ~EspdSpecificationCadReport()
        {
            Dispose();
        }

        public void Dispose()
        {
        } 

        // Create report
        public void Make(int documentID, IReportGenerationContext context,DocumentAttributes documentAttr)
        {
            //TFDDocument document = _tflexdocs.GetDocument(documentID);            
            TFDDocument document = new TFDDocument(documentID);
            Make(document, context,documentAttr);
        }

        // Create report
        public void Make(TFDDocument document, IReportGenerationContext context,DocumentAttributes documentAttr)
        {
            try
            {
                // Чтение данных из DOCs
 			    m_form.setStage(MainForm.Stages.DataAcquisition);	
 		        m_writeToLog("=== Получение данных ===");		
                Program program = new Program();
                program.ReadContent(document, documentAttr);

                // Подготовка данных для CAD
                m_form.setStage(MainForm.Stages.DataProcessing);
                m_writeToLog("=== Обработка данных ===");
 		        m_form.progressBarParam(2);	
                //DocumentAttributes attrs = program.GetAttributes(Settings.AttributesObjectName);
                // string hashString = DocumentHashAlgorithm.ConvertToHexString(program.CalcualteHashAlgorithm1());
                SpecFormEspd specForm = new SpecFormEspd(program.SpecDocument, documentAttr);

                program.AddSpecRowsFormEspd(specForm, documentAttr);

               
                // Формирование документа CAD
 			    m_form.setStage(MainForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");	
 			    m_form.progressBarParam(4);	

                Document cadDocument = GetCadDocument(context);
                FillCadDocument(specForm, cadDocument);
              
                // Установка параметров объекта DOCs
#if !TEST
                //string path = cadDocument.FilePath;
                // TFDDocument docsDocument = _tflexdocs.GetReportDocument(path);
                // SetDocsDocumentParameters(docsDocument, hashString);
                // attrs.CopyContent(docsDocument);

                // Сохранение отчета об иерархии
//                string[] report = program.GetHierarchyReport();
//                SaveHierarchyReport(docsDocument, report);
#endif

                // TFlex.Application.ExitSession();
 			    m_form.setStage(MainForm.Stages.Done);
                m_writeToLog("=== Работа завершена ===");
                m_writeToLog("Ведомость спецификаций ГОСТ 19.202 сформирована");
                m_form.Close();

                // Сохранение
                // SaveCadDocument(cadDocument);
                cadDocument.Close();
                System.Diagnostics.Process.Start(context.ReportFilePath);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString(), "Ошибка",
                                                     System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        Document GetCadDocument(IReportGenerationContext context)
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

        void SaveCadDocument(Document cadDocument)
        {
            bool saved = cadDocument.Save();
            if (!saved)
            {
                string message = String.Format("Ошибка сохранения файла отчета");
                throw new Exception(message);
            }
        }

        /* void SetDocsDocumentParameters(TFDDocument docsDocument, string hashString)
        {
            docsDocument.Params.Item(ParametersGUIDs.Hash, hashString);
            docsDocument.Params.Item(ParametersGUIDs.NeedBOM, "0");
            docsDocument.Params.Item(ParametersGUIDs.PartOfProduct, "0");

            int result = docsDocument.Save();
            if (result != 0)
                throw new Exception("Ошибка сохранения парамтров объекта отчета");
        } */

        /* void SaveHierarchyReport(TFDDocument docsDocument, string[] report)
        {
            // Путь ко временному файлу, в который сохраняется отчет об иерархии
            string name = _tflexdocs.GetParameterValue(docsDocument, ParametersGUIDs.Name);
            string reportPath = Path.GetTempPath() + Path.DirectorySeparatorChar + name + ".hr.txt";
            //string reportPath = Path.GetTempFileName();

            // Сохранение отчета об иерархии в файл
            StreamWriter reportStream = File.CreateText(reportPath);
            foreach (string s in report)
                reportStream.WriteLine(s);
            reportStream.Close();

            // Создание объекта T-FLEX DOCS, содержащего отчет об иерархии
            TFDDocument reportDocument = new TFDDocument();
            _tflexdocs.SetParameterValue(reportDocument, ParametersGUIDs.Name, "Отчет об иерархии");
            reportDocument.Class = (int)DocsClass.HierarchyReport;
            _tflexdocs.SetParameterValue(reportDocument, ParametersGUIDs.ObjectClass,
                                         Convert.ToString((int)DocsClass.Report));
            reportDocument.Category = (int)DocsCategory.Related;
            _tflexdocs.SetParameterValue(reportDocument, ParametersGUIDs.ObjectCategory,
                                         Convert.ToString((int)DocsCategory.Related));
            reportDocument.Parent = docsDocument.id;

            // создание объекта на рабочем столе
            TFDDocument desktopDocument = (TFDDocument)_tflexdocs.Application.CreateObject_v2(CREATE_MODE.DOCUMENT,
                                                                                              reportDocument, 0);
            if (desktopDocument == null || desktopDocument.id == 0)
            {
                string message = String.Format("Ошибка создания на рабочем столе документа {0}",
                    _tflexdocs.GetDocumentInfoString(reportDocument));
                throw new Exception(message);
            }

            desktopDocument.Category = (int)DocsCategory.Related;
            _tflexdocs.SetParameterValue(desktopDocument, ParametersGUIDs.ObjectCategory,
                                         Convert.ToString((int)DocsCategory.Related));

            TFDFile docsFile = new TFDFile();
            docsFile.state = OBJECT_STATE.STATE_ADD;
            TFDStringCollection additionalFiles = new TFDStringCollection();

            string fileTypeStr = desktopDocument.Params.ItemT(ParametersGUIDs.DocType);
            int fileType = System.Convert.ToInt32(fileTypeStr);
            bool saveWithOriginalName = true;

            int result = docsFile.CreateFromFile(desktopDocument, reportPath, additionalFiles, fileType,
                                                 saveWithOriginalName);
            if (result != 0)
            {
                string message = String.Format("Файл не создан. {0}. {1}", _tflexdocs.Application.Error.MessageT(),
                    _tflexdocs.Application.Error.MessageCommentT());
                throw new Exception(message);
            }
        } */

        void FillCadDocument(SpecFormEspd specForm, Document cadDocument)
        {
            cadDocument.BeginChanges("Формирование отчета");
            ParagraphText textTable = GetSpecificationTextTable(cadDocument);

            // Начало редактирования чертежа
            textTable.BeginEdit();

            // Заполнение таблицы
            Table table = textTable.GetTableByIndex(0);
            SpecRowFormEspd[] records = specForm.GetRows();
            foreach (SpecRowFormEspd record in records)
            {
                InsertRow(textTable, table, record);
            }
           
            // Добавление листа регистрации изменений
            string changeLogPageName = null;
            if (specForm.Attributes.AddChangelogPage)
            {
                changeLogPageName = AddMorePages(specForm, cadDocument);
            }

            // Добавление Титульного листа
            string titulPageName = null;
            if (specForm.Attributes.TitulList)
            {
                titulPageName = AddMorePages(specForm, cadDocument);
            }

            // Добавление листа утверждения
            string luPageName = null;
            if (specForm.Attributes.Lu)
            {
                luPageName = AddMorePages(specForm, cadDocument);
            }

            // Завершение редактирования чертежа
            textTable.EndEdit();

            // Добавление "форматок"
            AddPageForms(specForm, cadDocument, changeLogPageName, titulPageName, luPageName, specForm.Attributes.Lu);

            cadDocument.EndChanges();

            cadDocument.Save();
        } 

        ParagraphText GetSpecificationTextTable(Document document)
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
                throw new ApplicationException("На чертеже не найден объект $TextTablePg1");

            return textTable;
        }

        void InsertRow(ParagraphText parText, Table table, SpecRowFormEspd row)
        {
            table.InsertRows(1, (uint)(table.CellCount - 1), Table.InsertProperties.After);
            uint newCellsCount = (uint)table.CellCount;
            uint columnsCount = (uint)table.ColumnCount;

            InsertTableText(table, newCellsCount - columnsCount, 0, parText, row.Designation);
            InsertTableText(table, newCellsCount - columnsCount + 1, 0, parText, row.Name);
            InsertTableText(table, newCellsCount - columnsCount + 2, 0, parText, row.Note);
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

        /// Добавление листа регистрации изменений.
        /// @return
        ///     Имя страницы с листром регистрации
        string AddMorePages(SpecFormEspd specForm, Document cadDocument)
        {
            Page page = new Page(cadDocument);
            return page.Name;
        }

        /// Добавление форм в виде фгарментов. Заполнение переменных фрагментов.
        void AddPageForms(SpecFormEspd specForm, Document cadDocument, string changeLogPageName,
            string titulPageName, string luPageName, bool addLU)
        {
            FileReference fileReference = new FileReference();

            var nameSpecification = specForm.Attributes.NameProgram;

            if (specForm.Attributes.NameProgram.Contains("Программа"))
            {
                int indx = specForm.Attributes.NameProgram.IndexOf("Программа");
                if(indx>0)
                    nameSpecification = specForm.Attributes.NameProgram.Insert(indx - 1, "\n");
            }


            FileLink pageFormLink = null;
            FileLink regListFormLink = null;
            FileLink titulFormLink = null;
            FileLink luFormLink = null;

            int pageNumber = specForm.Attributes.FirstPageNumber - 1;
            foreach (Page page in cadDocument.Pages)
            {
                if (page.PageType != PageType.Normal)
                    continue;

                ++pageNumber;
                FileLink link = null;
                //Вставка титульного листа
                if (page.Name == titulPageName)
                {
                    if (titulFormLink == null)
                    {
                        var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.TitulListName);
                        fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                        titulFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                        /*titulFormLink = new FileLink(cadDocument,
                            fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.TitulListName).LocalPath,
                            true); */
                        /*titulFormLink = new FileLink(cadDocument,
                            Settings.PathToTemplates + @"\" + Settings.TitulListName, true);*/
                    }
                    link = titulFormLink;
                }
                //Вставка листа утверждения
                else if (page.Name == luPageName)
                {
                    if (luFormLink == null)
                    {
                        var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ListUtvergdeniaNameHard);
                        fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                        luFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                        /*luFormLink = new FileLink(cadDocument,
                            fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ListUtvergdeniaNameHard).LocalPath,
                            true);*/
                        /*luFormLink = new FileLink(cadDocument,
                            Settings.PathToTemplates + @"\" + Settings.ListUtvergdeniaNameHard, true);*/
                    }
                    link = luFormLink;
                }
                // вставка листа регистрации изменений
                else if (page.Name == changeLogPageName)
                {                    
                    if (regListFormLink == null)
                    {
                        var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ChangelogName);
                        fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                        regListFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                        /* regListFormLink = new FileLink(cadDocument,
                            fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.ChangelogName).LocalPath,
                            true); */
                        /* regListFormLink = new FileLink(cadDocument,
                            Settings.PathToTemplates + @"\" + Settings.ChangelogName,
                            true); */
                     }
                    link = regListFormLink;
                }
                // вставка формы 1а в виде фрагмента
                else if (page.Name.StartsWith("Страница"))
                {                 
                    if (pageFormLink == null)
                    {
                        var fileObject = fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.FormName);
                        fileObject.GetHeadRevision(); //Загружаем последнюю версию в рабочую папку клиента

                        pageFormLink = new FileLink(cadDocument, fileObject.LocalPath, true);
                        /* pageFormLink = new FileLink(cadDocument,
                            fileReference.FindByRelativePath(Settings.PathToTemplates + Settings.FormName).LocalPath,
                            true); */
                        /*pageFormLink = new FileLink(cadDocument,
                            Settings.PathToTemplates + @"\" + Settings.FormName,
                            true);*/
                    }
                    link = pageFormLink;
                }
                else
                {
                    throw new ApplicationException("Неизвестное наименование страницы документа CAD");
                }

                page.Rectangle = new Rectangle(0,0,210,297);

                Fragment formFragment = new Fragment(link);
                formFragment.Page = page;

                
                // задание значений переменных фрагмента
                foreach (FragmentVariableValue var in formFragment.VariableValues)
                {
                    switch (var.Name)
                    {
                        case "$nameProgram":
                            var.TextValue = nameSpecification;
                            break;
                        case "$specify":
                            var.TextValue = "Спецификация";
                            break;
                        case "$zakaz":
                            var.TextValue = specForm.Attributes.Zakaz;
                            break;
                        case "$year":
                            var.TextValue = specForm.Attributes.YearMake;
                            break;
                        case "$lifs":
                            long count = cadDocument.Pages.Count + specForm.Attributes.FirstPageNumber - 2;
                            if (addLU == true)
                                count--;
                            var.TextValue = "Листов "+count.ToString();
                            break;
                        case "$designlu":
                            var.TextValue = specForm.Designation + "-ЛУ";
                            break;
                        case "$goev":
                            var.TextValue = specForm.Attributes.Goev;
                            break;
                        case "$vpmo":
                            var.TextValue = specForm.Attributes.MilitaryChief;
                            break;
                        case "$razrab":
                            var.TextValue = specForm.Attributes.Designer;
                            break;
                        case "$n_contr":
                            var.TextValue = specForm.Attributes.Norm;
                            break;
                        case "$zgkz":
                            var.TextValue = specForm.Attributes.ChiefTeam;
                            break;
                        case "$vpmo1":
                            var.TextValue = specForm.Attributes.VPMO;
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
                            long total = cadDocument.Pages.Count + specForm.Attributes.FirstPageNumber - 1;
                            var.TextValue = total.ToString();
                            break;
                        case "$sidebar_text":
                            var.TextValue = GetSidebarText(specForm);
                            break;
                        // case "$product_guid":
                        //    var.TextValue = specForm.DocumentGuid;
                        //    break;
                        case "$created_date":
                            DateTime now = DateTime.Now;
                            var.TextValue = now.ToString(Settings.DateTimeFormat);
                            break;
                        /* case "$hash":
                            var.TextValue = specForm.Hash;
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
                            break; */
                        case "$vpmoChief":
                            if (specForm.Attributes.CheckMilitaryChief)
                                var.TextValue = "СОГЛАСОВАНО\n\nНачальник 454 ВП МО РФ";
                            else var.TextValue = "";
                            break;

                        case "$chiefZGKZ":
                            if (specForm.Attributes.Zgkz)
                                var.TextValue = "ЗГКЗ";
                             else var.TextValue = "";
                            break;
                        case "$surnameZGKZ":
                            if (specForm.Attributes.Zgkz)
                            var.TextValue = specForm.Attributes.SurnameZgkz;
                             else var.TextValue = "";
                            break;
                        case "$vpmoChiefDate":
                            if (specForm.Attributes.CheckMilitaryChief)
                                var.TextValue = @"""__""_________20__ г.";
                            else var.TextValue = "";
                            break;
                        case "$brigChief":
                            if (specForm.Attributes.CheckChiefTeam)
                                var.TextValue = "Нач. РБ";
                            else var.TextValue = "";
                            break;
                        case "$vpmoVisa":
                            if (specForm.Attributes.CheckVPMO)
                                var.TextValue = "ВП МО РФ";
                            else var.TextValue = "";
                            break;
                        case "$litera":
                            var.TextValue = specForm.Attributes.Litera;
                            break;
                        case "$literalu":
                            if (specForm.Attributes.Litera == "О1")
                                var.TextValue = specForm.Attributes.Litera;
                            else var.TextValue = string.Empty;
                            break;
                        case "$vpmoline":
                            if (specForm.Attributes.CheckMilitaryChief)
                                var.TextValue = "_________";
                            else var.TextValue = "";
                            break;
                        case "$maker":
                            var.TextValue = "Исполнитель";
                            break;
                    }
                
                }

                
            }
        } // AddPageForms

        string GetSidebarText(SpecFormEspd specForm)
        {
            DocumentAttributes attrs = specForm.Attributes;
            string text = String.Format("{0}               {1}     {2}               {3}     {4}               {5}",
                attrs.SidebarSignHeader1, attrs.SidebarSignName1,
                attrs.SidebarSignHeader2, attrs.SidebarSignName2,
                attrs.SidebarSignHeader3, attrs.SidebarSignName3);
            return text.Trim();
        }
    }

    //
    //  Настройки
    //
    internal sealed class Settings
    {
        public static readonly string PathToTemplates = @"Служебные файлы\Шаблоны отчётов\Спецификация ГОСТ 19.202 (ЕСПД)\";
            // = @"D:\Шаблоны отчетов\";
            // = @"D:\Тест - отчет"; // Environment.GetEnvironmentVariable("ALLUSERSPROFILE") + @"\Application Data\T-FLEX DOCs 10\Prototypes\";
            // = @"S:\Projects\GlobusTFlexSuite\devel\Source\DocsReports\SpecificationReport\Templates";

        public static readonly string FormName = "Спецификация форма ГОСТ 19.202.grb";

        public static readonly string ChangelogName = "Лист регистрации изменений ГОСТ 19.604.grb";

        public static readonly string DateTimeFormat = "d.MM.yyyy H:mm:ss";        

        public static readonly string TitulListName = "Титульный лист ГОСТ 19.104.grb";

        //public static readonly string ListUtvergdeniaName = "Лист утверждения ГОСТ 19.104.grb";

        public static readonly string ListUtvergdeniaNameHard = "Лист утверждения ГОСТ 19.104.grb";
    }

    /// Строка таблицы спецификации по форме ГОСТ 19.202-78
    public class SpecRowFormEspd
    {
        // Графы
        SpecFormCell _designation;
        SpecFormCell _name;
        SpecFormCell _note;

        public SpecRowFormEspd(SpecFormCell designation, SpecFormCell name, SpecFormCell note)
        {
            _designation = designation;
            _name = name;
            _note = note;
        }

        public override string ToString()
        {
            return String.Format("|{0}|{1}|{2}|",
                                 _designation.Text,
                                 _name.Text,
                                 _note.Text);
        }

        /// Форматирование записи спецификации с разбивкой значений полей на строки
        public static SpecRowFormEspd[] CreateRecord(string designation, string name, string note,
                                                     TextFormatter textFormatter)
        {
            List<SpecRowFormEspd> rows = new List<SpecRowFormEspd>();

            // Разбивка некоторых полей на несколько строк
            string[] designationLines = textFormatter.Wrap(designation, SpecRowFormEspd.DesignationColumnWidth);
            string[] nameLines = textFormatter.WrapName(name, SpecRowFormEspd.NameColumnWidth);
            string[] noteLines = textFormatter.Wrap(note, SpecRowFormEspd.NoteColumnWidth);

            // Создание строк спецификации
            while (true)
            {
                bool empty = true;

                SpecFormCell designationCell;
                if (rows.Count < designationLines.Length)
                {
                    designationCell = new SpecFormCell(designationLines[rows.Count]);
                    empty = false;
                }
                else
                {
                    designationCell = SpecFormCell.Empty;
                }

                // Ячейка графы "Наименование"
                SpecFormCell nameCell;
                if (rows.Count < nameLines.Length)
                {
                    nameCell = new SpecFormCell(nameLines[rows.Count]);
                    empty = false;
                }
                else
                {
                    nameCell = SpecFormCell.Empty;
                }

                // Ячейка графы "Прим."
                SpecFormCell noteCell;
                if (rows.Count < noteLines.Length)
                {
                    noteCell = new SpecFormCell(noteLines[rows.Count]);
                    empty = false;
                }
                else
                {
                    noteCell = SpecFormCell.Empty;
                }

                // Прерывание цикла, если очередная строка оказалась пустой
                if (empty)
                    break;

                // Создание строки таблицы
                SpecRowFormEspd row = new SpecRowFormEspd(designationCell, nameCell, noteCell);
                rows.Add(row);
            }

            return rows.ToArray();
        }

        static readonly SpecRowFormEspd _empty = new SpecRowFormEspd(SpecFormCell.Empty,
                                                                     SpecFormCell.Empty,
                                                                     SpecFormCell.Empty);

        public static SpecRowFormEspd Empty
        {
            get { return _empty; }
        }

        public SpecFormCell Designation
        {
            get { return _designation; }
        }

        public SpecFormCell Name
        {
            get { return _name; }
        }

        public SpecFormCell Note
        {
            get { return _note; }
        }

        public static float DesignationColumnWidth
        {
            get { return 80f; }
        }

        public static float NameColumnWidth
        {
            get { return 70f; }
        }

        public static float NoteColumnWidth
        {
            get { return 30f; }
        }
    }

    public class SpecFormEspd
    {
        TextFormatter _textFormatter;

        // Атрибуты основной надписи

        string _name;
        string _designation;
        // string _documentGuid;
        string _litera;

        DocumentAttributes _attributes;

        // string _hash;

        // Строки таблицы
        List<SpecRowFormEspd> _rows;

        static readonly int _firstPageRows = 29; // Всего 30 строк, но одна пустая в форматке уже имеется
        static readonly int _otherPageRows = 30;

        public SpecFormEspd(SpecDocument document, DocumentAttributes attributes)
        {
            Font font = GetSpecTableFont();
            _textFormatter = new TextFormatter(font);
            _rows = new List<SpecRowFormEspd>();

            if (!(document is SpecDocumentEspd))
                throw new ApplicationException("Документ не является документом класса SpecDocumentEspd");

            _name = document.Name;
            _designation = document.Designation;
            _litera = document.Litera;
            //attributes.SidebarSignHeader1 = FormAttributes.AttrFromForm.
            _attributes = attributes;
            // _documentGuid = document.GUID;
            // _hash = hash;
        }

        Font GetSpecTableFont()
        {
            return new Font("Times New Roman", 3.3f, GraphicsUnit.Millimeter);
        }

        public void Add(SpecRowFormEspd row)
        {
            _rows.Add(row);
        }

        public void Add(SpecRowFormEspd[] rows)
        {
            foreach (SpecRowFormEspd row in rows)
                Add(row);
        }

        public void AddEmptyRow()
        {
            Add(SpecRowFormEspd.Empty);
        }

        public void AddEmptyRows(int count)
        {
            for (int i = 0; i < count; ++i)
            {
#if DEBUG_REPORT_LINES
                Add(new SpecRowFormEspd(new SpecFormCell(i.ToString()), new SpecFormCell(count.ToString())));
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
                Add(new SpecRowFormEspd(new SpecFormCell("a"), new SpecFormCell("")));
#else
                AddEmptyRow();
#endif
            while (!AtTopOfPage())
#if DEBUG_REPORT_LINES
                Add(new SpecRowFormEspd(new SpecFormCell("b" + _rows.Count.ToString()), new SpecFormCell(((_rows.Count - _firstPageRows) % _otherPageRows).ToString())));
#else
                AddEmptyRow();
#endif
        }

        public SpecRowFormEspd[] GetRows()
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

        public SpecRowFormEspd[] GetRows(int page)
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

        public DocumentAttributes Attributes
        {
            get { return _attributes; }
        }

        /* public string DocumentGuid
        {
            get { return _documentGuid; }
        } */

        public string Litera
        {
            get { return _litera; }
        }

        /* public string Hash
        {
            get { return _hash; }
        } */
    }

    public class SpecDocumentEspd : SpecDocument
    {
        // Параметры для спецификации

        // string[] _litera;         // литера
        // string _note;           // примечание

        /**
         *  Constructor
         */
        public SpecDocumentEspd(TFDDocument doc, SpecDocument parent)
            : base(doc, parent)
        {
        }

        public SpecDocumentEspd(TFDDocument doc, SpecDocument parent, DocumentAttributes documentAttr)
            : base(doc, parent, documentAttr)
        {
        }

        //
        //  Чтение параметров документа
        //
        protected override void ReadDocumentParameters(TFDDocument doc, SpecDocument parent)
        {
            // Чтение основных параметров (базовый класс)
            base.ReadDocumentParameters(doc, parent);
        }
    }

    //
    // Сравнение документов для определения их порядка записи в спецификацию.
    //
    internal class SpecDocumentComparerEspd : SpecDocumentComparer
    {
        public SpecDocumentComparerEspd()
            : base()
        { }

        /**
         * Сравнение двух документов спецификации, определяющей состав одного изделия.
         * Функция учитывает:
         * - раздел спецификцаии;
         * - правила сортировки записей спецификации для каждого из разделов.
         */
        public override int Compare(SpecDocument doc1, SpecDocument doc2)
        {
            int cmp = 0;

            SpecDocumentEspd d1 = (SpecDocumentEspd)doc1;
            SpecDocumentEspd d2 = (SpecDocumentEspd)doc2;

            // Сравнение по разделам спецификации

            cmp = CompareByBomGroup(d1, d2);
            if (cmp != 0)
                return cmp;

            // Сравнение по номеру позиции, обозначению и наименованию

            cmp = CompareByDesignationAndName(d1, d2);
            if (cmp != 0)
                return cmp;

            cmp = CompareByID(d1, d2);
            return cmp;
        }

        /**
         *  Сравнение на основе разделов спецификации
         */
        protected int CompareByBomGroup(SpecDocumentEspd doc1,
                                        SpecDocumentEspd doc2)
        {
            int cmp = doc1.BomGroup.CompareTo(doc2.BomGroup);
            return cmp;
        }

        /**
         *  Сравнение документов из одного раздела по графам спецификации:
         *  - порядковый номер позиции;
         *  - обозначение;
         *  - наименование;
         */
        protected int CompareByDesignationAndName(SpecDocumentEspd doc1,
                                                  SpecDocumentEspd doc2)
        {
            Debug.Assert(CompareByBomGroup(doc1, doc2) == 0);

            int cmp = 0;

            // Сравнение по обозначению
            cmp = CompareByStringElements(doc1.Designation, doc2.Designation);
            if (cmp != 0)
                return cmp;

            // Сравнение по наименованию
            cmp = CompareByStringElements(doc1.Name, doc2.Name, true);

            return cmp;
        }


        protected int CompareByID(SpecDocumentEspd doc1, SpecDocumentEspd doc2)
        {
            if (doc1.ID < doc2.ID)
                return -1;
            else if (doc1.ID > doc2.ID)
                return 1;
            else
                return 0;
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

        public void progressBarParam (int step)
        {
            progressBar1.Step=step;
            progressBar1.PerformStep();
        }
        public void setStage(Stages stage)
        {
		    progressBarParam(4);
            switch (stage)
            {
                case Stages.Initialization:
                    ProcessingLabel.Text= "Инициализация...";
                    break;
                case Stages.DataAcquisition:
                    ProcessingLabel.Text= "Получение данных";
                    break;
                case Stages.DataProcessing:
                    ProcessingLabel.Text= "Обработка данных";
                    break;
                case Stages.ReportGenerating:
                    ProcessingLabel.Text= "Создание отчёта";
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
/*namespace FormAttributes
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
    using DocumentAttributes;

    public partial class Attributes : Form
    {
        public Attributes()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.labelZagolovok2 = new System.Windows.Forms.Label();
            this.labelZagolovok1 = new System.Windows.Forms.Label();
            this.tbSideBarSignName3 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignName2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignName1 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader3 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader1 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbApprovedBy = new System.Windows.Forms.TextBox();
            this.tbNControlBy = new System.Windows.Forms.TextBox();
            this.tbControlledBy = new System.Windows.Forms.TextBox();
            this.tbDesignedBy = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.cbFullNames = new System.Windows.Forms.CheckBox();
            this.cbAddChangelogPage = new System.Windows.Forms.CheckBox();
            this.nudFirstPageNumber = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.makeReportButton = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.nudCountBefore = new System.Windows.Forms.NumericUpDown();
            this.nudCountAfter = new System.Windows.Forms.NumericUpDown();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.labelZagolovok2);
            this.groupBox1.Controls.Add(this.labelZagolovok1);
            this.groupBox1.Controls.Add(this.tbSideBarSignName3);
            this.groupBox1.Controls.Add(this.tbSideBarSignName2);
            this.groupBox1.Controls.Add(this.tbSideBarSignName1);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader3);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader2);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(548, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Визы на поле подшивки";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(277, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Фамилия";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(277, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Фамилия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(277, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Фамилия";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Заголовок";
            // 
            // labelZagolovok2
            // 
            this.labelZagolovok2.AutoSize = true;
            this.labelZagolovok2.Location = new System.Drawing.Point(46, 48);
            this.labelZagolovok2.Name = "labelZagolovok2";
            this.labelZagolovok2.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok2.TabIndex = 8;
            this.labelZagolovok2.Text = "Заголовок";
            // 
            // labelZagolovok1
            // 
            this.labelZagolovok1.AutoSize = true;
            this.labelZagolovok1.Location = new System.Drawing.Point(46, 22);
            this.labelZagolovok1.Name = "labelZagolovok1";
            this.labelZagolovok1.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok1.TabIndex = 7;
            this.labelZagolovok1.Text = "Заголовок";
            // 
            // tbSideBarSignName3
            // 
            this.tbSideBarSignName3.Location = new System.Drawing.Point(349, 71);
            this.tbSideBarSignName3.Name = "tbSideBarSignName3";
            this.tbSideBarSignName3.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName3.TabIndex = 6;
            // 
            // tbSideBarSignName2
            // 
            this.tbSideBarSignName2.Location = new System.Drawing.Point(349, 45);
            this.tbSideBarSignName2.Name = "tbSideBarSignName2";
            this.tbSideBarSignName2.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName2.TabIndex = 5;
            // 
            // tbSideBarSignName1
            // 
            this.tbSideBarSignName1.Location = new System.Drawing.Point(349, 19);
            this.tbSideBarSignName1.Name = "tbSideBarSignName1";
            this.tbSideBarSignName1.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName1.TabIndex = 4;
            // 
            // tbSideBarSignHeader3
            // 
            this.tbSideBarSignHeader3.Location = new System.Drawing.Point(117, 71);
            this.tbSideBarSignHeader3.Name = "tbSideBarSignHeader3";
            this.tbSideBarSignHeader3.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader3.TabIndex = 3;
            // 
            // tbSideBarSignHeader2
            // 
            this.tbSideBarSignHeader2.Location = new System.Drawing.Point(117, 45);
            this.tbSideBarSignHeader2.Name = "tbSideBarSignHeader2";
            this.tbSideBarSignHeader2.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader2.TabIndex = 2;
            // 
            // tbSideBarSignHeader1
            // 
            this.tbSideBarSignHeader1.Location = new System.Drawing.Point(117, 19);
            this.tbSideBarSignHeader1.Name = "tbSideBarSignHeader1";
            this.tbSideBarSignHeader1.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader1.TabIndex = 1;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.tbApprovedBy);
            this.groupBox2.Controls.Add(this.tbNControlBy);
            this.groupBox2.Controls.Add(this.tbControlledBy);
            this.groupBox2.Controls.Add(this.tbDesignedBy);
            this.groupBox2.Location = new System.Drawing.Point(12, 118);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(548, 129);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Подписи основной надписи";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(154, 100);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "Утв.";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(154, 74);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 11;
            this.label7.Text = "Н. контр.";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(154, 48);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(36, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Пров.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(154, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Разраб.";
            // 
            // tbApprovedBy
            // 
            this.tbApprovedBy.Location = new System.Drawing.Point(223, 97);
            this.tbApprovedBy.Name = "tbApprovedBy";
            this.tbApprovedBy.Size = new System.Drawing.Size(156, 20);
            this.tbApprovedBy.TabIndex = 8;
            // 
            // tbNControlBy
            // 
            this.tbNControlBy.Location = new System.Drawing.Point(223, 71);
            this.tbNControlBy.Name = "tbNControlBy";
            this.tbNControlBy.Size = new System.Drawing.Size(156, 20);
            this.tbNControlBy.TabIndex = 7;
            // 
            // tbControlledBy
            // 
            this.tbControlledBy.Location = new System.Drawing.Point(223, 45);
            this.tbControlledBy.Name = "tbControlledBy";
            this.tbControlledBy.Size = new System.Drawing.Size(156, 20);
            this.tbControlledBy.TabIndex = 6;
            // 
            // tbDesignedBy
            // 
            this.tbDesignedBy.Location = new System.Drawing.Point(223, 19);
            this.tbDesignedBy.Name = "tbDesignedBy";
            this.tbDesignedBy.Size = new System.Drawing.Size(156, 20);
            this.tbDesignedBy.TabIndex = 5;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.nudCountAfter);
            this.groupBox3.Controls.Add(this.nudCountBefore);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.cbFullNames);
            this.groupBox3.Controls.Add(this.cbAddChangelogPage);
            this.groupBox3.Controls.Add(this.nudFirstPageNumber);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Location = new System.Drawing.Point(12, 253);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(548, 105);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Нумерация листов и содержание";
            // 
            // cbFullNames
            // 
            this.cbFullNames.AutoSize = true;
            this.cbFullNames.Location = new System.Drawing.Point(13, 78);
            this.cbFullNames.Name = "cbFullNames";
            this.cbFullNames.Size = new System.Drawing.Size(143, 17);
            this.cbFullNames.TabIndex = 3;
            this.cbFullNames.Text = "Полные наименования";
            this.cbFullNames.UseVisualStyleBackColor = true;
            // 
            // cbAddChangelogPage
            // 
            this.cbAddChangelogPage.AutoSize = true;
            this.cbAddChangelogPage.Checked = true;
            this.cbAddChangelogPage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbAddChangelogPage.Location = new System.Drawing.Point(13, 55);
            this.cbAddChangelogPage.Name = "cbAddChangelogPage";
            this.cbAddChangelogPage.Size = new System.Drawing.Size(234, 17);
            this.cbAddChangelogPage.TabIndex = 2;
            this.cbAddChangelogPage.Text = "Добавлять лист регистрации изменений";
            this.cbAddChangelogPage.UseVisualStyleBackColor = true;
            // 
            // nudFirstPageNumber
            // 
            this.nudFirstPageNumber.Location = new System.Drawing.Point(148, 26);
            this.nudFirstPageNumber.Name = "nudFirstPageNumber";
            this.nudFirstPageNumber.Size = new System.Drawing.Size(45, 20);
            this.nudFirstPageNumber.TabIndex = 1;
            this.nudFirstPageNumber.Value = 1;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(10, 26);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(117, 13);
            this.label9.TabIndex = 0;
            this.label9.Text = "Номер первого листа";
            // 
            // makeReportButton
            // 
            this.makeReportButton.Location = new System.Drawing.Point(453, 364);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(107, 23);
            this.makeReportButton.TabIndex = 3;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(253, 26);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(190, 13);
            this.label10.TabIndex = 4;
            this.label10.Text = "Количество пустых строк до записи";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(253, 56);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(208, 13);
            this.label11.TabIndex = 5;
            this.label11.Text = "Количество пустых строк после записи";
            // 
            // nudCountBefore
            // 
            this.nudCountBefore.Location = new System.Drawing.Point(468, 24);
            this.nudCountBefore.Name = "nudCountBefore";
            this.nudCountBefore.Size = new System.Drawing.Size(66, 20);
            this.nudCountBefore.TabIndex = 6;
            this.nudCountBefore.Value = 1;
            // 
            // nudCountAfter
            // 
            this.nudCountAfter.Location = new System.Drawing.Point(467, 54);
            this.nudCountAfter.Name = "nudCountAfter";
            this.nudCountAfter.Size = new System.Drawing.Size(66, 20);
            this.nudCountAfter.TabIndex = 7;
            this.nudCountAfter.Value = 1;
            // 
            // AttributesDesign
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(572, 397);
            this.Controls.Add(this.makeReportButton);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "AttributesDesign";
            this.Text = "Настройки отчета";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelZagolovok2;
        private System.Windows.Forms.Label labelZagolovok1;
        private System.Windows.Forms.TextBox tbSideBarSignName3;
        private System.Windows.Forms.TextBox tbSideBarSignName2;
        private System.Windows.Forms.TextBox tbSideBarSignName1;
        private System.Windows.Forms.TextBox tbSideBarSignHeader3;
        private System.Windows.Forms.TextBox tbSideBarSignHeader2;
        private System.Windows.Forms.TextBox tbSideBarSignHeader1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbApprovedBy;
        private System.Windows.Forms.TextBox tbNControlBy;
        private System.Windows.Forms.TextBox tbControlledBy;
        private System.Windows.Forms.TextBox tbDesignedBy;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NumericUpDown nudFirstPageNumber;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox cbAddChangelogPage;
        private System.Windows.Forms.CheckBox cbFullNames;
        private System.Windows.Forms.Button makeReportButton;
        private System.Windows.Forms.NumericUpDown nudCountAfter;
        private System.Windows.Forms.NumericUpDown nudCountBefore;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private bool _attributesFormClose = false;

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            _attributesFormClose = true;
        }

        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }
        public bool AttributesFormClose
        {
            get { return _attributesFormClose; }
        }

        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();         

            while (AttributesFormClose==false)
            {
                this.UpdateUI();                
            }

            DocumentAttributes documentAttr = new DocumentAttributes();
            documentAttr.SidebarSignHeader1 = (tbSideBarSignHeader1.Text == null) ? "" : tbSideBarSignHeader1.Text;
            documentAttr.SidebarSignHeader2 = (tbSideBarSignHeader2.Text == null) ? "" : tbSideBarSignHeader2.Text;
            documentAttr.SidebarSignHeader3 = (tbSideBarSignHeader3.Text == null) ? "" : tbSideBarSignHeader3.Text;

            documentAttr.SidebarSignName1 = (tbSideBarSignName1.Text == null) ? "" : tbSideBarSignName1.Text;
            documentAttr.SidebarSignName2 = (tbSideBarSignName2.Text == null) ? "" : tbSideBarSignName2.Text;
            documentAttr.SidebarSignName3 = (tbSideBarSignName3.Text == null) ? "" : tbSideBarSignName3.Text;

            documentAttr.DesignedBy = (tbDesignedBy.Text == null) ? "" : tbDesignedBy.Text;
            documentAttr.ControlledBy = (tbControlledBy.Text == null) ? "" : tbControlledBy.Text;
            documentAttr.ApprovedBy = (tbApprovedBy.Text == null) ? "" : tbApprovedBy.Text;
            documentAttr.NControlBy = (tbNControlBy.Text == null) ? "" : tbNControlBy.Text;

            documentAttr.FirstPageNumber = Convert.ToInt32(nudFirstPageNumber.Value);

            if (cbAddChangelogPage.CheckState == CheckState.Checked)
                documentAttr.AddChangelogPage = true;
            else documentAttr.AddChangelogPage = false;

            if (cbFullNames.CheckState == CheckState.Checked)
                documentAttr.FullNames = true;
            else documentAttr.FullNames = false;

            documentAttr.CountEmptyLinesBefore = Convert.ToInt32(nudCountBefore.Value);

            documentAttr.CountEmptyLinesAfter = Convert.ToInt32(nudCountAfter.Value);

            this.Close();
                        
            // ТОЧКА ВХОДА МАКРОСА
            SpecificationReport.Espd.EspdSpecificationCadReport.Make(context,documentAttr);
        }
    }
}*/




