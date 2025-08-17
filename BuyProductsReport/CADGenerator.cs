using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using TFlex;
using TFlex.Model;
using TFlex.Model.Model2D;
using TFlex.Reporting;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;
using System.Drawing;
using System.Diagnostics;
using BuyProductsReport.Form1Namespace;
using System.Runtime.InteropServices;
using Globus.DOCs.Technology.Reports.CAD;

namespace BuyProductsReport
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
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionParamsForm selectionParamsForm = new SelectionParamsForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            context.CopyTemplateFile();  //Создаем копию шаблона
            ApiCad.InitApi(context);
            TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD
           
            TFlex.Model.Document document = null;    //Документ CAD 
            
            try
            {
                // ТОЧКА ВХОДА МАКРОСА
                selectionParamsForm.DotsEntry(context);
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
    }



    public class DocumentAttributes
    {
        // для вкладки Настройка страниц
        private bool addList = false;
        private bool newListSection = false;
        // для вкладки Добавление изменений
        private bool _addModify = false;  // флаг добавления изменений
        private string numberModify = string.Empty;                 // номер изменения
        private string numberDocument = string.Empty; // номер документа
        private string dateModify = string.Empty;           // дата изменения
        private List<int> listPage = new List<int>(); // список страниц
        private bool addTitulList = false;
        private string listZam = string.Empty;
        private string listNov = string.Empty;
        private List<int> listPage2 = new List<int>();

        public bool AddList
        {
            get { return addList; }
            set { addList = value; }
        }

        public bool NewListSection
        {
            get { return newListSection; }
            set { newListSection = value; }
        }

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

    // ============== ТЕКСТ МАКРОСА ВП T-FLEX DOCs 11 ==========================
    /// <summary>
    /// Построитель отчета
    /// </summary>
    public delegate void LogDelegate(string line);

    public class BuyListReport
    {
        static public Form1 m_form;
        static public LogDelegate m_writeToLog;
        
        /// <summary>
        /// Точка запуска отчета
        /// </summary>
        [STAThread]
        public static void MakeBuyListReport(IReportGenerationContext context, DocumentAttributes documentAttr)
        {
            BuyListReport buyListReport = new BuyListReport();
            buyListReport.MakeReport(context, documentAttr);
        }

        public void MakeReport(IReportGenerationContext context, DocumentAttributes documentAttr)
        {
            m_form = new Form1();
            m_writeToLog = new LogDelegate(m_form.writeToLog);

            m_form.Show();
            m_form.setStage(Form1Namespace.Form1.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            try
            {
              //формируем данные
                m_writeToLog("=== Получение идентификатора корневого документа ===");
                int documentId = ApiCad.GetBaseDocID(context);
                if (ApiCad.GetBaseDocID(context) == -1) return;

                ReportData reportData = ApiDocs.GetDocsReportData(documentId, context);

                //получаем наименование, обозначение и разработчика документа, созданного на рабочем столе для отчета
                m_form.Refresh();
                ReportParams reportParameter = new ReportParams(reportData.BaseDoc);
                m_form.Refresh();

                //заполняем шаблон отчета данными
                ApiCad.BuildReport(reportParameter, reportData, documentAttr);
                System.Diagnostics.Process.Start(context.ReportFilePath);
            }
            catch (System.TypeInitializationException)
            {
                System.Windows.Forms.MessageBox.Show(
                "Неправильно заданы дифайны!\n" +
                "Для сборки в Редакторе макросов T-Flex CAD следует установить #undef DEBUG!\n" +
                "Для отладки в Visual Studio 2005 cледует установить #define DEBUG!\n\n" +
                "Измените дифайны и перекомпилируйте отчет!", "Ошибка!",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                if (ApiDocs.halt == false)
                    System.Windows.Forms.MessageBox.Show("В процессе формирования отчета возникла ошибка:\n" + ex.Message, "Ошибка!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
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
            text = text.Trim();
            Size size = TextRenderer.MeasureText(text, font);
            double width = 25.4f * widthFactor * (double)size.Width / (double)_displayResolutionX;
            return width;
        }

        public string[] Wrap(string text, double maxWidth)
        {
            return Wrap(text, _font, maxWidth);
        }

        /// Разбивка на строки по символам перевода строки \n
        public string[] Wrap(string text, Font font, double maxWidth)
        {
            string[] lines = text.Split(new string[] { @"\n" }, StringSplitOptions.None);

            var wrappedLines = new List<string>();
            foreach (string line in lines)
            {
                string[] wrappedBySpaces = WrapBySyllables(line, font, maxWidth);
                wrappedLines.AddRange(wrappedBySpaces);
            }
            return wrappedLines.ToArray();
        }
  
        public string[] Wrap(string text)
        {
            double maxWidth = 57f;
            var wrappedLines = new List<string>();

            string patternBeginPart = @"^(\S{2,}(\s+(?:[А-Я]{1})?[а-яё]{2,}(?=\s|$))*)";
            Regex regex;
            Match match;
            string beginPart = String.Empty;
            string remPart = String.Empty;
           
            if (GetTextWidth(text, _font) > maxWidth)
            {
                regex = new Regex(patternBeginPart);
                match = regex.Match(text);
                beginPart = match.Value;
                beginPart = beginPart.Trim();
                if (GetTextWidth(beginPart, _font) > maxWidth)
                    wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                else wrappedLines.Add(beginPart);
              
                remPart = text.Remove(0, beginPart.Length);
                remPart = remPart.Trim();

                if (remPart.Trim() != String.Empty)
                {
                    if (GetTextWidth(remPart, _font) > maxWidth)
                        wrappedLines.AddRange(Wrap(remPart, maxWidth));
                    else
                    {
                        remPart = remPart.Trim();
                        wrappedLines.Add(remPart);
                    }
                }
            }      
            else wrappedLines.Add(text); 
          
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

            var lines = new List<string>();
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

    #region Domain

    internal class TreeNode<TValue>
    {
        private TValue _value;
        private TreeNode<TValue> _parentNode;
        private List<TreeNode<TValue>> _childNodes;
        private TreeNode<TValue> _rootNode;

        public TreeNode(TValue value)
        {
            _value = value;
            _childNodes = new List<TreeNode<TValue>>();
            _rootNode = this;
        }

        public TreeNode(TValue value, TreeNode<TValue> parentNode)
            : this(value)
        {
            _parentNode = parentNode;
            _parentNode._childNodes.Add(this);
            _rootNode = parentNode._rootNode;
        }

        public TValue Value
        {
            get { return _value; }
        }

        public TreeNode<TValue> ParentNode
        {
            get { return _parentNode; }
            private set { _parentNode = value; }
        }

        internal TreeNode<TValue> RootNode
        {
            get { return _rootNode; }
        }

        public List<TreeNode<TValue>> ChildNodes
        {
            get { return _childNodes; }
        }

        public bool HasChildrens
        {
            get { return _childNodes.Count > 0; }
        }
    }

    internal class DocumentTreeNode : TreeNode<TFDDocument>
    {
        private DocParams _document;

        public DocumentTreeNode(TFDDocument docParams)
            : base(docParams)
        {
            _document = new DocParams(docParams, this);
        }

        public DocumentTreeNode(TFDDocument docParams, DocumentTreeNode parentNode)
            : base(docParams, parentNode)
        {
            _document = new DocParams(docParams, this);
        }

        public new DocumentTreeNode ParentNode
        {
            get { return base.ParentNode as DocumentTreeNode; }
        }

        public DocParams Document
        {
            get { return _document; }
        }

        public bool PathContains(TFDDocument doc)
        {
            if (Value.ObjectID == doc.ObjectID)
                return true;

            if (ParentNode == null)
                return false;

            return ParentNode.PathContains(doc);
        }
    }

    internal class DocumentEntrance : IComparable<DocumentEntrance>
    {
        private DocParams _source;
        private double _naIzdelie;
        private double _vKomplekt;
        private string _kudaVhoditOboznachenie;

        public DocumentEntrance(DocParams doc)
        {
            _source = doc;
            FillCounts(doc);
            TFDDocument ownerSpecification = GetOwnerSpecification(doc);
            _kudaVhoditOboznachenie = ownerSpecification == null ? String.Empty : ownerSpecification.Oboznachenie;
        }

        public double NaIzdelie
        {
            get { return _naIzdelie; }
        }

        public double VKomplekt
        {
            get { return _vKomplekt; }
        }

        public double Vsego
        {
            get { return NaIzdelie + VKomplekt; }
        }

        public string KudaVhoditOboznachenie
        {
            get { return _kudaVhoditOboznachenie; }
        }

        public int CompareTo(DocumentEntrance other)
        {
            int comapreResult = _source.CompareTo(other._source);
            if (comapreResult != 0)
                return comapreResult;

            return _kudaVhoditOboznachenie.CompareTo(other.KudaVhoditOboznachenie);
        }

        private void FillCounts(DocParams doc)
        {
            if (doc.IsKomplekt)
                _vKomplekt = doc.GetSummaryCount();
            else
                _naIzdelie = doc.GetSummaryCount();
        }
        private TFDDocument GetOwnerSpecification(DocParams doc)
        {
            if (doc.Node.ParentNode == null)
                return null;
            return doc.Node.ParentNode.Document.Source;
        }

        public void MergeWith(DocumentEntrance other)
        {
            _naIzdelie += other.NaIzdelie;
            _vKomplekt += other.VKomplekt;
        }
    }

    /// <summary>
    /// Параметры документа DOCs
    /// </summary>
    internal class DocParams : IComparable<DocParams>
    {
        private const string DocCodeRegEx = @"\s(?<code>\w+\.([0-9]+\.)+\s?[-|\w|/]*)$";
        private const string DocCodeRegTUEnd = @"(\S+)?\s?(МТУ|ТУ|ПС)+((\/)?[0-9]{0,3})?(\s|$)?(?:\)[$\n])?";
        private const string DocCodeRegTUStart = @"\s(ТУ-У|ТУ|ГОСТ)\s?[\w+\.-]+";
        private static char[] DocNameSeparator = new char[] { ' ' };
        private int idBaseDoc = Defines.BaseDoc.ObjectId;
        private TFDDocument _source;
        private int _id;
        private string _naimenovanie;
        private string _shortNaimenovanie;
        private string _oboznachenieNaPostavku;
        private string _productionCode;
        private DocumentTreeNode _node;
        private bool? _isKomplekt;
        private string[] _nameStringArray;


        public DocParams(TFDDocument doc, DocumentTreeNode path)
        {
            if (doc == null)
                throw new ArgumentNullException("doc");
            if (path == null)
                throw new ArgumentNullException("path");

            _source = doc;
            _node = path;

            int id = _source.ObjectID;
            string naimenovanie = _source.Naimenovanie;
            string oboznachenie = _source.Oboznachenie;
            string productCode = _source.ProductCode;
            _oboznachenieNaPostavku = String.Empty;
            _naimenovanie = naimenovanie;
            _id = id;

            System.Text.RegularExpressions.Match match1 = System.Text.RegularExpressions.Regex.Match(naimenovanie, DocCodeRegTUStart);
            if (match1.Success)
            {
                _oboznachenieNaPostavku = match1.Value;
                _naimenovanie = naimenovanie.Replace(match1.Value, String.Empty).Trim();
            }

            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(_naimenovanie, DocCodeRegEx);
            if (match.Success)
            {
                _oboznachenieNaPostavku = _oboznachenieNaPostavku + match.Groups["code"].Value;
                _naimenovanie = _naimenovanie.Replace(match.Value, String.Empty).Trim();
            }

            System.Text.RegularExpressions.Match match2 = System.Text.RegularExpressions.Regex.Match(_naimenovanie, DocCodeRegTUEnd);
            if (match2.Success)
            {
                _oboznachenieNaPostavku = _oboznachenieNaPostavku + match2.Value;
                _naimenovanie = _naimenovanie.Replace(match2.Value, String.Empty).Trim();
            }

            _shortNaimenovanie = _naimenovanie;
            _naimenovanie = _naimenovanie + ' ' + oboznachenie;
            _oboznachenieNaPostavku = _oboznachenieNaPostavku.Trim();

            _productionCode = productCode; // ApiDocs.GetTableStringParamValue(Defines.ProductCodesGroup.GROUP.ID, Defines.ProductCodesGroup.ProductCode.ID, Source.ProductCode, false).Trim();
            _nameStringArray = _naimenovanie.Split(DocNameSeparator);
        }

        public double GetSummaryCount()
        {
            DocParams currentDocument = this;
            double summaryCount = 1;            

            do
            {
                summaryCount *= currentDocument.Source.Amount; // * currentDocument.Source.InstanceAmount);
                currentDocument = currentDocument.Node.ParentNode.Document;

            } while (summaryCount != 0 && currentDocument != null && currentDocument.Node.ParentNode != null);

            return summaryCount;
        }

        internal DocumentTreeNode Node
        {
            get { return _node; }
        }

        /// <summary>
        /// Наименование
        /// </summary>
        public string Naimenovanie
        {
            get { return _naimenovanie; }
        }

        /// <summary>
        /// Обозначение на поставку
        /// </summary>
        public string OboznachenieNaPostavku
        {
            get { return _oboznachenieNaPostavku; }
        }

        /// <summary>
        /// Код продукции
        /// </summary>
        public string ProductionCode
        {
            get { return _productionCode; }
        }

        public TFDDocument Source
        {
            get { return _source; }
        }

        public string[] NameStringArray
        {
            get { return _nameStringArray; }
        }

        public bool IsKomplekt
        {            
            get
            {                
                if (!_isKomplekt.HasValue)
                {
                    if (Source.Class == Defines.TFDClass.Komplekt)
                    {
                        _isKomplekt = true;
                    }
                    else
                    {
                        DocumentTreeNode parentNode = Node.ParentNode;  
                        _isKomplekt = (parentNode == null) ? false : (parentNode.Document.IsKomplekt && parentNode.Document._id != idBaseDoc);             
                    }
                }
                return _isKomplekt.Value;
            }
            set { _isKomplekt = value; }
        }

        public int CompareTo(DocParams other)
        {
            int compareResult = this.Source.Naimenovanie.CompareTo(other.Source.Naimenovanie);
            if (compareResult != 0)
            {
                int compareResult2 = this.Source.Oboznachenie.CompareTo(other.Source.Oboznachenie);
                if (compareResult2 != 0) return compareResult;
            }

            compareResult = this.Source.Oboznachenie.CompareTo(other.Source.Oboznachenie);
            if (compareResult != 0)
                return compareResult;

            compareResult = this.Source.Naimenovanie.CompareTo(other.Source.Naimenovanie);
            if (compareResult != 0)
                return compareResult;

            compareResult = this.OboznachenieNaPostavku.CompareTo(other.OboznachenieNaPostavku);
            if (compareResult != 0)
                return compareResult;

            return this.Source.ProductCode.CompareTo(other.Source.ProductCode);
        }
    }

    internal class TFDDocument
    {
        /* Порядок полей в запросе чтения документа */
        public int PK;
        public int ObjectID;
        public int ParentID;
        public int Class;
        
        public double Amount;
        public string Naimenovanie;
        public string Oboznachenie;
        public string ProductCode;        
        public string BomGroup;

        private List<TFDDocument> _childDocuments;        

        public TFDDocument(int id, int parent, int docClass)
        {
            ObjectID = id;
            ParentID = parent;
            Class = docClass;            
        }     

        public TFDDocument(ReferenceObject doc)
        {
            NomenclatureReference refNomenclature = new NomenclatureReference();
            TFlex.DOCs.Model.Structure.ParameterInfo idInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.ObjectId);
            TFlex.DOCs.Model.Structure.ParameterInfo classInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
            TFlex.DOCs.Model.Structure.ParameterInfo nameInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
            TFlex.DOCs.Model.Structure.ParameterInfo denotationInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);            

            ObjectID = (int)doc[idInfo].Value;
            ParentID = 0;
            Class = (int)doc[classInfo].Value;
            Amount = 1;
            Naimenovanie = (string)doc[nameInfo].Value;
            Oboznachenie = (string)doc[denotationInfo].Value;
            BomGroup = String.Empty;
        }

        public List<TFDDocument> ChildDocuments
        {
            get
            {
                if (_childDocuments == null)
                {
                        int parentDocId = ObjectID;
                    if (parentDocId > 0 && ApiDocs.DocumentsByParentId.ContainsKey(parentDocId))
                        _childDocuments = ApiDocs.DocumentsByParentId[parentDocId];
                    else
                        _childDocuments = new List<TFDDocument>();
                }
                return _childDocuments;
            }
        }
    }

    internal class NamedDocumentGroup : List<DocParams>
    {
        private string _name;

        public NamedDocumentGroup(string name)
        {
            _name = name;
        }

        public NamedDocumentGroup(string name, IEnumerable<DocParams> collection)
            : base(collection)
        {
            _name = name;
        }

        public string Name
        {
            get { return _name; }
        }

        public List<List<DocParams>> GroupRows()
        {
            var resultList = new List<List<DocParams>>();

            if (Count == 0)
                return resultList;

            Sort();
            List<DocParams> lastList = new List<DocParams>();
            resultList.Add(lastList);
            lastList.Add(this[0]);
            for (int rowIndex = 1; rowIndex < Count; rowIndex++)
            {
                DocParams prevDoc = this[rowIndex - 1];
                DocParams currDoc = this[rowIndex];
       
                if (currDoc.CompareTo(prevDoc) == 0)
                    lastList.Add(currDoc); 
                else
                {
                    lastList = new List<DocParams>();
                    lastList.Add(currDoc);
                    resultList.Add(lastList);
                 }
            }
            return resultList;
        }
    }

    internal class ReportData
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

        private List<NamedDocumentGroup> _groups;

        private ReferenceObject _baseDoc;
        private ReferenceObject _notClassKomplDoc;

        public ReportData(ReferenceObject baseReportDoc, IReportGenerationContext context)
        {
            if (baseReportDoc == null)
                throw new ArgumentNullException("parentDocument");

            _baseDoc = baseReportDoc;
            _notClassKomplDoc = baseReportDoc;
     
            List<DocParams> _allReportDocuments;
            //определяем, какие документы будут использованы для отчета
            DocParams _masterDocument = LoadDocumentStruct(baseReportDoc, out _allReportDocuments,context);

            if (ApiDocs.halt == false)
            {
                _groups = GroupReportData(_allReportDocuments);
            }
        }

        public Dictionary<string, KeyValueList<RowType, string[]>> GetReportRows(string baseDocCode)
        {
            var resultDict = new Dictionary<string, KeyValueList<RowType, string[]>>();
            int columnsCount = 10;
            
            for (int groupIndex = 0; groupIndex < _groups.Count; groupIndex++)
            {
                NamedDocumentGroup group = _groups[groupIndex];
                if (!resultDict.ContainsKey(group.Name))
                    resultDict.Add(group.Name, new KeyValueList<RowType, string[]>());
                KeyValueList<RowType, string[]> groupList = resultDict[group.Name];

                string[] headerRow = new string[columnsCount];
                headerRow[0] = String.Concat(group.Name, "\n");
                groupList.Add(RowType.Header, headerRow);

                List<List<DocParams>> groupedRows = group.GroupRows();
                foreach (List<DocParams> rows in groupedRows)
                {
                    DocParams rowData = rows[0];
                    List<DocumentEntrance> entrances = GetDocumentGroupEntrances(rows , _baseDoc);
                    entrances.Sort();

                    string[] firstRow = new string[columnsCount];
                    firstRow[0] = rowData.Naimenovanie;
                    firstRow[1] = rowData.ProductionCode;
                    firstRow[2] = rowData.OboznachenieNaPostavku;
                    
                    double vsego = 0;
                    if (entrances.Count > 0)
                    {
                        DocumentEntrance firstEntrance = entrances[0];
                        if (String.Compare(firstEntrance.KudaVhoditOboznachenie, baseDocCode, true) != 0)
                            firstRow[4] = firstEntrance.KudaVhoditOboznachenie;
                        firstRow[5] = GetNonZeroString(firstEntrance.NaIzdelie);
                        firstRow[6] = GetNonZeroString(firstEntrance.VKomplekt);
                        firstRow[8] = GetNonZeroString(firstEntrance.Vsego);
                        vsego += firstEntrance.Vsego;
                    }
                    groupList.Add(RowType.FirstRecordRow, firstRow);

                    if (entrances.Count < 2)
                        continue;
                    for (int entranceIndex = 1; entranceIndex < entrances.Count; entranceIndex++)
                    {
                        DocumentEntrance entrance = entrances[entranceIndex];
                        string[] entrancetRow = new string[columnsCount];
                        if (String.Compare(entrance.KudaVhoditOboznachenie, baseDocCode, true) != 0)
                            entrancetRow[4] = entrance.KudaVhoditOboznachenie;
                        entrancetRow[0] = String.Empty;
                        entrancetRow[1] = String.Empty;
                        entrancetRow[2] = String.Empty;
                        
                        entrancetRow[5] = GetNonZeroString(entrance.NaIzdelie);
                        entrancetRow[6] = GetNonZeroString(entrance.VKomplekt);

                        if (entranceIndex == entrances.Count - 1)
                            entrancetRow[8] = String.Concat(" ", GetNonZeroString(entrance.Vsego), " ");
                        else
                            entrancetRow[8] = GetNonZeroString(entrance.Vsego);

                        vsego += entrance.Vsego;
                        groupList.Add(entranceIndex == entrances.Count - 1 ? RowType.RowBeforeSummary : RowType.Row, entrancetRow);
                    }
                    string[] summaryRow = new string[columnsCount];
                    summaryRow[8] = vsego.ToString();
                    groupList.Add(RowType.Summary, summaryRow);
                }
                if (groupIndex < _groups.Count - 1)
                    groupList.Add(new KeyValuePair<RowType, string[]>(RowType.Empty, new string[columnsCount]));
            }
            return resultDict;
        }

        public ReferenceObject BaseDoc
        {
            get 
            {
                if (_baseDoc.Class.ToString().Trim() == "Комплект")
                    return _notClassKomplDoc;
                else
                    return _baseDoc; 
            }
        }

        private static List<DocumentEntrance> GetDocumentGroupEntrances(List<DocParams> rows, ReferenceObject baseDoc)
        {
            var resultEntrances = new SortedDictionary<string, DocumentEntrance>();
            foreach (DocParams doc in rows)
            {
                 DocumentEntrance docEntrance = new DocumentEntrance(doc);
                if (resultEntrances.ContainsKey(docEntrance.KudaVhoditOboznachenie))
                {
                    resultEntrances[docEntrance.KudaVhoditOboznachenie].MergeWith(docEntrance);
                }
                else
                    resultEntrances.Add(docEntrance.KudaVhoditOboznachenie, docEntrance);
            }
            return new List<DocumentEntrance>(resultEntrances.Values);
        }

        private DocParams LoadDocumentStruct(ReferenceObject baseDocument, out List<DocParams> resultList, IReportGenerationContext context)
        {
            NomenclatureReference refNomenclature = new NomenclatureReference();
            TFlex.DOCs.Model.Structure.ParameterInfo idInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.ObjectId);
            int idObject = (int) baseDocument[idInfo].Value;            
            resultList = new List<DocParams>();
            TFDDocument baseDocToProcess;
            BuyListReport.m_form.setStage(Form1Namespace.Form1.Stages.DataAcquisition);
            BuyListReport.m_writeToLog("=== Получение данных ===");
            ReadReportDocuments(idObject,context);
            BuyListReport.m_writeToLog("=== Добавление покупных изделий ===");
            baseDocToProcess = ApiDocs.DocumentsById[0];
            DocumentTreeNode path = new DocumentTreeNode(baseDocToProcess);
              //загружаем дерево потомков головного документа на 1 уровень
            if (ApiDocs.halt == false)
            {
                FillChildDocumentsTree(path, resultList);
                BuyListReport.m_form.progressBarParam(26);
                BuyListReport.m_form.setStage(Form1Namespace.Form1.Stages.DataProcessing);
                Clipboard.SetText(strB.ToString());
                MessageBox.Show(strB.ToString());
                BuyListReport.m_writeToLog("=== Обработка данных ===");
            }
            return new DocParams(baseDocToProcess, path);
        }

        private void ReadReportDocuments(int documentId, IReportGenerationContext context)
        {

            using (SqlConnection connection = ApiDocs.GetConnection(true,context))
            {         
                int complexClassID = Defines.TFDClass.Complex;                   // Комплекс
                int assemblyClassID = Defines.TFDClass.Assembly;                 // Сборочная единица
                int detalClassID = Defines.TFDClass.Detal;                       // Деталь
                int standartItemClassID = Defines.TFDClass.StandartItem;         // Стандартное изделие
                int otherItemClassID = Defines.TFDClass.OtherItem;               // Прочее изделие
                int materialClassID = Defines.TFDClass.Material;                 // Материал
                int komplektClassID = Defines.TFDClass.Komplekt;                 // Комплект
                int complexProgramClassID = Defines.TFDClass.ComplexProgram;     // Комплекс (программы)
                int componentProgramClassID = Defines.TFDClass.ComponentProgram; // Компонент (программы)                
                                
                //подготавливаем дерево документов хранилища
                //получение s_objectID всех объектов структуры изделия на каждом уровне
                //без повторения во временную таблицу #TmpTree
                string prepareCollectionCommandText = String.Format(@"
                DECLARE @docid int 
                SET @docid = {0}
                DECLARE @InsertCount int   
                SET @InsertCount = 0
				
                IF OBJECT_ID('tempdb..#TmpTree') is null
                   CREATE TABLE #TmpTree (s_ObjectID int, [level] int)
                ELSE
                DELETE FROM #TmpTree 
				                
                INSERT INTO #TmpTree
                SELECT n.s_ObjectID,0
                FROM Nomenclature n
                WHERE n.s_ObjectID = @docid
                  AND n.s_Deleted = 0 
                  AND n.s_ActualVersion = 1 

                DECLARE @level int
                SET @level = 0

                WHILE 1=1
                BEGIN
                  INSERT INTO #TmpTree 
                  SELECT DISTINCT nh.s_ObjectID, @level+1  
                  FROM NomenclatureHierarchy nh 
                  INNER JOIN #TmpTree ON nh.s_ParentID = #TmpTree.s_ObjectID 
                         AND #TmpTree.[level] = @level
                         AND nh.s_ActualVersion = 1
                         AND nh.s_Deleted = 0          
                  INNER JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID 
                         AND n.s_ActualVersion = 1
                         AND n.s_Deleted = 0
                         AND n.s_ClassID IN ({1},{2},{3},{4},{5},{6},{7},{8},{9}) 
                         --695 - Комплекс
                         --693 - Сборочная единица
                         --692 - Деталь
                         --703 - Стандартное изделие
                         --759 - Прочее изделие
                         --859 - Материал
                         --694 - Комплект
                         --790 - Комплекс (программы)
                         --789 - Компонент (программы)                         
                  SET  @InsertCount = @@ROWCOUNT 
                  SET @level = @level + 1 
                    	  
                  IF  @InsertCount = 0  
                      GOTO END1
                END             	
                END1:",documentId, complexClassID, assemblyClassID, detalClassID, standartItemClassID,
                                   otherItemClassID, materialClassID, komplektClassID, complexProgramClassID,
                                   componentProgramClassID);
                SqlCommand prepareCollectionCommand = new SqlCommand(prepareCollectionCommandText, connection);
                prepareCollectionCommand.CommandTimeout = 0;
                BuyListReport.m_writeToLog(String.Format("===Обработка 1-го запроса к базе данных==="));
                prepareCollectionCommand.ExecuteNonQuery();
                BuyListReport.m_form.progressBarParam(3);              

                //читаем необходимые параметры для ведомости покупных изделий объектов #TmpTree
                string getDocTreeCommandText = String.Format(
                  @"SELECT DISTINCT n.s_ClassID,
                           nh.s_ParentID,
                           n.s_ObjectID,
	                       nh.Amount,
                           n.Denotation,
	                       n.[Name],
                           nh.BomSection,
                           nh.s_PK,
                           bdp.CodeRCP
                    FROM NomenclatureHierarchy nh  
                    JOIN #TmpTree ON nh.s_ObjectID = #TmpTree.s_ObjectID
                    JOIN #TmpTree t ON nh.s_ParentID = t.s_ObjectID
                    JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                    LEFT JOIN BuyDeviceParams bdp ON  nh.s_ObjectID = bdp.s_ObjectID
                    WHERE n.s_Deleted = 0
                      AND n.s_ActualVersion = 1
                      AND nh.s_Deleted = 0
                      AND nh.s_ActualVersion = 1
                      AND n.s_ClassID  IN ({0},{1},{2},{3},{4},{5},{6},{7},{8}) 
                      AND bdp.s_ActualVersion = 1
                      AND bdp.s_Deleted = 0
                    UNION ALL 
 SELECT DISTINCT n.s_ClassID,
                           nh.s_ParentID,
                           n.s_ObjectID,
	                       nh.Amount,
                           n.Denotation,
	                       n.[Name],
                           nh.BomSection,
                           nh.s_PK,
                           bdp.CodeRCP
                    FROM NomenclatureHierarchy nh  
                    JOIN #TmpTree ON nh.s_ObjectID = #TmpTree.s_ObjectID
                    JOIN #TmpTree t ON nh.s_ParentID = t.s_ObjectID
                    JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                    LEFT JOIN BuyDeviceParams bdp ON  nh.s_ObjectID = bdp.s_ObjectID
                    WHERE n.s_Deleted = 0
                      AND n.s_ActualVersion = 1
                      AND nh.s_Deleted = 0
                      AND nh.s_ActualVersion = 1
                      AND n.s_ClassID  IN ({0},{1},{2},{3},{4},{5},{6},{7},{8}) 
                      AND nh.s_ObjectID not in (SELECT DISTINCT 
                           n.s_ObjectID
	                      
                    FROM NomenclatureHierarchy nh  
                    JOIN #TmpTree ON nh.s_ObjectID = #TmpTree.s_ObjectID
                    JOIN #TmpTree t ON nh.s_ParentID = t.s_ObjectID
                    JOIN Nomenclature n ON nh.s_ObjectID = n.s_ObjectID
                    LEFT JOIN BuyDeviceParams bdp ON  nh.s_ObjectID = bdp.s_ObjectID
                    WHERE n.s_Deleted = 0
                      AND n.s_ActualVersion = 1
                      AND nh.s_Deleted = 0
                      AND nh.s_ActualVersion = 1
                      AND n.s_ClassID  IN ({0},{1},{2},{3},{4},{5},{6},{7},{8}) 
                      AND bdp.s_ActualVersion = 1
                      AND bdp.s_Deleted = 0)
                    UNION ALL 
                    SELECT n1.s_ClassID,
                           0,
                           #TmpTree.s_ObjectID,
                           1,
                           n1.Denotation,
                           n1.[Name],
                           '',
                           0,
                           NULL
                    FROM #TmpTree 
                    JOIN Nomenclature n1 ON #TmpTree.s_ObjectID = n1.s_ObjectID
                    WHERE #TmpTree.[level] = 0
                      AND n1.s_Deleted = 0
                      AND n1.s_ActualVersion = 1", complexClassID, assemblyClassID, detalClassID, standartItemClassID,
                                   otherItemClassID, materialClassID, komplektClassID, complexProgramClassID,
                                   componentProgramClassID);
                SqlCommand docTreeCommand = new SqlCommand(getDocTreeCommandText, connection);
                docTreeCommand.CommandTimeout = 0;
                BuyListReport.m_writeToLog(String.Format("===Обработка 2-го запроса к базе данных==="));
                SqlDataReader reader = docTreeCommand.ExecuteReader();
                BuyListReport.m_form.progressBarParam(3);
                int PK=1;
                ApiDocs.DocumentsById.Clear();
                ApiDocs.DocumentsByParentId.Clear();
                while (reader.Read())
                {
                    int docId = reader.GetInt32(2);
                    int parentID = GetSqlInt32(reader, 1);
                    int objectClass = GetSqlInt32(reader, 0);
                    TFDDocument doc = new TFDDocument(docId, parentID, objectClass);
                    doc.Amount = GetSqlDouble(reader, 3);
                    doc.Oboznachenie = GetSqlString(reader, 4);

                    if ((GetSqlString(reader, 5) == "Комплект запасных частей на фильтр 8Д2.966.603 согласно паспорту 8Д2.966.603 ПС"))
                        doc.Naimenovanie = "Комплект запасных частей на апрпар варвар кенекнек вапвапавп 8Д2.966.603 согласно паспорту 8Д2.966.603 ПС";
                    else 
                    doc.Naimenovanie = GetSqlString(reader, 5);
                    doc.BomGroup = GetSqlString(reader, 6);
                    PK = GetSqlInt32(reader, 7);  
                    doc.PK = PK;
                    if (GetSqlString(reader, 8) != null)
                        doc.ProductCode = GetSqlString(reader, 8);
                    else doc.ProductCode = string.Empty;

                    if (ApiDocs.DocumentsById.ContainsKey(PK))
                       continue;
                    else ApiDocs.DocumentsById.Add(PK, doc);
                    
                    if (ApiDocs.DocumentsByParentId.ContainsKey(parentID))
                    {
                        ApiDocs.DocumentsByParentId[parentID].Add(doc);
                    }
                    else
                    {
                        var childList = new List<TFDDocument>();
                        childList.Add(doc);
                        ApiDocs.ChildDocList.Add(doc);
                        ApiDocs.DocumentsByParentId.Add(parentID, childList);
                    }
                }
                reader.Close();
            }
        }

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

        private void FillChildDocumentsTree(DocumentTreeNode parentDocumentNode, List<DocParams> resultList)
        {                       
            foreach (TFDDocument childDocument in parentDocumentNode.Document.Source.ChildDocuments)
            {
                var childDocNode = new DocumentTreeNode(childDocument, parentDocumentNode);
                //обрабатываем только прочие изделия
                if (childDocument.Class == Defines.TFDClass.OtherItem)
                {   
                    resultList.Add(new DocParams(childDocument, childDocNode));
                    BuyListReport.m_writeToLog(String.Format("Добавлено покупное изделие: {0}", childDocument.Naimenovanie));
                    continue;
                }

                if (childDocument.Class == Defines.TFDClass.StandartItem && childDocument.Naimenovanie.ToLower().Contains("рукав металлический"))
                {
                    resultList.Add(new DocParams(childDocument, childDocNode));
                    BuyListReport.m_writeToLog(String.Format("Добавлено стандартное изделие: {0}", childDocument.Naimenovanie));
                    strB.AppendLine(childDocument.Naimenovanie + " " + childDocument.Oboznachenie + " " + childDocument.Amount + " " + childDocument.Class + " " + childDocument.BomGroup + " " + childDocument.ParentID);
                    continue;
                }

                FillChildDocumentsTree(childDocNode, resultList);
            }
        }

        StringBuilder strB = new StringBuilder();
        
        private List<NamedDocumentGroup> GroupReportData(List<DocParams> docList)
        {
            var resultList = new List<NamedDocumentGroup>();
            if (docList.Count == 0)
                return resultList;
            else if (docList.Count == 1)
            {
                var docGroup = new NamedDocumentGroup(docList[0].Naimenovanie);
                docGroup.Add(docList[0]);
                resultList.Add(docGroup);
                return resultList;
            }

            // Сортируем по наименованию
           docList.Sort(delegate(DocParams x, DocParams y) { return x.Naimenovanie.CompareTo(y.Naimenovanie); });

            var groupedDocuments = new Dictionary<string, List<DocParams>>();
            foreach (DocParams doc in docList)
            {
                string firstNameWord = doc.NameStringArray[0];

                if (!groupedDocuments.ContainsKey(firstNameWord))
                    groupedDocuments.Add(firstNameWord, new List<DocParams>());

                groupedDocuments[firstNameWord].Add(doc);
            }
            var otherGroup = new NamedDocumentGroup("Элементы разные");
            BuyListReport.m_form.setStage(Form1Namespace.Form1.Stages.ReportGenerating);
            BuyListReport.m_writeToLog("=== Формирование отчета ===");
            foreach (List<DocParams> docGroup in groupedDocuments.Values)
            {
                if (IsSingleDocument(docGroup))
                {
                    for (int i = 0; i < docGroup.Count; i++)
                        otherGroup.Add(docGroup[i]);
                }
                else
                    resultList.Add(new NamedDocumentGroup(GetGroupName(docGroup), docGroup));
            }

            if (otherGroup.Count > 0)
                resultList.Add(otherGroup);
            return resultList;

        }

        private bool IsSingleDocument(List<DocParams> list)
        {
            if (list.Count == 1)
                return true;
            DocParams firstDoc = list[0];
            for (int listIndex = 1; listIndex < list.Count; listIndex++)
            {
                if (firstDoc.CompareTo(list[listIndex]) != 0)
                    return false;
            }
            return true;
        }

        private string GetGroupName(List<DocParams> documents)
        {
            if (documents.Count == 0)
                return String.Empty;
            else if (documents.Count == 1)
                return documents[0].Naimenovanie;

            string resultName = String.Join(" ", documents[0].NameStringArray);
            for (int docIndex = 1; docIndex < documents.Count; docIndex++)
            {
                string commonName = GetCommonName(documents[docIndex - 1].NameStringArray, documents[docIndex].NameStringArray);
                if (commonName.Length < resultName.Length)
                    resultName = commonName;
            }
            return resultName;
        }

        private string GetCommonName(string[] name1, string[] name2)
        {
            const string DocCodeRegEx = @"\S+[-0-9\.]";

            if (name1.Length == 0 || name2.Length == 0)
                return String.Empty;

            string resultName = String.Empty;
            int maxResultLength = Math.Min(name1.Length, name2.Length);

            for (int commonIndex = 0; commonIndex < maxResultLength; commonIndex++)
            {
                string namePart1 = name1[commonIndex];
                string namePart2 = name2[commonIndex];

                if (String.Compare(namePart1, namePart2, true) != 0)
                    break;

                if (commonIndex == 0)
                    resultName = namePart1;
                else
                    resultName = String.Concat(resultName, " ", namePart1);
            }

            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(resultName, DocCodeRegEx);
            if (match.Success)
            {
                resultName = resultName.Replace(match.Value, String.Empty).Trim();
            }

            return resultName.TrimEnd(',', '.', ';', ':');
        }

        private string GetNonZeroString(double value)
        {
            if (value == 0)
                return String.Empty;
            return value.ToString();
        }
    }

    internal class KeyValueList<TKey, TValue> : List<KeyValuePair<TKey, TValue>>
    {
        public void Add(TKey key, TValue value)
        {
            Add(new KeyValuePair<TKey, TValue>(key, value));
        }
    }

    #endregion

    #region Config

    public class ReportParams
    {
        private string _naimenovanie = String.Empty;
        private string _obozna4enie = String.Empty;
        private string _razrabot4ik = String.Empty;
        private string _proveril = String.Empty;
        private string _proverilDate = String.Empty;
        private string _nKontr = String.Empty;
        private string _nKontrDate = String.Empty;
        private string _utverdil = String.Empty;
        private string _utverdilDate = String.Empty;

        /// <summary>
        /// Конструктор по умолчанию
        /// </summary>
        public ReportParams(ReferenceObject doc)
        {
            var refNomenclature = new NomenclatureReference();
            TFlex.DOCs.Model.Structure.ParameterInfo nameInfo =
                refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
            TFlex.DOCs.Model.Structure.ParameterInfo denotationInfo =
                refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);
            Наименование = (string) doc[nameInfo].Value;
            Обозначение = (string) doc[denotationInfo].Value;
        }

        public string Наименование
        {
            get { return _naimenovanie; }
            set { _naimenovanie = value; }
        }

        public string Обозначение
        {
            get { return _obozna4enie; }
            set { _obozna4enie = value; }
        }

        public string Разработчик
        {
            get { return _razrabot4ik; }
            set { _razrabot4ik = value; }
        }

        public string ДатаРазработки
        {
            get { return _razrabot4ik; }
            set { _razrabot4ik = value; }
        }

        public string Проверил
        {
            get { return _proveril; }
            set { _proveril = value; }
        }

        public string ДатаПроверки
        {
            get { return _proverilDate; }
            set { _proverilDate = value; }
        }

        public string НормКонтроль
        {
            get { return _nKontr; }
            set { _nKontr = value; }
        }

        public string ДатаНормКонтроля
        {
            get { return _nKontrDate; }
            set { _nKontrDate = value; }
        }

        public string Утвердил
        {
            get { return _utverdil; }
            set { _utverdil = value; }
        }

        public string ДатаУтверждения
        {
            get { return _utverdilDate; }
            set { _utverdilDate = value; }
        }
    }

    #endregion

    #region DOCs

    /// <summary>
    /// Работа с DOCs
    /// </summary>
    internal static class ApiDocs
    {
        public static SqlConnection GetConnection(bool open, IReportGenerationContext context)
        {
            SqlConnectionStringBuilder sqlConStringBuilder = new SqlConnectionStringBuilder();
            sqlConStringBuilder.DataSource = TFlex.DOCs.Model.ServerGateway.Connection.ServerName;
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
        public static List<TFDDocument> ChildDocList = new List<TFDDocument>();       
        public static Dictionary<int, TFDDocument> DocumentsById = new Dictionary<int, TFDDocument>();        
        public static Dictionary<int, List<TFDDocument>> DocumentsByParentId = new Dictionary<int, List<TFDDocument>>();
        
        public static bool halt;

       
        /// <summary>
        /// Формирует списки разделов ведомости на основе документов DOCs
        /// </summary>
        /// <param name="baseDocID">ID головного документа</param>
        public static ReportData GetDocsReportData(int baseDocID, IReportGenerationContext context)
        {
            NomenclatureReference refNomenclature = new NomenclatureReference();            
           
            // получаем корневой объект иерархии Номенклатура и изделия отчета
            TFlex.DOCs.Model.References.ReferenceObject baseReportDoc = refNomenclature.Find(baseDocID);  
            BuyListReport.m_writeToLog(String.Format("Идентификатор корневого документа: id {0}", baseDocID.ToString()));
            Defines.BaseDoc.ObjectId = baseDocID;
            return new ReportData(baseReportDoc,context);
        }
    }

    #endregion

    #region CAD

    internal class RowTable
    {
        public string Name;

        public string CodeRCP;
        public int CountPCodeRCP;
        public string Denotation;
        public int CountPDenotation;
        public string Supplier;
        public int CountPSupplier;
        public string WhereEnters;
        public int CountPWhereEnters;
        public string CountOnProduct;
        public int CountPCountOnProduct;
        public string CountInComplements;
        public int CountPCountInComplements;
        public string CountOnRegulation;
        public int CountPCountOnRegulation;
        public string CountTotal;
        public int CountPCountTotal;
        public string Remarks;
        public int CountPRemarks;

        public void SetCountPerenos(int countPerenos)
    {
        this.CountPCodeRCP = countPerenos;
        this.CountPDenotation = countPerenos;
        this.CountPSupplier = countPerenos;
        this.CountPWhereEnters = countPerenos;
        this.CountPCountOnProduct = countPerenos;
        this.CountPCountInComplements = countPerenos;
        this.CountPCountOnRegulation = countPerenos;
        this.CountPCountTotal = countPerenos;
        this.CountPRemarks = countPerenos;
    }
    }


    /// <summary>
    /// Работа с CAD APi
    /// </summary>
    internal  class ApiCad
    {
        /// <summary>
        /// ссылка на документ отчета
        /// </summary>
        private static TFlex.Model.Document _repDoc = null;
        
        /// <summary>
        /// Возвращает параграф-текст документа ReportDoc
        /// </summary>
        /// <param name="parTextName">имя параграф-текста</param>
        /// <returns>параграф текст</returns>
        public static ParagraphText GetParText(string parTextName)
        {
            
            ParagraphText ParText = ReportDoc.GetObjectByName(parTextName) as ParagraphText;

            if (ParText == null)
                throw new NullReferenceException("Не найден параграф текст с именем: " + parTextName);
            return ParText;
        }
        
        /// <summary>
        /// Документ отчета
        /// </summary>
        public static TFlex.Model.Document ReportDoc
        {
            get
            {
                if (_repDoc == null)
                    throw new NullReferenceException("Документ отчета не найден!");

                return _repDoc;
            }
        }
        
        /// <summary>
        /// Возвращает ID документа DOCs, для которого строится отчет
        /// </summary>
        /// <returns>ID документа</returns>
        public static int GetBaseDocID(IReportGenerationContext context)
        {
            if (context.ObjectsInfo.Count == 1) return context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }
        }

        /// <summary>
        /// Инициализация CAD APi
        /// </summary>
        public static void InitApi(IReportGenerationContext context)
        {
            var setup = new ApplicationSessionSetup
            {
                Enable3D = true,
                ReadOnly = false,
                PromptToSaveModifiedDocuments = false,
                EnableMacros = true,
                EnableDOCs = true,
                DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version13,
                ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
            };
            bool result = TFlex.Application.InitSession(setup);

            if (!result)
                throw new InvalidOperationException("Ошибка подключения к Cad Api");
                  
           _repDoc = TFlex.Application.OpenDocument(context.ReportFilePath);                 
       
        }
        /// <summary>
        /// Возвращает переменную документа по ее названию
        /// </summary>
        /// <param name="doc">документ</param>
        /// <param name="varName">имя переменной</param>
        /// <returns>найденная переменная документа</returns>
        public static Variable FindDocVariable(TFlex.Model.Document doc, string varName)
        {
            if (doc == null || String.IsNullOrEmpty(varName))
                return null;

            foreach (Variable currentVar in doc.Variables)
            {
                if (String.Compare(currentVar.Name, varName) == 0)
                    return currentVar;
            }

            return null;
        }
        /// <summary>
        /// Строит отчет на базе подготовленных данных DOCs
        /// </summary>
        /// <param name="reportPars">параметры отчета</param>
        /// <param name="reportData">данные отчета</param>
        public static void BuildReport(ReportParams reportPars, ReportData reportData, DocumentAttributes documentAttr)
        {         
            Dictionary<string, KeyValueList<ReportData.RowType, string[]>> rowGroups = reportData.GetReportRows(reportPars.Обозначение);       
            if (rowGroups.Count == 0)
                return;
            ReportDoc.BeginChanges("Формирование отчета");

            var reportText = ReportDoc.GetTextByName("$data_table");
            reportText.BeginEdit();
            var contentTable = reportText.GetFirstTable();
            var cadTable = new CadTable(contentTable.Value, reportText, 22, 28);
            var objCount = 0;
            var font = new Font("T-FLEX Type B", 2.5f, GraphicsUnit.Millimeter); 
            var textFormatter = new TextFormatter(font);

            foreach (var groupName in rowGroups.Keys)
            {
                KeyValueList<ReportData.RowType, string[]> groupRows = rowGroups[groupName];
                objCount += groupRows.Count();
            }
            BuyListReport.m_writeToLog("Количество элементов отчета: " + objCount);

            //==================================================-Старый избыточный код=============================================
            var listRowContent = new List<CadTableRow>();

            int countPage = 1;
            int countRow = 0;
            //Запись содержания ведомости без номеров страниц (только наименования групп) 
            foreach (string groupName in rowGroups.Keys)
            {
                string[] wrapCol = null;
                
                countRow++;
                //Если страница первая то вставляем до 22 строки
                if (countPage == 1)
                {
                    if (countRow == cadTable.countLinesOnFirstPage)
                    {
                        countPage++;
                        countRow = countRow - cadTable.countLinesOnFirstPage;
                    }
                }
                // иначе всталяем до 28
                else
                {
                    if (countRow == cadTable.countLinesOnOtherPages)
                    {
                        countPage++;
                        countRow = countRow - cadTable.countLinesOnOtherPages;
                    }
                }

                wrapCol = textFormatter.Wrap(groupName);
                //Название группы в одну строку
                if (wrapCol.Count() == 1)
                {
                    var rowContent = cadTable.CreateRow();
                    if (listRowContent.Count == 0)
                        rowContent.Insert((uint)(rowGroups.Count - 1), (Table)contentTable);
                    rowContent.Insert((uint) (rowGroups.Count - 1), (Table) contentTable);
                    rowContent.AddText(groupName, 0, CadTableCellJust.Left, TextStyle.Italic);
                    listRowContent.Add(rowContent);
                    countRow++;
                }
                //Название группы в несколько строк
                else
                {
                    for (int i = 0; i < wrapCol.Count(); i++)
                    {
                        var rowContent2 = cadTable.CreateRow();
                        countRow++;
                        rowContent2.Insert((uint)(rowGroups.Count - 1), (Table)contentTable);
                        rowContent2.AddText(wrapCol[i], 0, CadTableCellJust.Left, TextStyle.Italic);
                        if( i == 0)
                            listRowContent.Add(rowContent2);
                    }
                }
            }
            cadTable.CreateRow();
            // Если нажат флажок страница после содержания вставляем пустую страницу
            if (documentAttr.AddList)
            {
                var rowContent = cadTable.CreateRow();
                cadTable.NewPage(1);
                countPage += 2;
                rowContent.AddText("", CadTableCellJust.Center);
            }
            reportText.AutoUpdate = false;
            //вставляем данные в таблицу 
            CharFormat defaultCharFormat = reportText.CharacterFormat;
            CharFormat underlinedCharFormat = defaultCharFormat;
            underlinedCharFormat.isUnderline = true;

            //Количество ячеек в строке
            const uint columnsCount = 10;
            //Подчеркивание
            bool underlined = false;
            bool wrapName = false;
            var pageNumberOnAdding = new List<int>();
            var pageNumberAfterAdding = new List<int>();
            countRow = 0;

            foreach (string groupName in rowGroups.Keys)
            {
                //Все элементы раздела
                KeyValueList<ReportData.RowType, string[]> groupRows = rowGroups[groupName];
                int iR = 0;
                pageNumberOnAdding.Add(countPage);
                foreach (KeyValuePair<ReportData.RowType, string[]> row in groupRows)
                {
                    iR++;
                    string[] wrapCol = null;
                    var row1 = cadTable.CreateRow();
                    countRow++;
                    bool writeNameGroup = false;
                    wrapName = false;
                    //Добавляем страницу если строки на странице закончились
                    if (countRow >= cadTable.countLinesOnOtherPages)
                    {
                        countPage++;
                        countRow = countRow - cadTable.countLinesOnOtherPages;
                    }
                    //Запись строки по ячейкам
                    for (uint columnIndex = 0; columnIndex < columnsCount; columnIndex++)
                    {
                        underlined = false;
                        // Разделяем длинное наименование на несколько строк
                        if (row.Value[columnIndex] != null && columnIndex == 0)
                        {
                            wrapCol = textFormatter.Wrap(row.Value[columnIndex]); 
                        }

                        //Условие для наименования группы
                        if (row.Value[columnIndex] != null && columnIndex == 0 && row.Value[columnIndex].Trim() == groupName.Trim())
                        {
                            BuyListReport.m_writeToLog("Запись группы " + row.Value[0] + " на странице " + countPage);
                            underlined = true;
                            if (wrapCol.Count() > 1)
                            {
                                row1.AddText(wrapCol[0], CadTableCellJust.Center, underlined ? TextStyle.ItalicUnderline : TextStyle.Italic);
                                wrapName = true;
                                writeNameGroup = true;
                            }
                            else
                            {
                                row1.AddText(row.Value[columnIndex], CadTableCellJust.Center,
                                             underlined ? TextStyle.ItalicUnderline : TextStyle.Italic);
                                cadTable.CreateRow(2);
                                countRow += 2;
                            }
                        }
                        else
                        {
                            //Запись первой строки наименования , если она разбивается на несколько строк
                            if (wrapCol !=null && wrapCol.Count() > 1 && columnIndex ==0)
                            {
                                row1.AddText(wrapCol[0], CadTableCellJust.Left);
                                wrapName = true;
                            }
                            else
                            {
                                if ((row.Value[0] == null) && (row.Value[1] == null) && row.Value[2] == null && row.Value[3] == null && row.Value[4] == null
                                    && row.Value[5] == null)
                                {
                                    row1.AddText(row.Value[columnIndex], CadTableCellJust.Left);
                                }
                                else
                                {
                                    // Проверяем следующий элемент коллекции, если он итоговое число то подчеркиваем элемент из "Общего количествва"
                                    if (iR < groupRows.Count && columnIndex == 8 && groupRows[iR].Value[0] == null && groupRows[iR].Value[1] == null
                                        && groupRows[iR].Value[2] == null && groupRows[iR].Value[3] == null && groupRows[iR].Value[4] == null 
                                        && groupRows[iR].Value[5] == null && groupRows[iR].Value[6] == null && groupRows[iR].Value[7] == null 
                                        && groupRows[iR].Value[8] != null)
                                    {
                                        row1.AddText(row.Value[columnIndex].Trim(), CadTableCellJust.Left, TextStyle.ItalicUnderline);
                                    }
                                    else
                                    {
                                        row1.AddText(row.Value[columnIndex], CadTableCellJust.Left);
                                    }
                                }
                            }
                        }

                        //Добавляем строки для наименования, которые не вместились на первой строке
                        if (columnIndex == columnsCount - 1 && wrapName)
                        {
                            if (writeNameGroup)
                            {
                                for (int i = 1; i < wrapCol.Count(); i++)
                                {
                                    row1 = cadTable.CreateRow();
                                    countRow++;
                                    row1.AddText(wrapCol[i], CadTableCellJust.Center, TextStyle.ItalicUnderline );
                                }
                                cadTable.CreateRow(2);
                                countRow += 2;
                            }
                            else
                            {
                                for (int i = 1; i < wrapCol.Count(); i++)
                                {
                                    row1 = cadTable.CreateRow();
                                    countRow++;
                                    if (row.Value[0].Trim() == groupName.Trim())
                                    {
                                        row1.AddText(wrapCol[i], CadTableCellJust.Center,
                                                     underlined ? TextStyle.ItalicUnderline : TextStyle.Italic);
                                    }
                                    else
                                    {
                                        row1.AddText(wrapCol[i], CadTableCellJust.Left);
                                    }
                                }
                            }   
                        }
                    }
                }
                pageNumberAfterAdding.Add(countPage);
                // Добавляем страницу после раздела, если нажат соответствующий флажок
                if (documentAttr.NewListSection)
                {
                    cadTable.NewPage();
                    countPage++;
                    countRow = 0;
                }
                var row2 = cadTable.CreateRow();
                row2.AddText("", CadTableCellJust.Center);    
            }
           
            //заполняем номера страниц в содержании
            int k = 0;
            foreach (var rowContent in listRowContent)
            {
                string str = string.Empty;
                rowContent.AddText("", CadTableCellJust.Left, TextStyle.Italic);
                rowContent.AddText("", CadTableCellJust.Left, TextStyle.Italic);
                if (pageNumberOnAdding[k] == pageNumberAfterAdding[k])
                    str = (String.Format("Лист {0}", pageNumberOnAdding[k] + 1));
                else
                    str =  (String.Format("Листы {0}-{1}", pageNumberOnAdding[k] + 1, pageNumberAfterAdding[k] + 1)); 
                rowContent.AddText(str, CadTableCellJust.Left, TextStyle.Italic);
                k++;
            }
            cadTable.Apply();
            reportText.EndEdit();
            AddPageForms(ReportDoc,reportPars, documentAttr);
            ReportDoc.EndChanges();
            BuyListReport.m_writeToLog("=== Ведомость покупных изделий сформирована ===");
            BuyListReport.m_form.progressBarParam(20);
            Form1.ActiveForm.Close();

            ReportDoc.Save();
            ReportDoc.Close();   
            
        }
        // Вставка значений заданных атрибутов в форму
        public static void AddPageForms(TFlex.Model.Document doc, ReportParams reportPars, DocumentAttributes docAttr)//TFlex.Model.Document cadDocument, ReportParams reportPars, DocumentAttributes docAttr)
        {
            var pages = new List<Page>();
            foreach (Page page in doc.GetPages())
            {
                pages.Add(page);
            }
            int pageNumber = 1;
            foreach (Page page in pages)
            {
                ++pageNumber;
               
                // задание значений переменных фрагмента
                foreach (Fragment fragment in page.GetFragments())
                {
                    foreach (FragmentVariableValue var in fragment.GetVariables())
                    {
                        var.AttachedVariable = null;
                        switch (var.Name)
                        {
                            case "$oboznach":
                                    var.TextValue = reportPars.Обозначение;
                                    break;
                            case "$naimen1":
                                var.TextValue = reportPars.Наименование;
                                break;
                            case "$izm2":
                                if (docAttr.AddModify)
                                {

                                    foreach (int nPage in docAttr.ListPage)
                                        if (pageNumber == nPage)
                                        {
                                            var.TextValue = docAttr.NumberModify;

                                        }

                                    foreach (int nPage in docAttr.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.NumberModify;
                                }
                                break;
                            case "$il2":
                                if (docAttr.AddModify)
                                {
                                    foreach (int nPage in docAttr.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.ListZam;

                                    foreach (int nPage in docAttr.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.ListNov;
                                }
                                break;
                            case "$ido2":
                                if (docAttr.AddModify)
                                {
                                    foreach (int nPage in docAttr.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.NumberDocument.ToString();

                                    foreach (int nPage in docAttr.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.NumberDocument.ToString();
                                }
                                break;
                            case "$idata2":
                                if (docAttr.AddModify)
                                {
                                    foreach (int nPage in docAttr.ListPage)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.DateModify;

                                    foreach (int nPage in docAttr.ListPage2)
                                        if (pageNumber == nPage)
                                            var.TextValue = docAttr.DateModify;
                                }
                                break;
                        }
                    }
                }
            }
        } 

        public static void FormatCell(ParagraphText parText, ref Table parTextTable, uint cell, bool underlined, bool bold, bool italic, ParaFormat.Just justification)
        {
            FormatCell(parText, ref parTextTable, cell, underlined, bold, italic);
            FormatCell(parText, ref parTextTable, cell, justification);
        }

        public static void FormatCell(ParagraphText parText, ref Table parTextTable, uint cell, bool underlined, bool bold, bool italic)
        {
            CharFormat newCharacterFormat = parText.CharacterFormat;
            newCharacterFormat.isBold = bold;
            newCharacterFormat.isUnderline = underlined;
            newCharacterFormat.isItalic = italic;
            newCharacterFormat.isDefaultItalic = false;
            newCharacterFormat.Color = parText.CharacterFormat.Color;
            newCharacterFormat.FontName = parText.CharacterFormat.FontName;
            newCharacterFormat.FontSize = parText.CharacterFormat.FontSize;
            newCharacterFormat.isStrikeout = parText.CharacterFormat.isStrikeout;
            newCharacterFormat.Space = parText.CharacterFormat.Space;
            newCharacterFormat.VertOffset = parText.CharacterFormat.VertOffset;

            parTextTable.SetSelection(cell, cell);
            parText.CharacterFormat = newCharacterFormat;
        }

        public static void FormatCell(ParagraphText parText, ref Table parTextTable, uint cell, ParaFormat.Just justification)
        {
            ParaFormat newParagraphFormat = parText.ParagraphFormat;
            newParagraphFormat.HorJustification = justification;
            newParagraphFormat.FirstLineOffset = parText.ParagraphFormat.FirstLineOffset;
            newParagraphFormat.FitOneLine = parText.ParagraphFormat.FitOneLine;
            newParagraphFormat.LeftIndent = parText.ParagraphFormat.LeftIndent;
            newParagraphFormat.LineSpace = parText.ParagraphFormat.LineSpace;
            newParagraphFormat.LineSpaceMode = parText.ParagraphFormat.LineSpaceMode;
            newParagraphFormat.NumberFormat = parText.ParagraphFormat.NumberFormat;
            newParagraphFormat.Numbering = parText.ParagraphFormat.Numbering;
            newParagraphFormat.NumProps = parText.ParagraphFormat.NumProps;
            newParagraphFormat.ReduceExtension = parText.ParagraphFormat.ReduceExtension;
            newParagraphFormat.RightIndent = parText.ParagraphFormat.RightIndent;
            newParagraphFormat.SpaceAfterLast = parText.ParagraphFormat.SpaceAfterLast;
            newParagraphFormat.SpaceBeforeFirst = parText.ParagraphFormat.SpaceBeforeFirst;
            newParagraphFormat.TabSize = parText.ParagraphFormat.TabSize;

            parTextTable.SetSelection(cell, cell);
            parText.ParagraphFormat = newParagraphFormat;
        }
    }

    #endregion

    #region Defines

    namespace Defines
    {
        internal static class TFDClass
        {
            public static readonly int Assembly = new NomenclatureReference().Classes.AllClasses.Find("Сборочная единица").Id;
            public static readonly int Detal = new NomenclatureReference().Classes.AllClasses.Find("Деталь").Id;
            public static readonly int Document = new NomenclatureReference().Classes.AllClasses.Find("Документ").Id;
            public static readonly int Komplekt = new NomenclatureReference().Classes.AllClasses.Find("Комплект").Id;
            public static readonly int Complex = new NomenclatureReference().Classes.AllClasses.Find("Комплекс").Id;
            public static readonly int OtherItem = new NomenclatureReference().Classes.AllClasses.Find("Прочее изделие").Id;
            public static readonly int StandartItem = new NomenclatureReference().Classes.AllClasses.Find("Стандартное изделие").Id;
            public static readonly int Material = new NomenclatureReference().Classes.AllClasses.Find("Материал").Id;
            public static readonly int ComponentProgram = new NomenclatureReference().Classes.AllClasses.Find("Компонент (программы)").Id;
            public static readonly int ComplexProgram = new NomenclatureReference().Classes.AllClasses.Find("Комплекс (программы)").Id;
        }

        internal class BaseDoc
        {
            public static int ObjectId;
        }

        /// <summary>
        /// Дифайны переменных документа отчета
        /// </summary>
        internal static class ReportDocVars
        {
            /// <summary>
            /// Кол-во строк до заголовка типа покупного изделия по умолчанию
            /// </summary>
            public const int LinesBeforeHeaderDefault = 0;
            /// <summary>
            /// Кол-во строк после заголовка типа покупного изделия по умолчанию
            /// </summary>
            public const int LinesAfterHeaderDefault = 0;

            /// <summary>
            /// Имя параграф текста, в который заносятся данные отчета
            /// </summary>
            public const string DataTableParText = "$data_table";
            /// <summary>
            /// Имя переменной, хранящей кол-во строк до названия типа покупного изделия
            /// </summary>
            public const string LinesBeforeHeader = "$lines_before_header";
            /// <summary>
            /// Имя переменной, хранящей кол-во строк после названия типа покупного изделия
            /// </summary>
            public const string LinesAfterHeader = "$lines_after_header";
        }
    }

    #endregion

   
    namespace Form1Namespace
    {
        using System;

        public partial class Form1 : System.Windows.Forms.Form
        {
            private System.Windows.Forms.TextBox LogTextBox;

            private System.Windows.Forms.Button CanselButton;

            private System.Windows.Forms.Label LabelProgressBar;

            private System.Windows.Forms.ProgressBar progressBar;

            private System.Windows.Forms.Label Loglabel;

            private System.Windows.Forms.OpenFileDialog openFileDialog1;

            private void InitializeComponent()
            {
                this.LogTextBox = new System.Windows.Forms.TextBox();
                this.CanselButton = new System.Windows.Forms.Button();
                this.LabelProgressBar = new System.Windows.Forms.Label();
                this.progressBar = new System.Windows.Forms.ProgressBar();
                this.Loglabel = new System.Windows.Forms.Label();
                this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
                this.SuspendLayout();
                // 
                // LogTextBox
                // 
                this.LogTextBox.BackColor = System.Drawing.SystemColors.ActiveCaption;
                this.LogTextBox.Location = new System.Drawing.Point(27, 37);
                this.LogTextBox.Multiline = true;
                this.LogTextBox.Name = "LogTextBox";
                this.LogTextBox.Size = new System.Drawing.Size(430, 168);
                this.LogTextBox.TabIndex = 6;
                // 
                // CanselButton
                // 
                this.CanselButton.Location = new System.Drawing.Point(379, 288);
                this.CanselButton.Name = "CanselButton";
                this.CanselButton.Size = new System.Drawing.Size(78, 23);
                this.CanselButton.TabIndex = 5;
                this.CanselButton.Text = "Отмена";
                this.CanselButton.UseCompatibleTextRendering = true;
                this.CanselButton.UseVisualStyleBackColor = true;
                this.CanselButton.Click += new System.EventHandler(this.CanselButton_Click);
                // 
                // LabelProgressBar
                // 
                this.LabelProgressBar.Location = new System.Drawing.Point(27, 226);
                this.LabelProgressBar.Name = "LabelProgressBar";
                this.LabelProgressBar.Size = new System.Drawing.Size(169, 15);
                this.LabelProgressBar.TabIndex = 4;
                this.LabelProgressBar.Text = "Инициализация...";
                this.LabelProgressBar.UseCompatibleTextRendering = true;
                // 
                // progressBar
                // 
                this.progressBar.Location = new System.Drawing.Point(27, 244);
                this.progressBar.Name = "progressBar";
                this.progressBar.Size = new System.Drawing.Size(430, 23);
                this.progressBar.TabIndex = 3;
                // 
                // Loglabel
                // 
                this.Loglabel.Location = new System.Drawing.Point(27, 20);
                this.Loglabel.Name = "Loglabel";
                this.Loglabel.Size = new System.Drawing.Size(150, 14);
                this.Loglabel.TabIndex = 2;
                this.Loglabel.Text = "Журнал операций";
                this.Loglabel.UseCompatibleTextRendering = true;
                // 
                // openFileDialog1
                // 
                this.openFileDialog1.FileName = "openFileDialog1";
                // 
                // Form1
                // 
                this.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
                this.ClientSize = new System.Drawing.Size(481, 323);
                this.Controls.Add(this.LogTextBox);
                this.Controls.Add(this.CanselButton);
                this.Controls.Add(this.LabelProgressBar);
                this.Controls.Add(this.progressBar);
                this.Controls.Add(this.Loglabel);
                this.DoubleBuffered = false;
                this.Text = "Ведомость покупных изделий ГОСТ 2.106-96";
                this.ResumeLayout(false);
                this.PerformLayout();
            }
        }
    }
    
    namespace Form1Namespace
    {
        public partial class Form1
	    {
            public void writeToLog(string line)
            {
                DateTime now = DateTime.Now;
                string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
                LogTextBox.AppendText(s);
                LogTextBox.Invalidate();
                this.Update();
                System.Windows.Forms.Application.DoEvents();
            }

	        public enum Stages
            {
                Initialization,
                DataAcquisition,
                DataProcessing,
                ReportGenerating,
                Done
            } 

		    public Form1()
		    {
                System.Windows.Forms.Application.EnableVisualStyles();	
			    InitializeComponent();
		    }
        
            public  void progressBarParam (int step)
            {
		        progressBar.Step=step;
                progressBar.PerformStep();
            }
		
            public void setStage(Stages stage)
            {
		        progressBarParam(3);
                switch (stage)
                {
                    case Stages.Initialization:				
                        LabelProgressBar.Text= "Инициализация...";                 
				        break;
                    case Stages.DataAcquisition:				
                        LabelProgressBar.Text= "Получение данных";
                        break;
                    case Stages.DataProcessing:
                        LabelProgressBar.Text= "Обработка данных";                  
                        break;
                    case Stages.ReportGenerating:				
                        LabelProgressBar.Text= "Создание отчёта";					
                        break;
                    case Stages.Done:
                        break;
                    default:
                        Debug.Assert(false);
                        break;
                }
                this.Update();
            }

		    private void CanselButton_Click(object sender, System.EventArgs e)
		    {            
                Form1.ActiveForm.Dispose();			
		    }
	    }
    }
}
