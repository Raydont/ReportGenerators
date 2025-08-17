using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Defines;
using TFlex.DOCs.Model.Configuration.Variables;

namespace GroupSpecVarB
{
    /// <summary>
    /// Обработчик строки обозначения
    /// </summary>
    internal static class StringProcessor
    {
        private static NumberFormatInfo _numberFormat = null;

        /// <summary>
        /// Разделяет код обозначения на сам код и цифру, идущую за кодом
        /// </summary>
        /// <param name="codeWithNum">исходный код с цифрой</param>
        /// <param name="codeWithoutNum">код без цифры</param>
        /// <returns>цифра после кода</returns>
        private static void TrimCodeAndNum(string codeWithNum, out string codeWithoutNum1, out int codeNum1, out string codeWithoutNum2, out int codeNum2)
        {
            codeWithoutNum1 = String.Empty;
            codeWithoutNum2 = String.Empty;
            codeNum1 = 0;
            codeNum2 = 0;

            if (String.IsNullOrEmpty(codeWithNum))
                return;

            string CodeWithNumPattern = @"(?<code1>[a-z|а-я]+)[\s|-]?(?<num1>[0-9]*)((?<code2>[a-z|а-я]+)[\s|-]?(?<num2>[0-9]*))?";
            Match CodeNumFinder = Regex.Match(codeWithNum, CodeWithNumPattern, RegexOptions.IgnoreCase);

            if (!CodeNumFinder.Success)
                return;

            codeWithoutNum1 = CodeNumFinder.Groups["code1"].Value;
            string codeNumStr1 = CodeNumFinder.Groups["num1"].Value;
            codeWithoutNum2 = CodeNumFinder.Groups["code2"].Value;
            string codeNumStr2 = CodeNumFinder.Groups["num2"].Value;

            if (!String.IsNullOrEmpty(codeNumStr1))
                codeNum1 = Int32.Parse(codeNumStr1);

            if (!String.IsNullOrEmpty(codeNumStr2))
                codeNum2 = Int32.Parse(codeNumStr2);
        }
        /// <summary>
        /// Удаляет префикс в строке обозначения
        /// </summary>
        /// <param name="obozna4enie">обозначение</param>
        /// <param name="prefix">удаляемый префикс</param>
        /// <returns>обозначение без префикса</returns>
        private static string DelPrefix(string obozna4enie, out string prefix)
        {
            prefix = null;

            if (String.IsNullOrEmpty(obozna4enie))
                return null;

            string PrefixPattern = String.Concat(@"\A", Defines.OboznTypes.PrefixPattern, @"[\s|-]*");
            Match PrefixFinder = Regex.Match(obozna4enie, PrefixPattern, RegexOptions.IgnoreCase);

            if (!PrefixFinder.Success)
                return obozna4enie;

            prefix = PrefixFinder.Value.TrimEnd('-', ' ');
            return obozna4enie.Substring(PrefixFinder.Value.Length);
        }
        /// <summary>
        /// Удаляет постфикс УД / ЛУ из обозначения
        /// </summary>
        /// <param name="obozna4enie">обозначение</param>
        /// <param name="postfix">удаляемый постфикс</param>
        /// <returns>обозначние без постфикса</returns>
        private static string DelPostfix(string obozna4enie, out string postfix)
        {
            postfix = null;

            if (String.IsNullOrEmpty(obozna4enie))
                return null;

            postfix = obozna4enie.Substring(obozna4enie.Length - 2);
            string OboznWithoutPostfix = obozna4enie.Substring(0, obozna4enie.Length - 2);
            OboznWithoutPostfix = OboznWithoutPostfix.TrimEnd('-');
            OboznWithoutPostfix = OboznWithoutPostfix.TrimEnd(' ');

            return OboznWithoutPostfix;
        }
        /// <summary>
        /// Возвращает обозначение без номера исполнения
        /// </summary>
        /// <param name="obozna4enie">обозначение</param>
        /// <param name="varNumber">номер исполения</param>
        /// <returns>обозначение без номера исполнения</returns>
        private static string DelVariantNumber(string obozna4enie, out int varNumber)
        {
            varNumber = 0;

            if (String.IsNullOrEmpty(obozna4enie))
                return obozna4enie;

            string Pattern = String.Concat(@"\A([\S|\s]*)", Defines.OboznTypes.VarEndPart);
            Match TemplateFinder = Regex.Match(obozna4enie, Pattern, RegexOptions.IgnoreCase);

            if (!TemplateFinder.Success || TemplateFinder.Groups.Count < 2)
                return obozna4enie;

            string VarNumber = TemplateFinder.Groups[2].Value;
            varNumber = Convert.ToInt32(VarNumber.TrimStart('-'));

            return TemplateFinder.Groups[1].Value;
        }
        /// <summary>
        /// Определяет, соответствует ли строка шаблону
        /// </summary>
        /// <param name="val">строковое значение</param>
        /// <param name="pattern">шаблон</param>
        /// <returns>результат сопоставления шаблона и строки</returns>
        private static bool ValidateStringWithPattern(string val, string pattern)
        {
            if (String.IsNullOrEmpty(val) || String.IsNullOrEmpty(pattern))
                return false;

            Match Validator = Regex.Match(val, pattern, RegexOptions.IgnoreCase);
            if (Validator.Success)
                return true;

            return false;
        }

        /// <summary>
        /// Формат преобразования double в строку для отчета
        /// </summary>
        public static NumberFormatInfo NumberFormat
        {
            get
            {
                if (_numberFormat == null)
                {
                    _numberFormat = new NumberFormatInfo();
                    _numberFormat.CurrencyDecimalSeparator = ".";
                }

                return _numberFormat;

            }
        }
        /// <summary>
        /// Возвращает тип формата обозначения
        /// </summary>
        /// <param name="obozna4enie">обозначение</param>
        /// <returns>тип обозначения</returns>
        public static Defines.OboznTypes.FormatType GetOboznType(string obozna4enie)
        {
            if (String.IsNullOrEmpty(obozna4enie))
                return Defines.OboznTypes.FormatType.Unknown;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.Full))
                return Defines.OboznTypes.FormatType.Full;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.Full_Var))
                return Defines.OboznTypes.FormatType.Full_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullPostfix_LU))
                return Defines.OboznTypes.FormatType.FullPostfix_LU;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullPostfix_LU_Var))
                return Defines.OboznTypes.FormatType.FullPostfix_LU_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullPostfix_UD))
                return Defines.OboznTypes.FormatType.FullPostfix_UD;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullPostfix_UD_Var))
                return Defines.OboznTypes.FormatType.FullPostfix_UD_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCodePostfix_LU))
                return Defines.OboznTypes.FormatType.FullCodePostfix_LU;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCodePostfix_LU_Var))
                return Defines.OboznTypes.FormatType.FullCodePostfix_LU_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCodePostfix_UD))
                return Defines.OboznTypes.FormatType.FullCodePostfix_UD;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCodePostfix_UD_Var))
                return Defines.OboznTypes.FormatType.FullCodePostfix_UD_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCode))
                return Defines.OboznTypes.FormatType.FullCode;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.FullCode_Var))
                return Defines.OboznTypes.FormatType.FullCode_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortPostfix_LU))
                return Defines.OboznTypes.FormatType.ShortPostfix_LU;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortPostfix_LU_Var))
                return Defines.OboznTypes.FormatType.ShortPostfix_LU_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortPostfix_UD))
                return Defines.OboznTypes.FormatType.ShortPostfix_UD;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortPostfix_UD_Var))
                return Defines.OboznTypes.FormatType.ShortPostfix_UD_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCodePostfix_LU))
                return Defines.OboznTypes.FormatType.ShortCodePostfix_LU;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCodePostfix_LU_Var))
                return Defines.OboznTypes.FormatType.ShortCodePostfix_LU_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCodePostfix_UD))
                return Defines.OboznTypes.FormatType.ShortCodePostfix_UD;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCodePostfix_UD_Var))
                return Defines.OboznTypes.FormatType.ShortCodePostfix_UD_Var;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCode))
                return Defines.OboznTypes.FormatType.ShortCode;

            if (ValidateStringWithPattern(obozna4enie, Defines.OboznTypes.ShortCode_Var))
                return Defines.OboznTypes.FormatType.ShortCode_Var;

            return Defines.OboznTypes.FormatType.Unknown;
        }
        /// <summary>
        ///  Определяет тип обозначения и разбирает обозначение на: код обозначения, число после кода обозначения, постфикc, префикс и номер исполнения.
        /// </summary>
        /// <param name="obozna4enie">обозначение</param>
        /// <param name="type">тип обозначения</param>
        /// <param name="codeWithoutNum">код обозначения</param>
        /// <param name="codeNum">число после кода обозначения</param>
        /// <param name="postfix">постфикс ЛУ или УД</param>
        /// <param name="prefix">префикс (шифр формата слово.цифры.цифры.цифры...)</param>
        /// <param name="variantNumber">номер исполения</param> 
        public static void DivideObozn(string obozna4enie, out Defines.OboznTypes.FormatType type, out string codeWithoutNum1, out int codeNum1, out string codeWithoutNum2, out int codeNum2, out string postfix, out string prefix, out int variantNumber)
        {
            codeWithoutNum1 = null;
            codeWithoutNum2 = null;
            codeNum1 = 0;
            codeNum2 = 0;
            postfix = null;
            prefix = null;
            variantNumber = 0;
            type = GetOboznType(obozna4enie);

            if (String.IsNullOrEmpty(obozna4enie) || type == Defines.OboznTypes.FormatType.Unknown)
                return;

            string TempObozn = obozna4enie;

            if (Defines.OboznTypes.IsVariant(type))
                TempObozn = DelVariantNumber(TempObozn, out variantNumber);

            if (Defines.OboznTypes.HavePostfix(type))
                TempObozn = DelPostfix(TempObozn, out postfix);

            if (!Defines.OboznTypes.IsShort(type))
                TempObozn = DelPrefix(TempObozn, out prefix);

            TrimCodeAndNum(TempObozn, out codeWithoutNum1, out codeNum1, out codeWithoutNum2, out codeNum2);
        }
        /// <summary>
        /// Возвращает ID файла документа из пути к GRB файлу отчета
        /// </summary>
        /// <param name="grbReportFileDir">путь к директории, в которой расположен grb файл отчета</param>
        /// <returns>ID файла документа</returns>
        public static int GetReportDocFileID(string grbReportFileDir)
        {
            string GrbDirIdPattern = Defines.ReportDocVars.GrbDirIdPattern;
            Match Match = Regex.Match(grbReportFileDir.TrimEnd('\\'), GrbDirIdPattern, RegexOptions.IgnoreCase);

            if (!Match.Success)
                return 0;

            int ResultID;
            Int32.TryParse(Match.Groups[1].Value, out ResultID);

            return ResultID;
        }
        /// <summary>
        /// Извлекает номер исполнения из обозначения
        /// </summary>
        /// <param name="obozna4enie">обозначение документа</param>
        /// <returns>номер исполнения</returns>
        public static int GetVariantNumber(string obozna4enie)
        {
            if (String.IsNullOrEmpty(obozna4enie))
                return 0;

            Match RegexMatch = Regex.Match(obozna4enie, Defines.OboznTypes.VarEndPart, RegexOptions.IgnoreCase);

            if (!RegexMatch.Success || RegexMatch.Groups.Count < 2)
                return 0;

            string VarPart = RegexMatch.Groups[1].Value.TrimStart('-');

            return Int32.Parse(VarPart);
        }
    }

    /// <summary>
    /// Конфигурация отчета
    /// </summary>
    internal static class ReportConfig
    {
        /// <summary>
        /// Параметры отчета
        /// </summary>
        internal class ReportParams
        {
            private string _naimenovanie;
            private string _obozna4enie;
            private string _razrabot4ik;

            public ReportParams()
            {
                _naimenovanie = String.Empty;
                _obozna4enie = String.Empty;
                _razrabot4ik = String.Empty;
            }

            /// <summary>
            /// Наименование отчета
            /// </summary>
            public string Naimenovanie
            {
                get
                {
                    return _naimenovanie;
                }
                set
                {
                    if (value == null)
                        _naimenovanie = String.Empty;
                    else
                        _naimenovanie = value;
                }
            }
            /// <summary>
            /// Обозначение отчета
            /// </summary>
            public string Obozna4enie
            {
                get
                {
                    return _obozna4enie;
                }
                set
                {
                    if (value == null)
                        _obozna4enie = String.Empty;
                    else
                        _obozna4enie = value;
                }
            }
            /// <summary>
            /// Разработчик отчета
            /// </summary>
            public string Razrabot4ik
            {
                get
                {
                    return _razrabot4ik;
                }
                set
                {
                    if (value == null)
                        _razrabot4ik = String.Empty;
                    else
                        _razrabot4ik = value;
                }
            }
        }

        private static int _linesBeforeHeader = Defines.ReportDocVars.LinesBeforeHeaderDefault;
        private static int _linesAfterHeader = Defines.ReportDocVars.LinesAfterHeaderDefault;
        private static int _linesBeforeSection = Defines.ReportDocVars.LinesBeforeSectionDefault;
        private static int _linesAfterSection = Defines.ReportDocVars.LinesAfterSectionDefault;
        private static Dictionary<string, int> _codePriorMap = null;
        private static List<Variable> _configVars = null;

        /// <summary>
        /// Преобразует строку в int, при неудаче оставляет предыдущее значение переменной numResult
        /// </summary>
        /// <param name="var">переменная документа CAD</param>
        /// <param name="numResult">переменная в которую в случае успешной операции будет занесено значение переменной</param>
        private static void TryFillNumValWithVar(Variable var, ref int numResult)
        {
            if (var == null)
                return;

            int TempResult = -1;

            try
            {
                TempResult = Convert.ToInt32(var.Value);
            }
            catch (Exception)
            {
                return;
            }

            if (TempResult >= 0)
                numResult = TempResult;
        }
        /// <summary>
        /// Считывает список переменной и на его основе заполняет словарь кодов обозначений
        /// </summary>
        /// <param name="codePriorVar">перемнная, хранящая список приоритетов обозначений</param>
        /*private static void FillCodePriorMap(Variable codePriorVar)
        {
            if (codePriorVar == null)
                throw new NullReferenceException("Не найдена переменная $code_priority!");

            _codePriorMap = new Dictionary<string, int>();

            //читаем строки списка переменной codePriorVar
      //      for (int ListItemID = 0; ListItemID < codePriorVar.; ListItemID++)
            {
                //читаем код обозначения текущей строки списка кодов обозначений
                string CodeVal = GetCodeFromVarListItem(codePriorVar.GetValueListString(ListItemID));

                if (CodeVal != null)
                    if (!_codePriorMap.ContainsKey(CodeVal))
                        _codePriorMap.Add(CodeVal, ListItemID);
            }
        }*/
        /// <summary>
        /// Возвращает код обозначения записанный в строке в угловых скобках
        /// </summary>
        /// <param name="varListItem">строка с кодом обозначения</param>
        /// <returns>найденный код обозначения</returns>
        private static string GetCodeFromVarListItem(string varListItem)
        {
            int StartCodeIndex = varListItem.IndexOf('<');
            int EndCodeIndex = varListItem.IndexOf('>');

            if (StartCodeIndex == -1 || EndCodeIndex == -1 || EndCodeIndex < StartCodeIndex)
                return null;

            return varListItem.Substring(StartCodeIndex + 1, EndCodeIndex - StartCodeIndex - 1);
        }

        /// <summary>
        /// Число пустых строк, оставляемых до названия раздела спецификации
        /// </summary>
        public static int LinesBeforeHeader
        {
            get
            {
                return _linesBeforeHeader;
            }
        }
        /// <summary>
        /// Число пустых строк, оставляемых после названия раздела спецификации
        /// </summary>
        public static int LinesAfterHeader
        {
            get
            {
                return _linesAfterHeader;
            }
        }
        /// <summary>
        /// Число строк до "Обозн. исполн" для > 10 исполнений
        /// </summary>
        public static int LinesBeforeSection
        {
            get
            {
                return _linesBeforeSection;
            }
        }
        /// <summary>
        /// Число строк после "Обозн. исполн" для > 10 исполнений
        /// </summary>
        public static int LinesAfterSection
        {
            get
            {
                return _linesAfterSection;
            }
        }

        /// <summary>
        /// Возвращает приоритет кода обозначения
        /// </summary>
        /// <param name="code">код обозначения</param>
        /// <returns>приоритет кода обозначения</returns>
        public static int GetCodePriority(string code)
        {
            if (String.IsNullOrEmpty(code) || !_codePriorMap.ContainsKey(code))
                return 9999;

            return _codePriorMap[code];
        }
        /// <summary>
        /// Инициализация конфигурации отчета
        /// </summary>
        public static void InitConfig()
        {
            TFlex.Model.Document ReportDocument = ApiCad.ReportDoc;
            _codePriorMap = new Dictionary<string, int>();
            _configVars = new List<Variable>();
            Variable DocConfigVar = null;

           
          //  FillCodePriorMap(DocConfigVar);
        }
        /// <summary>
        /// Список переменных документа отчета, хранящих конфигурационные данные
        /// </summary>
        public static List<Variable> ConfigVars
        {
            get
            {
                return _configVars;
            }
        }
    }
    /// <summary>
    /// Данные строки отчета
    /// </summary>
    class ReportRow :  TFDDocument , IComparable<ReportRow>
    {
        /// <summary>
        /// Режим обработки обозначения объекта
        /// </summary>
        public enum Obozna4StringProcessType
        {
            /// <summary>
            /// Не обрабатывать обозначение вообще
            /// </summary>
            DontProcess,
            /// <summary>
            /// Обрабатывать только номер исполнения
            /// </summary>
            OnlyVariantIndex,
            /// <summary>
            /// Полная обработка
            /// </summary>
            FullProcessing
        }

        private string _code1 = String.Empty;
        private string _code2 = String.Empty;
        private int _codeNum1 = 0;
        private int _codeNum2 = 0;
        private string _prefix = null;
        private string _postfix = null;
        private string _sortIndex = null;
        private Defines.OboznTypes.FormatType _oboznType = Defines.OboznTypes.FormatType.Unknown;
        private int _variantIndex = 0;
        private bool _rowSortDataChanged = true;
        private int _parentDocVariantNumber = -1;
        private List<ReportRow> _sameRows = new List<ReportRow>();

        /// <summary>
        /// Устанавливает пустое обозначение и сбрасывает все связанные с ним параметры
        /// </summary>
        private void SetEmptyObozna4enie()
        {
            base.Denotation = String.Empty;
            _oboznType = Defines.OboznTypes.FormatType.Unknown;
            _code1 = String.Empty;
            _code2 = String.Empty;
            _codeNum1 = 0;
            _codeNum2 = 0;
            _variantIndex = 0;
            _postfix = String.Empty;
            _prefix = String.Empty;
        }
     
        /// <summary>
        /// Возвращает строковое значение, которое будет использовано для занесения в свойство
        /// </summary>
        /// <param name="val">исходное значение</param>
        /// <returns>значение для занесения в свойство</returns>
        private string GetStringValForSet(string val)
        {
            if (val == null)
                return null;
            else
                return val.Trim();
        }
        /// <summary>
        /// Формирует код приоритета сортировки
        /// </summary>
        /// <returns></returns>
       /* private string GetRowPriorSortIndex()
        {
            //приоритет головного раздела спецификации
            StringBuilder ResultPrior = new StringBuilder(Specification.BossTypeID.ToString(Defines.SortDefines.BaseSpecTypeDigits));
            ResultPrior.Append('!');

            //приоритет раздела спецификации
      //      ResultPrior.Append(Specification.SortPriority.ToString(Defines.SortDefines.CurrentSpecTypeDigits));
            ResultPrior.Append('!');

            //если объект принадлежит к разделу "Докуметация"
            if (Specification.BelongToDocumentation)
            {
                //записываем приоритет кода обозначения #1
                if (String.IsNullOrEmpty(Code1))
                    ResultPrior.Append(Defines.SortDefines.CodePriorDigits);
                else
                    ResultPrior.Append(ReportConfig.GetCodePriority(Code1).ToString(Defines.SortDefines.CodePriorDigits));

                ResultPrior.Append('!');

                //добавляем приоритет цифры после кода обозначения #1
                ResultPrior.Append(CodeNum1.ToString(Defines.SortDefines.CodeNumPriorDigits));

                ResultPrior.Append('!');

                //записываем приоритет кода обозначения #2
                if (String.IsNullOrEmpty(Code2))
                    ResultPrior.Append(Defines.SortDefines.CodePriorDigits);
                else
                    ResultPrior.Append(ReportConfig.GetCodePriority(Code2).ToString(Defines.SortDefines.CodePriorDigits));

                ResultPrior.Append('!');

                //добавляем приоритет цифры после кода обозначения #2
                ResultPrior.Append(CodeNum2.ToString(Defines.SortDefines.CodeNumPriorDigits));

                ResultPrior.Append('!');

                //формируем строку для сортировки префикса обозначения
                String PrefixString = new string('!', SortDefines.MaxPrefixLenght);

                //добавляем префикс обозначения
                if (!String.IsNullOrEmpty(PrefixObozna4enija))
                {
                    //записываем в начало PrefixString символы из PrefixObozna4enija
                    PrefixString = PrefixString.Insert(0, PrefixObozna4enija);
                    PrefixString = PrefixString.Substring(0, SortDefines.MaxPrefixLenght);
                    ResultPrior.Append(PrefixString);
                }
                else
                    ResultPrior.Append(PrefixString); //объекты без префикса имеют больший приоритет 

                ResultPrior.Append('!');

                //добавляем номер исполнения
                ResultPrior.Append(VariantIndex.ToString(Defines.SortDefines.VariantIndexPriorDigits));

                ResultPrior.Append('!');

                //добавляем постфикс обозначения (ЛУ / УД)
                if (!String.IsNullOrEmpty(PostfixObozna4enija))
                    ResultPrior.Append(PostfixObozna4enija);
                else
                    ResultPrior.Append("!!"); //объекты без постфикса имеют больший приоритет
            }
            //else
                //приоритет позиции на чертеже
               // ResultPrior.Append(PositionNumber.ToString(Defines.SortDefines.DocPositionDigits));

            return ResultPrior.ToString();
        }*/

        /// <summary>
        /// Инициализация
        /// </summary>
        /// <param name="doc">документ DOCs</param>
        /// <param name="loadParams">загружать ли все параметры документа DOCs</param>
        /// <param name="fillOnlyBaseParams">читать только основные парамтеры объекта</param>
        /// <param name="parentVariantNumber">порядковый номер родительского исполнения</param>
        /// <param name="stringProcessType">механизм обработки обозначения документа</param>
        public ReportRow(ref TFDDocument doc, bool loadParams, bool fillOnlyBaseParams, int parentVariantNumber, Obozna4StringProcessType stringProcessType)
           // : base(ref doc, loadParams, fillOnlyBaseParams)
        {
            if (stringProcessType == Obozna4StringProcessType.FullProcessing)
                this.Obozna4enie = base.Denotation; //заполняем Обозначение и связанные с ним поля 
            else if (stringProcessType == Obozna4StringProcessType.OnlyVariantIndex)
                _variantIndex = StringProcessor.GetVariantNumber(base.Denotation);

            _parentDocVariantNumber = parentVariantNumber;
        }
        /// <summary>
        /// Приоритет сортировки
        /// </summary>
        public string SortIndex
        {
            get
            {
                if (_rowSortDataChanged)
                {
                  //  _sortIndex = GetRowPriorSortIndex();
                    _rowSortDataChanged = false;
                }

                return _sortIndex;
            }
        }
        /// <summary>
        /// Код обозначения
        /// </summary>
        public string Code1
        {
            get { return _code1; }
        }
        /// <summary>
        /// Цифра кода обозначения
        /// </summary>
        public int CodeNum1
        {
            get { return _codeNum1; }
        }
        public string Code2
        {
            get { return _code2; }
        }
        public int CodeNum2
        {
            get { return _codeNum2; }
        }
        /// <summary>
        /// Количество для отчета
        /// </summary>
        public double Koli4estvo
        {
            get
            {
                //***исправить
                return  0;
            }
            set
            {
                Koli4estvo = value;
            }
        }
        /// <summary>
        /// Префикс обознчения (напр. "ПРОФ.9543.9543.3454")
        /// </summary>
        public string PrefixObozna4enija
        {
            get
            {
                return _prefix;
            }
        }
        /// <summary>
        /// Постфикс обозначения (УД или ЛУ)
        /// </summary>
        public string PostfixObozna4enija
        {
            get
            {
                return _postfix;
            }
        }
        /// <summary>
        /// Номер исполнения в обозначении
        /// </summary>
        public int VariantIndex
        {
            get
            {
                return _variantIndex;
            }
        }
        /// <summary>
        /// Порядковый номер родительского исполнения
        /// </summary>
        public int ParentDocVariantNumber
        {
            get
            {
                if (_parentDocVariantNumber < 0)
                    throw new InvalidDataException("Номер родительского исполнения не задан!");

                return _parentDocVariantNumber;
            }
        }
        /// <summary>
        /// Набор подобных объектов в других групповых исполнениях
        /// </summary>
        internal List<ReportRow> SameRows
        {
            get
            {
                return _sameRows;
            }
        }
        /// <summary>
        /// Поле "Формат" спецификации
        /// </summary>
        public string c1_Format
        {
            get
            {
                

                return Format;
            }
        }
        /// <summary>
        /// Поле "Зона" спецификации
        /// </summary>
        public string c2_Zona
        {
            get
            {
                return base.Zone;
            }
        }
        /// <summary>
        /// Поле "Поз." спецификации
        /// </summary>
        public string c3_Position
        {
            get
            {
             //   if (PositionNumber <= 0 || Specification.BelongToDocumentation)
               //     return String.Empty;
             //   else
                return string.Empty; //PositionNumber.ToString();
            }
        }
        /// <summary>
        /// Поле "Обозначение" спецификации
        /// </summary>
        public string c4_Obozna4enie
        {
            get
            {
                return Obozna4enie;
            }
        }
        /// <summary>
        /// Поле "Наименование" спецификации
        /// </summary>
        public string c5_Naimenovanie
        {
            get
            {
                string ResultVal = Naimenovanie;

                if (Defines.OboznTypes.HavePostfixLU(_oboznType))
                {
                    ResultVal = ResultVal.Replace("Лист утверждения", "\nЛист утверждения\n");
                    return ResultVal.Trim('\n');
                }
                else if (Defines.OboznTypes.HavePostfixUD(_oboznType))
                {
                    ResultVal = ResultVal.Replace("Удостоверяющий лист", "\nУдостоверяющий лист\n");
                    return ResultVal.Trim('\n');
                }
                else
                    return ResultVal;
            }
        }
        /// <summary>
        /// Поле "Количество" спецификации
        /// </summary>
        public string c6_Koli4estvo
        {
            get
            {
               // if (Specification.BelongToDocumentation)
                   // return "X";

                if (Koli4estvo == 0)
                    return String.Empty;

                return Koli4estvo.ToString(StringProcessor.NumberFormat);
            }
        }
        /// <summary>
        /// Поле "Примечание" спецификации
        /// </summary>
        public string c7_Prime4anie
        {
            get
            {
                
                    return Format;
              
            }
        }

        /// <summary>
        /// Производит сравнение с другим объектом
        /// </summary>
        /// <param name="row">объект, с которым производится сравнение</param>
        /// <returns>результат сравнения</returns>
        public bool LookLike(ReportRow row)
        {
          //  if (row == null || row.Specification == null)
          //      return false;

           // if (this.Specification.ID != row.Specification.ID)
           //     return false;

            if (String.Compare(this.c5_Naimenovanie, row.c5_Naimenovanie, true) != 0 ||
                String.Compare(this.c4_Obozna4enie, row.c4_Obozna4enie, true) != 0 ||
                String.Compare(this.c1_Format, row.c1_Format, true) != 0 ||
                String.Compare(this.c2_Zona, row.c2_Zona, true) != 0 ||
                String.Compare(this.c7_Prime4anie, row.c7_Prime4anie, true) != 0) //||
               // this.PositionNumber != row.PositionNumber)
                return false;

            return true;
        }
        /// <summary>
        /// Сравнение с объектом ReportRowData
        /// </summary>
        /// <param name="other">объект для сравнения</param>
        /// <returns>результат сравнения</returns>
        public int CompareTo(ReportRow other)
        {
            if (other == null)
                throw new ArgumentException("Ошибка сравнения строки спецификации!");

            int CompareResult = String.Compare(this.SortIndex, other.SortIndex, true);

            if (CompareResult == 0)
            {
                int Obozna4CompareResult = String.Compare(this.c4_Obozna4enie, other.c4_Obozna4enie, true);

                if (Obozna4CompareResult != 0)
                    return Obozna4CompareResult;

                return String.Compare(this.c5_Naimenovanie, other.c5_Naimenovanie, true);
            }

            return CompareResult;
        }
        /// <summary>
        /// Объединяет с данными другой строки
        /// </summary>
        /// <param name="other">данные другой строки</param>
        public void MergeWith(ReportRow other)
        {
            this.Koli4estvo += other.Koli4estvo;
        }

        /// <summary>
        /// Обозначение
        /// </summary>
        public  string Obozna4enie
        {
            get
            {
                return base.Denotation;
            }
            set
            {
                _rowSortDataChanged = true;

                if (String.IsNullOrEmpty(value))
                    SetEmptyObozna4enie();
                else
                {
                    base.Denotation = value;
                    //определяем тип обозначения, код обозначения, число после кода, префикс, постфикс и номер исполнения
                    StringProcessor.DivideObozn(base.Denotation, out _oboznType, out _code1, out _codeNum1, out _code2, out _codeNum2, out _postfix, out _prefix, out _variantIndex);
                }
            }
        }
        /*
        /// <summary>
        /// Раздел спецификации
        /// </summary>
        public SpecTypes.SpecificationType Specification
        {
            get
            {
                return base.BomSection;
            }
            set
            {
                _rowSortDataChanged = true;
                base.Specification = value;
            }
        }
        /// <summary>
        /// Номер позиции на чертеже
        /// </summary>
        public override int PositionNumber
        {
            get
            {
                return base.PositionNumber;
            }
            set
            {
                _rowSortDataChanged = true;
                base.PositionNumber = value;
            }
        }
        */
    }

    /// <summary>
    /// Коллекция разделов спецификации
    /// </summary>
    internal static class SpecTypes
    {
        /// <summary>
        /// Класс раздела спецификации
        /// </summary>
        public class SpecificationType
        {
            private string _specificationName;
            private int _parentID = 0;
            private int _id = 0;
            private int _priority = 0;
            private bool _bossType = false;
            private int _bossTypeID = 0;

            public SpecificationType(int id, string specName, int parentID)
            {
                _id = id;
                _specificationName = specName;
                _parentID = parentID;
                _bossTypeID = -1;
                _priority = 0;
                _bossType = false;
            }

            /// <summary>
            /// ID раздела спецификации
            /// </summary>
            public int ID
            {
                get
                {
                    return _id;
                }
            }
            /// <summary>
            /// ID родительского раздела спецификации
            /// </summary>
            public int ParentID
            {
                get
                {
                    return _parentID;
                }
            }
            /// <summary>
            /// Название раздела спецификации
            /// </summary>
            public string SpecificationName
            {
                get
                {
                    return _specificationName;
                }
            }
            /// <summary>
            /// Название головного раздела спецификации
            /// </summary>
            public string BossSpecName
            {
                get
                {
                    SpecificationType BossSpecType = SpecTypes.GetSpecificationData(BossTypeID);
                    return BossSpecType.SpecificationName;
                }
            }
            /// <summary>
            /// Код приоритета сортировки данного раздела спецификации
            /// </summary>
            public int SortPriority
            {
                set
                {
                    _priority = value;
                }
                get
                {
                    return _priority;
                }
            }
            /// <summary>
            /// Отображает, является ли данный раздел базовым
            /// </summary>
            public bool IsBossType
            {
                set
                {
                    _bossType = value;
                }
                get
                {
                    return _bossType;
                }
            }
            /// <summary>
            /// Принадлежит ли данный раздел к "Документации"
            /// </summary>
            public bool BelongToDocumentation
            {
                get
                {
                    if (BossTypeID == Defines.SpecificationTypes.DocumentationSpecID)
                        return true;

                    return false;
                }
            }
            /// <summary>
            /// ID раздела, базового для данного
            /// </summary>
            public int BossTypeID
            {
                get
                {
                    if (_bossTypeID <= 0)
                        _bossTypeID = SpecTypes.GetBossSpecID(_id);

                    return _bossTypeID;
                }
            }
            /// <summary>
            /// Возвращает название раздела спецификации
            /// </summary>
            /// <returns></returns>
            public override string ToString()
            {
                return SpecificationName;
            }
        }

        private static Dictionary<int, SpecificationType> _specDict = null;

        /// <summary>
        /// заполняет словарь разделов спецификации записями древовидного справочника Разделов спецификации
        /// </summary>
        /// <param name="masterRecID">ID головной записи</param>
        /// <param name="inputRefData">данные справочника разделов спецификации</param>
        /// <param name="mapToFill">словарь разделов спецификаций, который заполняется данными справчоника</param>
        /// <param name="priorVal">величина предыдущего приоритета</param>
        private static void FillChildParams(int masterRecID, Dictionary<int, Dictionary<string, string>> inputRefData, ref Dictionary<int, SpecificationType> mapToFill, ref int priorVal)
        {
            foreach (int RecID in inputRefData.Keys)
            {
                Dictionary<string, string> CurrentRecParams = inputRefData[RecID];

                int ParentID = Int32.Parse(CurrentRecParams[Defines.SpecificationTypes.ParentFieldGUID]);
                if (ParentID != masterRecID)
                    continue;

                string SpecName = CurrentRecParams[Defines.SpecificationTypes.NameFieldGUID];

                SpecificationType NewSpecType = new SpecificationType(RecID, SpecName, ParentID);

                if (masterRecID == Defines.SpecificationTypes.MasterSpecRecID)
                    NewSpecType.IsBossType = true;

                NewSpecType.SortPriority = priorVal;
                priorVal++;

                mapToFill.Add(RecID, NewSpecType);
                FillChildParams(RecID, inputRefData, ref mapToFill, ref priorVal);
            }
        }
        /// <summary>
        /// Загружает данные справочника разделов спецификаций
        /// </summary>
        /*public static void Load()
        {
            _specDict = new Dictionary<int, SpecificationType>();

            TFDGroup SpecGroup = CADGeneratorGroupESKD.ApiDocs.DocsGroups.GetItem(Defines.SpecificationTypes.GroupGUID) as TFDGroup;

            if (SpecGroup == null)
                throw new NullReferenceException("Не удалось найти группу справочника спецификаций!");

            int PriorityVal = 1;
            Dictionary<int, Dictionary<string, string>> SpecRefData = CADGeneratorGroupESKD.ApiDocs.GetAllGroupRecords(Defines.SpecificationTypes.GroupGUID, false);
            FillChildParams(Defines.SpecificationTypes.MasterSpecRecID, SpecRefData, ref _specDict, ref PriorityVal);
        }*/
        /// <summary>
        /// Возвращает ID базового раздела спецификации для текущего
        /// </summary>
        /// <param name="specID">ID раздела спецификации</param>
        /// <returns>ID базового раздела спецификации</returns>
        public static int GetBossSpecID(int specID)
        {
            if (!IsValidSpecificationType(specID))
                return 0;

            SpecificationType CurrentType = GetSpecificationData(specID);

            if (!CurrentType.IsBossType)
                return GetBossSpecID(CurrentType.ParentID);
            else
                return specID;
        }
        /// <summary>
        /// Возвращает данные о разделе спецификации
        /// </summary>
        /// <param name="specID">ID раздела спецификации</param>
        /// <returns>данные о разделе спецификации</returns>
        public static SpecificationType GetSpecificationData(int specID)
        {
            if (IsValidSpecificationType(specID))
                return _specDict[specID];

            throw new KeyNotFoundException("Нет раздела спецификации с ID = " + specID.ToString());
        }
        /// <summary>
        /// Проверяет на валидность раздел спецификации
        /// </summary>
        /// <param name="specID">ID проверямого раздела спецификации</param>
        /// <returns>результат проверки</returns>
        public static bool IsValidSpecificationType(int specID)
        {
            if (_specDict == null)
                throw new NullReferenceException("Данные о разделах спецификации не найдены!");

            if (!_specDict.ContainsKey(specID))
                return false;

            return true;
        }
        /// <summary>
        /// Проверяет, принадлежит ли раздел спецификации к документации
        /// </summary>
        /// <param name="specID">ID проверяемого раздела спецификации</param>
        /// <returns>результат проверки</returns>
        public static bool BelongToDocumentation(int specID)
        {
            int BossSpecID = GetBossSpecID(specID);

            if (BossSpecID == Defines.SpecificationTypes.DocumentationSpecID)
                return true;

            return false;
        }
    }
}
namespace Defines
{
    internal static class SortDefines
    {
        /// <summary>
        /// количество разрядов в коде приоритета головного раздела спецификации
        /// </summary>
        public const string BaseSpecTypeDigits = "0000";
        /// <summary>
        /// количество разрядов в номере позиции
        /// </summary>
        public const string DocPositionDigits = "00000";
        /// <summary>
        /// количество разрядов в коде приоритета раздела спецификации документа
        /// </summary>
        public const string CurrentSpecTypeDigits = "0000";
        /// <summary>
        /// количество разрядов в приоритете кода обозначения
        /// </summary>
        public const string CodePriorDigits = "000";
        /// <summary>
        /// количество разрядов в цифре после кода обозначения
        /// </summary>
        public const string CodeNumPriorDigits = "000";
        /// <summary>
        /// количество разрядов в цифре индекса исполнения
        /// </summary>
        public const string VariantIndexPriorDigits = "000";
        /// <summary>
        /// Максимальная длинна строки префикса обозначения
        /// </summary>
        public const int MaxPrefixLenght = 40;
    }

    /// <summary>
    /// Дифайны справочника разделов спецификаций 
    /// </summary>
    internal static class SpecificationTypes
    {
        /// <summary>
        /// ID записи главного раздела спецификации
        /// </summary>
        public static int MasterSpecRecID = 1;
        /// <summary>
        /// ID записи раздела спецификации "Документация"
        /// </summary>
        public static int DocumentationSpecID = 50;
        /// <summary>
        /// GUID группы "Разделы спецификации"
        /// </summary>
        public const string GroupGUID = "80BF50CE-C4F8-4E15-8F54-254174D2B4EB";
        /// <summary>
        /// Поле указывающее на родительский раздел спецификации
        /// </summary>
        public const string ParentFieldGUID = "69F9952E-5458-4368-B700-2818B3A2F896";
        /// <summary>
        /// Название спецификации
        /// </summary>
        public const string NameFieldGUID = "F769BBBE-9B88-4224-A2C5-2A55D0B9570F";
    }

    /// <summary>
    /// Дифайны параметров документов
    /// </summary>
    internal static class DocParams
    {
        /// <summary>
        /// Значения параметра ClassEx документа (Дополнительный параметр док-та)
        /// </summary>
        public enum ClassEx
        {
            /// <summary>
            /// ведомость спецификаций
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_SPECIFICATION = 19964,
            /// <summary>
            /// ведомость ссылочных документов
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_REF_DOC = 19965,
            /// <summary>
            /// ведомость покупных изделий
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_BOUGHT_ARTICLE = 19966,
            /// <summary>
            /// ведомость разрешения применения покупных изделий
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_PERMIT_USE_BOUGHT_ARTICLE = 19967,
            /// <summary>
            /// ведомость держателей подлинников
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_ORIGINAL = 19968,
            /// <summary>
            /// ведомости технического предложения, эскизного и технического проектов
            /// </summary>
            CLASSOBJ_EX_BOM_LIST_PROPOSITION = 19969,
            /// <summary>
            /// групповая ведомость спецификаций
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_LIST_SPECIFICATION = 19970,
            /// <summary>
            /// групповая ведомость покупных изделий
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_LIST_BOUGHT_ARTICLE = 19971,
            /// <summary>
            /// простая спецификация
            /// </summary>
            CLASSOBJ_EX_BOM_SIMPLE = 19972,
            /// <summary>
            /// групповая спецификация по типу А
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_A = 19973,
            /// <summary>
            /// групповая спецификация по типу Б
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_B = 19974,
            /// <summary>
            /// плазовая спецификация
            /// </summary>
            CLASSOBJ_EX_BOM_SIMPLE_PLAZ = 19975,
            /// <summary>
            /// групповая спецификация по типу В
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_V = 19976,
            /// <summary>
            /// групповая спецификация по типу Г
            /// </summary>
            CLASSOBJ_EX_BOM_GROUP_G = 19977
        }

        /// <summary>
        /// "Идентификатор объекта"
        /// </summary>
        public const string DOCID = "B53FAB61-FB5E-4C5E-B3B2-A714A26A137B";
        /// <summary>
        /// "Входит в спецификацию"
        /// </summary>
        public const string NEEDBOM = "57237DFC-D192-4E2F-8E09-F8C1CDBECB93";
        /// <summary>
        /// "Входит в состав изделия"
        /// </summary>
        public const string PARTOFPRODUCT = "DD88DC39-9B23-4025-85A7-8B841D8F2965";
        /// <summary>
        /// "Раздел спецификации"
        /// </summary>
        public const string BOMGROUP = "DC41ACEE-4941-4B0D-9716-454BF887529A";
        /// <summary>
        /// "Наименование объекта"
        /// </summary>
        public const string DOCNAME = "2C30F1FA-7933-42A2-BA10-E0547080C144";
        /// <summary>
        /// "Обозначение объекта"
        /// </summary>
        public const string DOCCODE = "FF5BBC2C-E6AB-4A96-A3BA-E0BC2B2D1004";
        /// <summary>
        /// "Зона на чертеже"
        /// </summary>
        public const string SHEETZONE = "F2B636DC-627A-43A1-A7B9-D401548FAF20";
        /// <summary>
        /// "Примечание спецификации"
        /// </summary>
        public const string BOMNOTE = "58C00440-6D3C-42B3-A9AC-E0B1D8BF3A15";
        /// <summary>
        /// "Формат чертежа детали"
        /// </summary>
        public const string SHEETFORMAT = "4F18ABBE-A407-46AA-968F-ED5B3726FC2A";
        /// <summary>
        /// "Автоматическое нумерование позиции"
        /// </summary>
        public const string SHEETPOSITIONFLAG = "4B59CBC4-D4B8-4081-8870-41738BD3481E";
        /// <summary>
        /// "Позиция на чертеже"
        /// </summary>
        public const string SHEETPOSITION = "DE7F484C-9D3D-426C-BB66-F3B23382416B";
        /// <summary>
        /// "Количество"
        /// </summary>
        public const string REFCOUNT = "89C72DC2-5343-403D-A918-754B20AE2DF0";
        /// <summary>
        /// "Количество экземпляров"
        /// </summary>
        public const string INSTANCE_AMOUNT = "DA288BF7-C9B8-4462-8A74-CBD41E935F10";
        /// <summary>
        /// "Идентификатор файла объекта"
        /// </summary>
        public const string DOCFILEID = "8EFF44C2-447D-49AB-A3F4-946D24B855AA";
        /// <summary>
        /// "Автор"
        /// </summary>
        public const string CREATEUSERID = "8DEF9DF7-C238-4031-81CF-E79FDBC87CFC";
        /// <summary>
        /// "Дополнительный класс объекта"
        /// </summary>
        public const string OBJECTCLASSEX = "AFEA49A1-29A8-4935-9715-C5DFD6AF2654";
    }

    /// <summary>
    /// Дифайны шаблонов регулярных выражений для определения типа обозначения
    /// </summary>
    internal static class OboznTypes
    {
        /// <summary>
        /// Тип строки занесенной в обозначение
        /// ____________________________________
        /// УСЛОВНЫЕ ОБОЗНАЧЕНИЯ В КОММЕНТАРИЯХ:
        /// 'БУКВЫ.XXX.XXXX' - префикс обозначения (например НКШР.467444.023)
        /// 'КОД' - код обозначения (например ФО или Э2)
        /// '-XX' - номер исполнения (например -01 или -02)
        /// </summary>
        public enum FormatType
        {
            /// <summary>
            /// "БУКВЫ.XXX.XXXX"
            /// </summary>
            Full,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX-XX"
            /// </summary>
            Full_Var,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД"
            /// </summary>
            FullCode,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД-XX"
            /// </summary>
            FullCode_Var,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX-ЛУ"
            /// </summary>
            FullPostfix_LU,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX-ЛУ-XX"
            /// </summary>
            FullPostfix_LU_Var,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX УД"
            /// </summary>
            FullPostfix_UD,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX УД-XX"
            /// </summary>
            FullPostfix_UD_Var,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД-ЛУ"
            /// </summary>
            FullCodePostfix_LU,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД-ЛУ-XX"
            /// </summary>
            FullCodePostfix_LU_Var,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД-УД"
            /// </summary>
            FullCodePostfix_UD,
            /// <summary>
            /// "БУКВЫ.XXX.XXXX КОД-УД-XX"
            /// </summary>
            FullCodePostfix_UD_Var,

            /// <summary>
            /// "ЛУ"
            /// </summary>
            ShortPostfix_LU,
            /// <summary>
            /// "ЛУ-XX"
            /// </summary>
            ShortPostfix_LU_Var,
            /// <summary>
            /// "УД"
            /// </summary>
            ShortPostfix_UD,
            /// <summary>
            /// "УД-XX"  
            /// </summary>
            ShortPostfix_UD_Var,
            /// <summary>
            /// "КОД"
            /// </summary>
            ShortCode,
            /// <summary>
            /// "КОД-XX"
            /// </summary>
            ShortCode_Var,
            /// <summary>
            /// "КОД-ЛУ"
            /// </summary>
            ShortCodePostfix_LU,
            /// <summary>
            /// "КОД-ЛУ-XX"
            /// </summary>
            ShortCodePostfix_LU_Var,
            /// <summary>
            /// "КОД-УД"
            /// </summary>
            ShortCodePostfix_UD,
            /// <summary>
            /// "КОД-УД-XX"
            /// </summary>
            ShortCodePostfix_UD_Var,
            /// <summary>
            /// тип обозначения не распознан :(
            /// </summary>
            Unknown,
        }

        /// <summary>
        /// Определяет, содержит ли данный тип обозначения код
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool HaveCode(FormatType type)
        {
            switch (type)
            {
                case FormatType.FullCode:
                case FormatType.FullCode_Var:
                case FormatType.FullCodePostfix_LU:
                case FormatType.FullCodePostfix_LU_Var:
                case FormatType.FullCodePostfix_UD:
                case FormatType.FullCodePostfix_UD_Var:
                case FormatType.ShortCode:
                case FormatType.ShortCode_Var:
                case FormatType.ShortCodePostfix_LU:
                case FormatType.ShortCodePostfix_LU_Var:
                case FormatType.ShortCodePostfix_UD:
                case FormatType.ShortCodePostfix_UD_Var:
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Определяет, содержит ли данный тип обозначения номер исполнения
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool IsVariant(FormatType type)
        {
            switch (type)
            {
                case FormatType.Full_Var:
                case FormatType.FullCode_Var:
                case FormatType.FullCodePostfix_LU_Var:
                case FormatType.FullCodePostfix_UD_Var:
                case FormatType.FullPostfix_LU_Var:
                case FormatType.FullPostfix_UD_Var:
                case FormatType.ShortCode_Var:
                case FormatType.ShortCodePostfix_LU_Var:
                case FormatType.ShortCodePostfix_UD_Var:
                case FormatType.ShortPostfix_LU_Var:
                case FormatType.ShortPostfix_UD_Var:
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Определяет, содержит ли данный тип обозначения постфикс
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool HavePostfix(FormatType type)
        {
            switch (type)
            {
                case FormatType.FullCodePostfix_LU:
                case FormatType.FullCodePostfix_LU_Var:
                case FormatType.FullCodePostfix_UD:
                case FormatType.FullCodePostfix_UD_Var:
                case FormatType.FullPostfix_LU:
                case FormatType.FullPostfix_LU_Var:
                case FormatType.FullPostfix_UD:
                case FormatType.FullPostfix_UD_Var:
                case FormatType.ShortCodePostfix_LU:
                case FormatType.ShortCodePostfix_LU_Var:
                case FormatType.ShortCodePostfix_UD:
                case FormatType.ShortCodePostfix_UD_Var:
                case FormatType.ShortPostfix_LU:
                case FormatType.ShortPostfix_LU_Var:
                case FormatType.ShortPostfix_UD:
                case FormatType.ShortPostfix_UD_Var:
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Определяет, имеет ли данный тип обозначения посфикс ЛУ
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool HavePostfixLU(FormatType type)
        {
            switch (type)
            {
                case FormatType.FullCodePostfix_LU:
                case FormatType.FullCodePostfix_LU_Var:
                case FormatType.FullPostfix_LU:
                case FormatType.FullPostfix_LU_Var:
                case FormatType.ShortCodePostfix_LU:
                case FormatType.ShortCodePostfix_LU_Var:
                case FormatType.ShortPostfix_LU:
                case FormatType.ShortPostfix_LU_Var:
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Определяет, имеет ли данный тип обозначения посфикс УД
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool HavePostfixUD(FormatType type)
        {
            switch (type)
            {
                case FormatType.FullCodePostfix_UD:
                case FormatType.FullCodePostfix_UD_Var:
                case FormatType.FullPostfix_UD:
                case FormatType.FullPostfix_UD_Var:
                case FormatType.ShortCodePostfix_UD:
                case FormatType.ShortCodePostfix_UD_Var:
                case FormatType.ShortPostfix_UD:
                case FormatType.ShortPostfix_UD_Var:
                    return true;
            }

            return false;
        }
        /// <summary>
        /// Определяет, записано ли обозначение в краткой форме
        /// </summary>
        /// <param name="type">тип обозначения</param>
        /// <returns>результат проверки</returns>
        public static bool IsShort(FormatType type)
        {
            switch (type)
            {
                case FormatType.ShortCode:
                case FormatType.ShortCode_Var:
                case FormatType.ShortCodePostfix_LU:
                case FormatType.ShortCodePostfix_LU_Var:
                case FormatType.ShortCodePostfix_UD:
                case FormatType.ShortCodePostfix_UD_Var:
                case FormatType.ShortPostfix_LU:
                case FormatType.ShortPostfix_LU_Var:
                case FormatType.ShortPostfix_UD:
                case FormatType.ShortPostfix_UD_Var:
                    return true;
            }

            return false;
        }

        /* Комментарии по обозначениям см. в enum FormatType */

        public const string PrefixPattern = @"(\S+(\.[0-9]+)+)"; //шаблон префикса
        public const string StartPart = @"\A"; //начало обозначение
        public const string EndPart = @"\Z"; //конец обозначения
        public const string VarEndPart = @"(-[0-9]*)\Z"; //окончания обозначения исполнения
        public const string CodePart = @"([\s|-]?[а-я]+[\s|-]?[0-9]*([а-я]+[0-9]*)?)"; //шаблон кода обозначения
        public const string PartLU = @"([\s|-]?(ЛУ)|(УЛ))"; //шаблон '-ЛУ/УЛ'
        public const string PartUD = @"([\s|-]?УД)"; //шаблон '-УД'

        public static readonly string Full = String.Concat(StartPart, PrefixPattern, EndPart);
        public static readonly string Full_Var = Full.Replace(EndPart, VarEndPart);
        public static readonly string FullCode = String.Concat(StartPart, PrefixPattern, CodePart, EndPart); //вызывать после проверки на ЛУ/УД!!!
        public static readonly string FullCode_Var = FullCode.Replace(EndPart, VarEndPart); //вызывать после проверки на ЛУ/УД!!!
        public static readonly string FullPostfix_LU = String.Concat(StartPart, PrefixPattern, PartLU, EndPart);
        public static readonly string FullPostfix_LU_Var = FullPostfix_LU.Replace(EndPart, VarEndPart);
        public static readonly string FullPostfix_UD = String.Concat(StartPart, PrefixPattern, PartUD, EndPart);
        public static readonly string FullPostfix_UD_Var = FullPostfix_UD.Replace(EndPart, VarEndPart);
        public static readonly string FullCodePostfix_LU = String.Concat(StartPart, PrefixPattern, CodePart, PartLU, EndPart);
        public static readonly string FullCodePostfix_LU_Var = FullCodePostfix_LU.Replace(EndPart, VarEndPart);
        public static readonly string FullCodePostfix_UD = String.Concat(StartPart, PrefixPattern, CodePart, PartUD, EndPart);
        public static readonly string FullCodePostfix_UD_Var = FullCodePostfix_UD.Replace(EndPart, VarEndPart);

        public static readonly string ShortPostfix_LU = String.Concat(StartPart, PartLU, EndPart);
        public static readonly string ShortPostfix_LU_Var = ShortPostfix_LU.Replace(EndPart, VarEndPart);
        public static readonly string ShortPostfix_UD = String.Concat(StartPart, PartUD, EndPart);
        public static readonly string ShortPostfix_UD_Var = ShortPostfix_UD.Replace(EndPart, VarEndPart);
        public static readonly string ShortCode = String.Concat(StartPart, CodePart, EndPart); //вызывать после проверки на ЛУ/УД!!!
        public static readonly string ShortCode_Var = ShortCode.Replace(EndPart, VarEndPart); //вызывать после проверки на ЛУ/УД!!!
        public static readonly string ShortCodePostfix_LU = String.Concat(StartPart, CodePart, PartLU, EndPart);
        public static readonly string ShortCodePostfix_LU_Var = ShortCodePostfix_LU.Replace(EndPart, VarEndPart);
        public static readonly string ShortCodePostfix_UD = String.Concat(StartPart, CodePart, PartUD, EndPart);
        public static readonly string ShortCodePostfix_UD_Var = ShortCodePostfix_UD.Replace(EndPart, VarEndPart);
    }

    /// <summary>
    /// Дифайны переменный документа отчета
    /// </summary>
    internal static class ReportDocVars
    {
        /// <summary>
        /// Кол-во строк до заголовка раздела спецификации по умолчанию
        /// </summary>
        public const int LinesBeforeHeaderDefault = 2;
        /// <summary>
        /// Кол-во строк после заголовка раздела спецификации по умолчанию
        /// </summary>
        public const int LinesAfterHeaderDefault = 1;
        /// <summary>
        /// Число строк до "Обозн. исполн." по умолчанию
        /// </summary>
        public const int LinesBeforeSectionDefault = 1;
        /// <summary>
        /// Число строк после "Обозн. исполн." по умолчанию
        /// </summary>
        public const int LinesAfterSectionDefault = 1;

        /// <summary>
        /// Имя переменной, хранящей кол-во строк до названия раздела спецификации
        /// </summary>
        public const string LinesBeforeHeader = "$lines_before_header";
        /// <summary>
        /// Имя переменной, хранящей кол-во строк после названия раздела спецификации
        /// </summary>
        public const string LinesAfterHeader = "$lines_after_header";
        /// <summary>
        /// Имя переменной, хранящей число пустых строк до "Обозн. исполн."
        /// </summary>
        public const string LinesBeforeSection = "$lines_before_section";
        /// <summary>
        /// Имя переменной, хранящей число пустых строк после "Обозн. исполн."
        /// </summary>
        public const string LinesAfterSection = "$lines_after_section";
        /// <summary>
        /// Имя переменной, хранящей список кодов обозначений
        /// </summary>
        public const string CodePriority = "$code_priority";
        /// <summary>
        /// Имя параграф текста, в который заносятся данные спецификации
        /// </summary>
        public const string DataTableParText = "$data_table";
        /// <summary>
        /// Имя параграф текста, в который заносятся порядковые номера спецификаций
        /// </summary>
        public const string NumHeaderParText = "$num_header";
        /// <summary>
        /// Имя переменной, созданной для хранения номеров страниц исполнений (10-19... и т.д.)
        /// </summary>
        public const string VariantSectionsPageNumbers = "$var_section_pages";

        /// <summary>
        /// Формат имени директории, в которую записывается документ отчета
        /// </summary>
        public const string GrbDirIdPattern = @"tfd([0-9]+)\.grb_files\Z";
    }

    /// <summary>
    /// Дифайны справочника Пользователи
    /// </summary>
    internal static class UsersGroup
    {
        /// <summary>
        /// GUID справочника Пользователи
        /// </summary>
        public const string GroupGUID = "8EE861F3-A434-4C24-B969-0896730B93EA";
        /// <summary>
        /// GUID поля Полное имя пользователя
        /// </summary>
        public const string UserNameParamGUID = "B2C4C5C7-358B-4A3B-BEF5-82BCDF4CF1FA";
    }
}
