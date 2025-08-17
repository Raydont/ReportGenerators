using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.References.Documents;
using System.Threading;
using TFlex.DOCs.Model;

namespace ListElementsReport
{
    public class ReportParameters
    {
        //Параметры файла "RefDes";"Type";"Value";"Допуск";"TU";"Обозначение";"Диап. номин."
        public List<ElementPki> Elements = new List<ElementPki>();

        public string DenotationDevice;  // Обозначение документа
        public string NameDevice; //Наименование устройства
        public string FirstUse;          //Первичная применяемость
        //public List<string> FileData = new List<string>();          //Первоначальные данные из загружаемого файла atr
        public List<FileData> FileData = new List<FileData>();          //Первоначальные данные из загружаемого файла xls
        public List<DataBaseElements> dBElements = new List<DataBaseElements>();
        public string filePath = string.Empty;
        public List<LinkedObjectDB> LinkedObjects = new List<LinkedObjectDB>();    //Таблица связанных объектов 
        public int CountHeadRows;
        public int BeginPage;

        public string AuthorBy;          // Разраю.
        public string MusteredBy;          // Пров.
        public string NControlBy;          // Н.Контр.
        public string ApprovedBy;          // Утв.

        public DateTime AuthorByDate;
        public DateTime MusteredByDate;
        public DateTime NControlByDate;
        public DateTime ApprovedByDate;

        public bool AuthorCheck;
        public bool MusteredBCheck;
        public bool NControlByCheck;
        public bool ApprovedByCheck;

        public char Litera1;
        public char Litera2;
        public string NumberSol;
        public string Code;
        public bool SelectTKS;
        public bool IsLoadKre;
        public bool IsManyDevice;
        public bool IsKonstr;
    }

    public class FileData
    {
        public string Admission;
        public string Value;
        public string Designator;
        public string Type;
        public string Tks;
        public string Наименование;
        public string Обозначение;
        public string Примечание;

        public FileData(ListElementReport.Xls xls, NumbersColumns numbersColumns, int row)
        {
            Type = GetValue(xls, numbersColumns.Type, row);
            Designator = GetValue(xls, numbersColumns.Designator, row);
            Admission = GetValue(xls, numbersColumns.Admission, row);
            Value = GetValue(xls, numbersColumns.Value, row).Replace(".", ",");
            Tks = GetValue(xls, numbersColumns.Tks, row);
            Наименование = GetValue(xls, numbersColumns.Наименование, row);
            Обозначение = GetValue(xls, numbersColumns.Обозначение, row);
            Примечание = GetValue(xls, numbersColumns.Примечание, row);
        }

        public FileData(string[] data, NumbersColumns numbersColumns)
        {
            for (int i = 0; i < data.Count(); i++)
            {
                if (data[i][0] == '\"' && data[i][data[i].Length - 1] == '\"')
                {
                    data[i] = data[i].Substring(1, data[i].Length - 2);
                }
            }
            Type = data[numbersColumns.Type];
            Designator = data[numbersColumns.Designator];
            Admission = numbersColumns.Admission != -1 ? data[numbersColumns.Admission] : string.Empty;
            Value = data[numbersColumns.Value];
            Tks = numbersColumns.Tks != -1 ? data[numbersColumns.Tks] : string.Empty;
        }

        private string GetValue(ListElementReport.Xls xls, int col, int row)
        {
            if (col != 0 && col != -1)
            {
                var range = xls[col, row];
                if (range != null)
                {
                    return range.Text.Trim();
                }
            }
            return string.Empty;
        }
    }

    public class NumbersColumns
    {
        public int Type = -1;
        public int Designator = -1;
        public int Value = -1;
        public int Admission = -1;
        public int Tks = -1;
        public int Наименование = -1;
        public int Обозначение = -1;
        public int Примечание = -1;

        public NumbersColumns(string firstString)
        {
            var headers = firstString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < headers.Length; i++)
            {
                switch (headers[i].Replace("\"", "").Trim())
                {
                    case "RefDes":
                        Designator = i;
                        break;
                    case "Type":
                        Type = i;
                        break;
                    case "Value":
                        Value = i;
                        break;
                    case "Допуск":
                        Admission = i;
                        break;
                    case "ТКС":
                        Tks = i;
                        break;
                    case "TKC":
                        Tks = i;
                        break;
                    default:
                        break;
                }
            }
        }
        public NumbersColumns(ListElementReport.Xls xls)
        {
            for (int i = 1; i < 12; i++)
            {
                var range = xls[i, 1];
                var header = range != null && !string.IsNullOrEmpty(range.Text) ? (string)range.Text.Trim().ToLower() : string.Empty;
                switch (header)
                {
                    case "designator":
                        Designator = i;
                        break;
                    case "comment":
                        Type = i;
                        break;
                    case "part number":
                        Type = i;
                        break;
                    case "value":
                        Value = i;
                        break;
                    case "допуск":
                        Admission = i;
                        break;
                    case "обозначение":
                        Обозначение = i;
                        break;
                    case "наименование":
                        Наименование = i;
                        break;
                    case "примечание":
                        Примечание = i;
                        break;
                    case "ткс":
                        Tks = i;
                        break;
                    case "tkc":
                        Tks = i;
                        break;
                    default:
                        break;
                }
            }
        }
    }

public class ElementPki
    {
        private string posDenotation;
        private string type;
        private string value;
        private string admission;
        private string tks;
        private string tU;
        private string denotation;
        public string OldName;
        public string Name = string.Empty;
        public string Comment = string.Empty;
        private string letterPosDenotation;
        private int count;
        public List<DataBaseElements> dbElements = new List<DataBaseElements>();
        public bool RezWithOutNominal = false;
        public bool SelectionObject;
        public bool MainSelectionObject;
        public bool Assembly = false;
        public ReferenceObject NomObj = null;
        public string ClassObject = string.Empty;
        public string Preparation = string.Empty;
        public string NameWithOutBlank;
        public string ErrorValueText = string.Empty;
        public string Parent = string.Empty;
        public bool AddedNotLinkedObj = false;
        public string DenotationDevice = string.Empty;
        public string NameDevice = string.Empty;
        public string Remark = string.Empty;

        public double SizeMaterial
        {
            get
            {
                string patternSize = @"\A((\d*(\.|\,)\d*)|\d*)";
                Regex regexPattern = new Regex(patternSize);
                try
                {
                    if (regexPattern.Match(Comment).Success)
                    {

                        return Convert.ToDouble(regexPattern.Match(Comment).Value.Trim().Replace(',', '.'));
                    }
                    else return 0;
                }
                catch
                {
                    return 0;
                }
            }
        }

        public string PosDenotation
        {
            get
            {
                return posDenotation;
            }

            set
            {
                if (value != string.Empty && value.Length >= 2 && value.Trim().Last() == '"' && value.Trim().First() == '"')
                {
                    var t = value.Remove(0, 1);
                    t = t.Remove(t.Length - 1);

                    this.posDenotation = t;
                }
                else
                    this.posDenotation = value;
            }
        }

        public string Type
        {
            get
            {
                return type;
            }
            set
            {
                if (value != string.Empty && value.Length >= 2 && value.Trim().Last() == '"' && value.Trim().First() == '"')
                {
                    var t = value.Remove(0, 1);
                    t = t.Remove(t.Length - 1);

                    this.type = t;
                }
                else
                    this.type = value;
            }
        }

        public string Value
        {
            get
            {
                return value;
            }

            set
            {
                try
                {
                    if (value != string.Empty && value.Length >= 1)
                    {
                        var t = value;
                        var patternValue = @"(\d*(\.|\,|/)\d*)|\d*";

                        var val = Regex.Match(t, patternValue);

                        var unitMeasure = string.Empty;
                        TextNormalizer textNorm = new TextNormalizer();
                        if (val.Success)
                            unitMeasure = t.Replace(val.Value.Trim().ToLower(), "");
                        else
                        {
                            this.value = string.Empty;
                            return;
                        }

                        //Выбор разрешенных единиц измерений
                        switch (textNorm.GetNormalForm(unitMeasure))
                        {
                            case "ма":
                                this.value = val.Value.Replace(".", ",") + " мА";
                                break;
                            case "мкф":
                                this.value = val.Value.Replace(".", ",") + " мкФ";
                                break;
                            case "мк":
                                this.value = val.Value.Replace(".", ",") + " мкФ";
                                break;
                            case "к":
                                this.value = val.Value.Replace(".", ",") + " кОм";
                                break;
                            case "k":
                                this.value = val.Value.Replace(".", ",") + " кОм";
                                break;
                            case "ком":
                                this.value = val.Value.Replace(".", ",") + " кОм";
                                break;
                            case "мом":
                                this.value = val.Value.Replace(".", ",") + " МОм";
                                break;
                            case "м":
                                this.value = val.Value.Replace(".", ",") + " МОм";
                                break;
                            case "гн":
                                this.value = val.Value.Replace(".", ",") + " гОм";
                                break;
                            case "г":
                                this.value = val.Value.Replace(".", ",") + " гОм";
                                break;
                            case "мкгн":
                                this.value = val.Value.Replace(".", ",") + " мкГн";
                                break;
                            case "нгн":
                                this.value = val.Value.Replace(".", ",") + " нГн";
                                break; 
                            case "мкг":
                                this.value = val.Value.Replace(".", ",") + " мкГн";
                                break;
                            case "ом":
                                this.value = val.Value.Replace(".", ",") + " Ом";
                                break;
                            case "пф":
                                this.value = val.Value.Replace(".", ",") + " пФ";
                                break;
                            case "":
                                this.value = val.Value.Replace(".", ",") + " пФОм";
                                break;
                            default:
                                {
                                    if (posDenotation.ToLower().Contains("v") || posDenotation.ToLower().Contains("a") || posDenotation.ToLower().Contains("b") || posDenotation.ToLower().Contains("d"))
                                        this.value = string.Empty;
                                    else
                                    {
                                        //     MessageBox.Show("Некорректный номинал или единица измерения: " + value +
                                        //" в объекте [" + posDenotation + "  " + Type + " " + Comment + " " + TU + "]" +
                                        //"\nДопустимые значения единиц измерения:\nмк\nк\nм\nг\nмкГ\nнГн\nма", "Ошибка!!!",
                                        //MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        //     Clipboard.SetText(t + " " + value);
                                        ErrorValueText = t + " " + value;
                                    }
                                    break;
                                }
                        }
                    }
                    else
                        this.value = string.Empty;

                }
                catch
                {
                    //MessageBox.Show(value);
                }
            }
        }

        public string Admission
        {
            get
            {


                return admission;
            }

            set
            {
                if (value != string.Empty && value.Length >= 2 && value.Trim().Last() == '"' && value.Trim().First() == '"')
                {
                    var t = value.Replace(('"').ToString(), "").Trim();


                    this.admission = t;
                }
                else
                {
                    this.admission = value.Replace(('"').ToString(), "").Trim();
                }
            }
        }

        public string TU
        {
            get
            {
                return tU;
            }

            set
            {
                this.tU = value;
            }
        }

        public string Denotation
        {
            get
            {

                return denotation;
            }

            set
            {
                if (value != string.Empty && value.Length >= 2 && value.Trim().Last() == '"' && value.Trim().First() == '"')
                {
                    var t = value.Remove(0, 1);
                    t = t.Remove(t.Length - 1);

                    this.denotation = t;
                }
                else
                    this.denotation = value;
            }
        }



        public int Count
        {
            get
            {
                return count;
            }

            set
            {

                count = value;

            }
        }

        public string LetterPosDenotation
        {
            get
            {
                var positions = PosDenotation.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                Regex regexLetter = new Regex(@"\D{1,}");
                return regexLetter.Match(positions[0].Trim()).Value.Trim();
            }


        }

        public string Tks
        {
            get
            {
                return tks;
            }

            set
            {
                if (value != string.Empty && value.Length >= 2 && value.Trim().Last() == '"' && value.Trim().First() == '"')
                {
                    var t = value.Replace(('"').ToString(), "").Trim();


                    this.tks = t;
                }
                else
                {
                    this.tks = value.Replace(('"').ToString(), "").Trim();
                }
            }
        }

        public List<string> ModifyPosDenotation()
        {
            var listResult = new List<string>();
            var positions = PosDenotation.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            var seq = new Dictionary<int, string>();
            int lengthLine = 9;

            if (positions.ToList().Count <= 2)
            {
                listResult.Add(PosDenotation);
                return listResult;
            }

            foreach (var pos in positions)
            {
                try
                {
                    string pattern = @"\p{L}(\d+)(?:\(\p{L}(\d+)\))?";
                    Match matchDigit = Regex.Match(pos.Trim(), pattern);
                    if (matchDigit.Success)
                    {
                        var number = matchDigit.Groups[2].Success ? matchDigit.Groups[2].Value : matchDigit.Groups[1].Value;

                        try
                        {
                            seq.Add(Convert.ToInt32(number), pos);
                        }
                        catch
                        {

                        }

                    }
                    else
                    {
                        MessageBox.Show("Неправильное значение Поз. обозначение\n" + pos + "\nу объекта" + Name + "\nОбратитесь в отдел 911", "Ошибка №1!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        var Lstr = new List<string>();
                        Lstr.Add("0");
                        return Lstr;
                    }
                }
                catch
                {
                    MessageBox.Show("Неправильное значение Поз. обозначение\n" + pos + "\nу объекта" + Name + "\nОбратитесь в отдел 911", "Ошибка №2!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    var Lstr = new List<string>();
                    Lstr.Add("0");
                    return Lstr;
                }
            }
            string resultString = string.Empty;
            int order = 0;
            for (int i = 0; i < seq.Count; i++)
            {

                if ((i + 1) < seq.Count && (seq.Keys.ToList()[i] + 1) == seq.Keys.ToList()[i + 1])
                {
                    if (order == 0)
                    {
                        if (i == 0)
                            resultString = seq.Values.ToList()[i];
                        else
                            resultString += "," + seq.Values.ToList()[i];
                    }
                    else
                    {

                    }
                    order++;
                }
                else
                {
                    if (order == 0)
                    {
                        if (i == 0)
                            resultString = seq.Values.ToList()[i];
                        else
                            resultString += "," + seq.Values.ToList()[i];
                    }
                    else
                    {
                        if (order == 1)
                        {
                            resultString += "," + seq.Values.ToList()[i];
                        }
                        else
                        {
                            if (order >= 2)
                            {
                                resultString += "-" + seq.Values.ToList()[i];
                            }
                        }

                    }

                    order = 0;
                }
                if (i == seq.Count - 1 && !resultString.Contains(seq.Values.ToList().Last()))
                {
                    resultString += "-" + seq.Values.ToList()[i];
                }
            }


            string bufStr = string.Empty;
            listResult.Add(string.Empty);

            for (int i = 0; i < resultString.Length; i++)
            {

                bufStr += resultString[i].ToString();
                if (resultString[i].ToString() == ",")
                {
                    if ((listResult[listResult.Count - 1].Length + bufStr.Length) <= lengthLine || listResult[listResult.Count - 1] == string.Empty)
                    {
                        listResult[listResult.Count - 1] += bufStr;
                        bufStr = string.Empty;
                    }
                    else
                    {
                        listResult.Add(bufStr);
                        bufStr = string.Empty;
                    }
                }
            }

            if ((listResult[listResult.Count - 1].Length + bufStr.Length) <= lengthLine)
            {
                listResult[listResult.Count - 1] += bufStr;
            }
            else
            {
                if (listResult.Count() == 1 && listResult[0].Trim() == string.Empty)
                {
                    listResult[listResult.Count - 1] += bufStr;
                }
                else
                {
                    listResult.Add(bufStr);
                }
            }

            if (listResult[0] == string.Empty)
            {
                listResult[0] = "0";
            }
            return listResult;
        }
    }

    public class DataBaseElements
    {
        public string Id;
        public string PartNumber;
        public string NormalPartNumber;
        public string Comment;
        public string Description;
        public string Value;
        public string Admission;
        public string OptionalParameter;
        public string TU;
        public List<string> RangeValue = new List<string>();
        public string Measurement;
        public string InstallationVariant;
        public List<LinkedObjectDB> LinkedObject = new List<LinkedObjectDB>();
        public bool NotWrite = false;

        public List<string> NormalRangeValue
        {
            get
            {
                var listStr = new List<string>();
                TextNormalizer textNorm = new TextNormalizer();
                foreach (var val in RangeValue)
                {
                    listStr.Add(textNorm.GetNormalForm(val));
                }
                return listStr;
            }

            set
            {
                RangeValue = value;
            }
        }
    }

    public class LinkedObjectDB
    {
        public int Id;
        public string CodePKI;
        public int Count;
        public string Name;
        public string Denotation;
        public string Preparation;
        public string Comment;
        public string Type;

    }

    public class SortPKIElements
    {
        public ElementPki ElPKI = new ElementPki();
        public int Measure;
        public double Nominal;
        public string FirstPartName;
        public string LetterSecondPartNameKond;
        public double SecondPartName;
        public string SecondNameMcx;
        public int ThirdPartName;

        public SortPKIElements(ElementPki elPKI)
        {

            ElPKI = elPKI;
            if (elPKI.Type != "Прочее изделие")
                return;

            string name = name = elPKI.OldName;

            if (name == null || name == string.Empty)
                  name = elPKI.Name;

            
            string measurePattern = @"(мкФ|пФ|Ом|кОм|МОм)";
            string nominalPattern = @"(?<=(-|-\())((\d+)[.,]?\d*)(?=\s?(пФ|мкФ|Ом|кОм|МОм|/))";
            string firstPartNamePattern = @".*(?=-\d+[.,/]?\d*\s?(пФ|мкФ|Ом|кОм|МОм))";
            //    string kondensatorSecondPattern = @"(((?<=[A-ZА-Я]{1})\d{1,})|(\d{1,3}))(?=(\s?[A-ZА-Я]{1})?-((\d+)[.,]?\d*\s?(мкФ|пФ)))";
            string kondensatorSecondPattern = @"(((?<=[A-ZА-Я]{1})\d{1,})|(\d{1,3}(,\d{1,})?))(?=(\s?[A-ZА-Яa-zа-я]{1,2})?(-|-\()((\d+)[.,]?\d*\s?(мкФ|пФ)))";
            string kondensatorFirstPartPattern = @"(?<=(Конденсатор)\s)([A-ZА-Я]{1}\d{1,3}([A-ZА-Я]{1})?-\d{1,3}([a-zа-я]{1}|\s?\W[A-ZА-Я]\W|-)|(КТ4-25В)|(КТ4-25в))";
            // string kondensatorSecondPartLetterPattern = @"(([A-ZА-Я])|())(?=(\d{1,})?-((\d+)[.,]?\d*\s?(мкФ|пФ)))";
            string kondensatorSecondPartLetterPattern = @"(([A-ZА-Яa-zа-я]{1,2})|())(?=(\d{1,})?(-|-\()((\d+)[.,]?\d*\s?(мкФ|пФ)))";

            //Регулярные выражения для микросхем, транзисторов и диодов
            // string mcxSerialPattern = @"(?<=Микросхема\s)\d{1,}";
            string mcxSerialPattern = @"(?<=(Микросхема\s)|(Диод\s2Д)|(Транзистор\s2Т))\d{1,}";
            string mcxNumberPattern = @"(?<=Микросхема\s\d{3,4}[А-Я]{1,})\d{1,}";
            // string mcxSecondNamePattern = @"(?<=Микросхема\s\d{3,4})[А-Я]{1,}";
            string mcxSecondNamePattern = @"(?<=(Микросхема\s\d{3,4})|(Диод\s2Д\d{3,4})|(Транзистор\s2Т\d{3,4}))[А-Я]{1,}";

            string nameOtherItemsPattern = @"\A\w{1,}";

            Regex regexMeasure = new Regex(measurePattern);
            Regex regexNominal = new Regex(nominalPattern);
            Regex regexFirstPartName = new Regex(firstPartNamePattern);
            Regex regexSecondPartNameKond = new Regex(kondensatorSecondPattern);
            Regex regexFirstPartNameKond = new Regex(kondensatorFirstPartPattern);
            Regex regexLetterSecondPartNameKond = new Regex(kondensatorSecondPartLetterPattern);
            Regex regexMCXSerialPattern = new Regex(mcxSerialPattern);
            Regex regexMCXNumberPattern = new Regex(mcxNumberPattern);
            Regex regexMCXSecondNamePattern = new Regex(mcxSecondNamePattern);
            Regex regexNameOtherItemsPattern = new Regex(nameOtherItemsPattern);

                if (name.ToLower().Contains("конденсатор"))
                {

                    if (regexSecondPartNameKond.Match(name).Success)
                    {
      
                        SecondPartName = Convert.ToDouble(regexSecondPartNameKond.Match(name).Value);
                        LetterSecondPartNameKond = regexLetterSecondPartNameKond.Match(name).Value.ToLower().Trim();
                    }
                    else
                        SecondPartName = 0;
 
                    if (!regexFirstPartNameKond.Match(name).Success)
                    {
                        MessageBox.Show(name + " в файле структуры отсортирован некорректно. Обратитесь в отдел 911 для корректировки регулярного выражения.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        FirstPartName = regexFirstPartNameKond.Match(name).Value;
                    }
                }
                else
                {
                    if (regexMCXSerialPattern.Match(name).Success)
                    {
          
                        FirstPartName = regexNameOtherItemsPattern.Match(name).Value;
                        SecondPartName = Convert.ToInt32(regexMCXSerialPattern.Match(name).Value);

                        if (regexMCXNumberPattern.Match(name).Success)
                        {
                            ThirdPartName = Convert.ToInt32(regexMCXNumberPattern.Match(name).Value);
                        }
                        else
                        {
                            ThirdPartName = 0;
                        }

                        if (regexMCXSecondNamePattern.Match(name).Success)
                        {
                            SecondNameMcx = regexMCXSecondNamePattern.Match(name).Value;
                        }
                    }
                    else
                    {            
                        if (regexFirstPartName.Match(name).Success)
                        {
                            FirstPartName = regexFirstPartName.Match(name).Value;
                        }
                        else
                            FirstPartName = name;
                    }
                }
           
                if (regexNominal.Match(name).Success)
                {
              
                    try
                    {
                        string nom = regexNominal.Match(name).Value;
                        nom = nom.Trim().Replace(" ", "").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        nom = nom.Replace(".", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
                        Nominal = Convert.ToDouble(nom);
                    }
                    catch
                    {
                        MessageBox.Show("Выделить номинал " + regexNominal.Match(name).Value + " из " + name + " из элемента " + elPKI.Name + " " + elPKI.Denotation + " нельзя!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        Clipboard.SetText("Выделить номинал " + regexNominal.Match(name).Value + " из " + name + " из элемента " + elPKI.Name + " " + elPKI.Denotation + " нельзя!");
                    }
                }
                else
                {
                 
                    Nominal = 0;
                }
                switch (regexMeasure.Match(name).Value)
                {
                    case "пФ":
                        Measure = 1;
                        break;
                    case "мкФ":
                        Measure = 2;
                        break;
                    case "Ом":
                        Measure = 3;
                        break;
                    case "кОм":
                        Measure = 4;
                        break;
                    case "МОм":
                        Measure = 5;
                        break;
                    default:
                        Measure = 0;
                        break;
                }
          
        }

    }
}



// Параметры объектов Номенклатуры
public class ParametersInfo
{
    // Наименование
    public static readonly ParameterInfo Name = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
    // Обозначение
    public static readonly ParameterInfo Denotation = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);
    //Примечание
    public static readonly ParameterInfo Comment = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce"));
    // Класс объекта
    public static readonly ParameterInfo Class = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
    // ID объекта
    public static readonly ParameterInfo ObjectId = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(SystemParameterType.ObjectId);
    // Формат объекта 
    public static readonly ParameterInfo Format = new NomenclatureReference(ServerGateway.Connection).ParameterGroup.OneToOneParameters.Find(new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c"));


}


public class NomHierarchyLink
{
    public static class Fields
    {
        public static readonly Guid Pos = new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e");
        public static readonly Guid Count = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");
        public static readonly Guid Comment = new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce");
        public static readonly Guid BomSection = new Guid("7e2425f7-15ea-4921-be03-b60db93fbe28");
        public static readonly Guid Zone = new Guid("1367dda9-7850-4c15-b123-636ab692034c");
        public static readonly Guid MeasureUnit = new Guid("530922fa-8490-49c8-b93a-5e604edc1d7d");
    }
}



