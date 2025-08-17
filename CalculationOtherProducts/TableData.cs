using System;
using System.Text.RegularExpressions;

namespace CalculationOtherProducts
{
    public class TableData
    {
        public string Name;
        public string Denotation;
        public string ShortName;
        public string MeasureUnit;
        public double Count;
        public string NameTextNorm;
        public bool FindedObj = false;
        public string RegularExpression;
        public string Section;

        public TableData(string name, string denotation, string measureUnit, double count, string nameTextNorm, string regExpr, string remark)
        {
            Name = name;
            Denotation = denotation;
            MeasureUnit = measureUnit;
            if (count != Math.Truncate(count))
                Count = CountSelectionRezistor(remark);
            else
                Count = count;
            NameTextNorm = nameTextNorm;
            RegularExpression = regExpr;
        }

        private int CountSelectionRezistor(string remark)
        {
            var remarkWithOutLetter = remark.Replace("*", "");
            Regex regexDigitalOne = new Regex(@"(C|С|R|W)\d{1,}");
            Regex regexDigital = new Regex(@"(?<=(C|С|R|W))\d{1,}");
            Regex regexDigitalTwo = new Regex(@"(C|С|R|W)\d{1,}(-|(\s{1,})?\.{3}(\s{1,})?)(C|С|R|W)\d{1,}");
            int countRezist = 0;
            var matchrez = string.Empty;
            try
            {
                foreach (Match match in regexDigitalTwo.Matches(remarkWithOutLetter))
                {
                    matchrez = regexDigital.Matches(match.Value.ToString())[0].Value + " " + regexDigital.Matches(match.Value.ToString())[1].Value;
                    countRezist += Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[1].Value) - Convert.ToInt32(regexDigital.Matches(match.Value.ToString())[0].Value) + 1;
                    remarkWithOutLetter = remarkWithOutLetter.Replace(match.Value.ToString(), "");
                }
                countRezist += regexDigitalOne.Matches(remarkWithOutLetter).Count;
            }
            catch
            {
                return 0;
               // System.Windows.Forms.MessageBox.Show("7. Не могу посчитать количество резистора\r\n" + nomObj + "\r\nОшибка в примечании: " + remark + "\r\nКод ошибки " + errorId + "\r\n" + matchrez, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return countRezist;
        }


    }
}