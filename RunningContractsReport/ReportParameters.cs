using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RunningContractsReport
{
    public class ReportParameters
    {
        public bool DOD = true;
        public DateTime BeginDate;
        public DateTime EndDate;
        public bool ActualContract = true;
    }

    public class ReportItem
    {
        public string NameContarctor;
        public string Number;
        public DateTime Date;
        public double Cost;
        public double CostNDS;
        public string Currency;
        public double AdvancePayment;
        public string Content;
        public bool ActualContract;
        public List<DateTime> ListTerm = new List<DateTime>();
        public string NumberOrder;
        public List<string> ListNumberOrder = new List<string>();
        public string RealNumber;
        public DateTime RealDate;

        //public string GetLines(string text, int countSymbolsOnLine)
        //{
        //    var textWords = text.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        //    string res = string.Empty;
        //    var lines = new List<string>();
        //    var line = "";
        //    foreach (string word in textWords)
        //    {
        //        if (line.Length + word.Length > countSymbolsOnLine)
        //        {
        //            if (line.Length == 0)
        //            {
        //                lines.Add(word);

        //            }
        //            else
        //            {
        //                lines.Add(line.Trim());
        //                line = word;
        //            }
        //        }
        //        else
        //        {
        //            line += " " + word;
        //        }
        //    }
        //    if (!string.IsNullOrEmpty(line.Trim()))
        //    {
        //        lines.Add(line.Trim());
        //    }


        //    foreach (var str in lines.Where(t=>t.Trim() != string.Empty).Select(t=>t.Trim()))
        //    {
        //        if (str != lines[0])
        //            res += "\n" + str;
        //        else
        //            res = str;
        //    }
        //    return res;
        //}
    }
}
