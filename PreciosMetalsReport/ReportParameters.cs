using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PreciosMetalsReport
{
    public class ReportParameters
    {
        public List<StoreData> PKIData = new List<StoreData>(); 
    }

    public class StoreData
    {
        public string Name = string.Empty;
        public string Article = string.Empty;
        public ReferenceObject StoreObject;
        public int RemainsBeginCount = 0;
        public int ComingCount = 0;
        public int ExpenditureCount = 0;
        public int RemainsEndCount = 0;
        public string Error = string.Empty;

        public StoreData(string nameRow, string remainsBeginRow, string comingRow, string expenditureRow, string remainsEndRow)
        {
            var errorId = 0;
            try
            {
                var patternArticle = @"^(\s*)?(\d{9}|(\S*))";
                var patternCount = @"^(-)?\d*";
                errorId = 1;
                var articleMatch = Regex.Match(nameRow.Trim(), patternArticle);
                if (articleMatch.Success)
                {
                    Article = articleMatch.Value;
                }
                errorId = 2;
                var indxBeginName = nameRow.IndexOf("(");
                var indxEndName = nameRow.IndexOf(")");
                if (indxEndName >= 0 && indxBeginName >= 0)
                {
                    Name = nameRow.Substring(indxBeginName + 1, indxEndName - indxBeginName - 1);
                }
                errorId = 3;

                var remainsBeginMatch = Regex.Match(remainsBeginRow.Trim(), patternCount);
                if (remainsBeginMatch.Success)
                {
                    RemainsBeginCount = Convert.ToInt32(remainsBeginMatch.Value);
                }
                errorId = 4;
                var comingCountMatch = Regex.Match(comingRow.Trim(), patternCount);
                errorId = 5;
                if (comingCountMatch.Success)
                {
                    ComingCount = Convert.ToInt32(comingCountMatch.Value);
                }
                errorId = 6;
                var expenditureCountMatch = Regex.Match(expenditureRow.Trim(), patternCount);
                if (expenditureCountMatch.Success)
                {
                    ExpenditureCount = Convert.ToInt32(expenditureCountMatch.Value);
                }
                errorId = 7;
                var remainsEndMatch = Regex.Match(remainsEndRow.Trim(), patternCount);
                if (remainsEndMatch.Success)
                {
                    RemainsEndCount = Convert.ToInt32(remainsEndMatch.Value);
                }
                errorId = 8;
            }
            catch
            {
                Error =  comingRow + " " + errorId.ToString();
            }
        }

    }
}
