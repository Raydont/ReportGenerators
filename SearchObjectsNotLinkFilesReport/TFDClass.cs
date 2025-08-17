using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SearchObjectsNotLinkFilesReport
{
    public static class TFDClass
    {
        public static readonly string ComplementName = "Комплекты";
        public static readonly string DetailName = "Детали";
        public static readonly string DocumentName = "Документы";

        public static string InsertSection(string typeName)
        {
            string sectionName = string.Empty;
            if (typeName == "Комплект")
                sectionName = ComplementName;
            if (typeName == "Деталь")
                sectionName = DetailName;
            if (typeName == "Документ")
                sectionName = DocumentName;
            

            return sectionName;
        }
   }
}
