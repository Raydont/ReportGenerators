using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CalculationOtherProducts
{

    public class FirstSection
    {
       // public List<string> OINames = new List<string>();
        public List<string> OINamesTextNorm = new List<string>();
        public string Name;
        public List<OtherItem> OtherItems = new List<OtherItem>();
        public List<SecondSection> SecondSection = new List<SecondSection>();
    }

    public class SecondSection
    {
        public string Name;
        public List<string> OINamesTextNorm = new List<string>();
        public List<OtherItem> OtherItems = new List<OtherItem>();
    }

    public class StructureTableData
    {
        public string NameSections;
        public List<TableData> Data = new List<TableData>();
    }

    public class OtherItem
    {
        public string Name;
        public string Denotation;
        public string NameTextNorm;
        public string SingularName;
        public string RegularExpression;

        public OtherItem(string name, string denotation, string nameTextNorm, string singularName, string regularExpression)
        {
            SingularName = singularName;
            Name = name;
            Denotation = denotation;
            RegularExpression = regularExpression;
            NameTextNorm = nameTextNorm;
        }
    }
}
