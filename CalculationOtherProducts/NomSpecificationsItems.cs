using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References.Nomenclature;

namespace CalculationOtherProducts
{
    public class NomSpecificationItems
    {
        public NomenclatureObject NomObject;
        public double Count;
        public List<string> Remarks = new List<string>();
        public string Number;
        public string NameLoadList = string.Empty;

        public NomSpecificationItems(NomenclatureObject nomObject, double count, List<string> remarks)
        {
            NomObject = nomObject;
            Count = count;

            foreach (var remark in remarks)
            {
                if (!Remarks.Contains(remark) && remark.Trim() != string.Empty)
                    Remarks.Add(remark);
            }
        }

        public NomSpecificationItems(NomenclatureObject nomObject, string nameLoadList,  double count)
        {
            NomObject = nomObject;
            NameLoadList = nameLoadList;
            Count = count;
        }

        public NomSpecificationItems(NomenclatureObject nomObject, double count)
        {
            NomObject = nomObject;
            Count = count;
        }
    }
}
