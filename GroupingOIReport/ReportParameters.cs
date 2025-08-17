using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References.Nomenclature;

namespace GroupingOIReport
{
    public class ReportParameters
    {
        public bool StandartItemsType = false;
        public bool OtherItems = false;
        public bool DetailsItems = false;
        public bool AddCodeDevices = false;
        public bool AddAllCodeDevice = false;
        public bool NotCodeDevices = false;
        public bool IsResultManyFiles = false;
        public bool IsCreateObjectCost = false;
        public string OrderNumber;
        public string SpecificationOrder;
        public List<NomSpecificationItems> NomObjects = new List<NomSpecificationItems>();

    }
}
