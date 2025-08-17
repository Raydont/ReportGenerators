using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;

namespace ObjectUsageCodeDevicesReport
{
    public class ReportParameters
    {
        public List<ReferenceObject> PKIObjects = new List<ReferenceObject>();
        public bool FullInformation = false;
        public bool OnlyCodeDevicesNames = false;
    }
}
