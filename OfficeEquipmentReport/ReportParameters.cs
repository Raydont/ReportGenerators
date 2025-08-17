using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;

namespace OfficeEquipmentReport
{
    public class ReportParameters
    {
        public ReferenceObject Department;
        public bool Computer = false;
        public bool Monitor = false;
        public bool MFU = false;
        public bool Printer = false;
        public bool Scanner = false;
        public bool OtherDevices = false;
        public bool Storage = false;
        public bool Notebook = false;
    }
}
