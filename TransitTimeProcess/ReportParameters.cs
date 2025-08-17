using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;

namespace TransitTimeProcess
{
    public class ReportParameters
    {
        public bool SelectedObject;
        public bool ObjectsByPeriods;
        public DateTime BeginPeriod;
        public DateTime EndPeriod;
        public ReferenceObject RefObject; 
    }
}
