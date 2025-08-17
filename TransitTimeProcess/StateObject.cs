using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;

namespace TransitTimeProcess
{
    public class StateObject
    {
        public string Name;
        public DateTime BeginDate;
        public DateTime EndDate;
        public TimeSpan Term;
        public ReferenceObject RefObject;
    }
}
