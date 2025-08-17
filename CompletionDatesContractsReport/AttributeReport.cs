using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Users;

namespace CompletionDatesContractsReport
{
    class AttributeReport
    {
        public bool AllJournalReport;
        public bool DepartmentReport;
        public bool MakerReport;
        public bool SumOrder500000;
        public UserReferenceObject Department;
        public User Maker;
    }
}
