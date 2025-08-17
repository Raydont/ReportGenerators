using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;

namespace PlannedEconomicReport
{
    static class CommonClasses
    {
        public static void GetChilds(ReferenceObject parent, Guid guid, List<ReferenceObject> childs)
        {
            if (parent.Class.IsInherit(guid))
            {
                childs.Add(parent);
            }

            foreach (var child in parent.Children)
            {
                GetChilds(child, guid, childs);
            }
        }
    }
}
