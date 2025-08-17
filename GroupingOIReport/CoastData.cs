using System.Collections.Generic;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References.Nomenclature;

namespace GroupingOIReport
{
    public class CoastData
    {
        public List<CoastObject> Objects = new List<CoastObject>();
        public void SearchCostObject(List<NomenclatureObject> nomObjects, LogDelegate logDelegate)
        {
            var invoiceReferenceInfo = ReferenceCatalog.FindReference("Счета-фактуры");
            var invoiceReference = invoiceReferenceInfo.CreateReference();
            invoiceReference.LoadSettings.AddAllParameters();
            var invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.LinkInvoicesToContractor);
            invoiceRelation = invoiceReference.LoadSettings.AddRelation(Guids.Invoices.ListObjectDeliveredItems);
            invoiceRelation.AddAllParameters();

            foreach (var nomenclatureObject in nomObjects)
            {
                var doc = new CoastObject(nomenclatureObject, logDelegate, invoiceReference, invoiceReferenceInfo);
                Objects.Add(doc);
            }

        }
    }
}
