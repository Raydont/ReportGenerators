using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;

namespace GroupingOIReport
{
    public class StoreObject
    {
        public string Name;
        public ReferenceObject RefObj;
        public Invoice InvoiceObject;

        public StoreObject(ReferenceObject refObj, Reference invoiceReference, ReferenceInfo invoiceReferenceInfo)
        {
            if (refObj != null)
            {
                RefObj = refObj;
                Name = refObj[Guids.NomStore.NameObjectStore].Value.ToString();
                InvoiceObject = SearchInvoice(invoiceReference, invoiceReferenceInfo);
            }
        }

        private Invoice SearchInvoice(Reference invoiceReference, ReferenceInfo invoiceReferenceInfo)
        {
            var invoices = new List<Invoice>();
            var filter = new Filter(invoiceReferenceInfo);
            filter.Terms.AddTerm("[Поставленные изделия].[ПКИ]->[ID]", ComparisonOperator.Equal, RefObj.SystemFields.Id);
            var refObjects = invoiceReference.Find(filter).ToList();
            invoices.AddRange(refObjects.Select(t => new Invoice(t)).ToList());
            return invoices.Count != 0 ? invoices.OrderByDescending(t => t.Date).FirstOrDefault() : null;
        }
    }

    public class CoastObject
    {
        public List<StoreObject> StoreObjects = new List<StoreObject>();
        public NomenclatureObject NomObj;

        public CoastObject(NomenclatureObject nomObject, LogDelegate logDelegate, Reference invoiceReference, ReferenceInfo invoiceReferenceInfo)
        {
            NomObj = nomObject;
            List<ReferenceObject> storeObjects = new List<ReferenceObject>();

            try
            {
                storeObjects.AddRange(nomObject.GetObjects(Guids.NomenclatureParameters.LinkStoreObject));
                if (storeObjects != null && storeObjects.Count > 0)
                {
                    logDelegate(String.Format("Получение цены объекта: " + nomObject));
                    foreach (var obj in storeObjects)
                    {
                        var storeObject = new StoreObject(obj, invoiceReference, invoiceReferenceInfo);
                        StoreObjects.Add(storeObject);
                    }
                }
            }
            catch
            {
            }
        }
    }
    
    public class Invoice
    {
        public string Number;
        public DateTime Date;
        public string ContractorInvoice;
        public ReferenceObject RefObject;
        public List<PKI> pkiObjects = new List<PKI>();

        public Invoice(ReferenceObject invoice)
        {
            RefObject = invoice;
            Number = invoice[Guids.Invoices.Number].Value.ToString();
            Date = Convert.ToDateTime(invoice[Guids.Invoices.Date].Value);
            var contractorObject = invoice.GetObject(Guids.Invoices.LinkInvoicesToContractor);
            ContractorInvoice = contractorObject != null ? contractorObject.ToString() : invoice[Guids.Invoices.ContractorName].Value.ToString();           
            pkiObjects.AddRange(invoice.GetObjects(Guids.Invoices.ListObjectDeliveredItems).Select(t => new PKI(t)));
        }
    }

    public class PKI
    {
        public double Price;
        public ReferenceObject StoreObject = null;

        public PKI(ReferenceObject pki)
        {
            Price = Math.Round(Convert.ToDouble(pki[Guids.DeliveredItems.Price].Value) / Convert.ToDouble(pki[Guids.DeliveredItems.Count].Value), 2);
            StoreObject = pki.GetObject(Guids.DeliveredItems.LinkNomStore);
        }
    }
}
