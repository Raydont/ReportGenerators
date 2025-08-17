using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

namespace ObjectsOIWithCoastStore
{
    public class StoreObject
    {
        public string Name;
        public string MeasureUnit;
        public int Count;
        public ReferenceObject RefObj;
        public List<NomenclatureObject> NomObjects = new List<NomenclatureObject>();
        public int Status;
        public string Article = string.Empty;
        public DateTime DateSZ = DateTime.MinValue;
        public string NumberSZ = string.Empty;


        public StoreObject(string name, string measureUnit, int count, ReferenceObject refObj)
        {
            RefObj = refObj;
            Name = name;
            MeasureUnit = measureUnit;
            Count = count;
            Article = refObj[Guids.NomStore.ArticleObjectStore].Value.ToString();
            foreach (var nomObj in refObj.GetObjects(Guids.NomenclatureParameters.LinkStoreObject))
                NomObjects.Add((NomenclatureObject)nomObj);
        }
        public StoreObject(ReferenceObject refObj)
        {
            if (refObj != null)
            {
                RefObj = refObj;
                Name = refObj[Guids.NomStore.NameObjectStore].Value.ToString();
                Article = refObj[Guids.NomStore.ArticleObjectStore].Value.ToString();
                Status = Convert.ToInt32(refObj[Guids.NomStore.StatusStore].Value);
            }
        }
    }

    public class NomObject
    {
        public string Name;
        public string Denotation;
        public string MeasureUnit;
        public string Remark;
        public double Count;
        public List<StoreObject> StoreObjects = new List<StoreObject>();
        public NomenclatureObject NomObj;

        public NomObject(NomenclatureObject nomObject)
        {
            NomObj = nomObject;
            Name = nomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString();
            Denotation = nomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString();
            // MeasureUnit = nhl[NomHierarchyLink.Fields.MeasureUnit].Value.ToString();
            //   Remark = nhl[NomHierarchyLink.Fields.Comment].Value.ToString();
            //   Count = Convert.ToDouble(nhl[NomHierarchyLink.Fields.Count].Value);
            //  if (Count != Math.Truncate(Count))
            //  {
            //     Count = CountSelectionRezistor(nomObject, Remark);
            // }
           // StoreObject
            var storeObjects = nomObject.GetObjects(Guids.NomenclatureParameters.LinkStoreObject);
            if (storeObjects != null)
            {
                foreach (var obj in storeObjects)
                {
                    StoreObjects.Add(new StoreObject(obj));
                }
            }
        }
    }

    public class Invoice
    {
        public string Number;
        public DateTime Date;
        public string ContractorInvoice;
        //  public double PriceWithNDS;
        // public double NDS;
        // public string Store;
        public ReferenceObject RefObject;
        public ReferenceObject ContractorObject = null;
        public List<PKI> pkiObjects = new List<PKI>();

        public Invoice(ReferenceObject invoice)
        {
            RefObject = invoice;
            Number = invoice[Guids.Invoices.Number].Value.ToString();
            Date = Convert.ToDateTime(invoice[Guids.Invoices.Date].Value);
            ContractorInvoice = invoice[Guids.Invoices.ContractorName].Value.ToString();

            ContractorObject = invoice.GetObject(Guids.Invoices.LinkInvoicesToContractor);

            pkiObjects.AddRange(invoice.GetObjects(Guids.Invoices.ListObjectDeliveredItems).Select(t => new PKI(t)));

          //  foreach (var pki in invoice.GetObjects(Guids.Invoices.ListObjectDeliveredItems))
         //   {
          //      pkiObjects.Add(new PKI(pki));
          //  }
        }
    }

    public class PKI
    {
        public ReferenceObject RefObject = null;
        public string Name;
        public int Count;
        public double Price;
        public ReferenceObject StoreObject = null;

        public PKI(ReferenceObject pki)
        {
            RefObject = pki;
            Name = pki[Guids.DeliveredItems.Name].Value.ToString();
            Price = Math.Round( Convert.ToDouble(pki[Guids.DeliveredItems.Price].Value)/ Convert.ToDouble(pki[Guids.DeliveredItems.Count].Value),2);
            StoreObject = pki.GetObject(Guids.DeliveredItems.LinkNomStore);
        }

    }
}
