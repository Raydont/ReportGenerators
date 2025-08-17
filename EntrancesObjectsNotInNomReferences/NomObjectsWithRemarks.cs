using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References.Nomenclature;

namespace EntrancesObjectsNotInNomReferences
{
    public class NomObjectsWithRemarks
    {
        public NomenclatureObject NomObject;
        public string Remark = string.Empty;
        public bool Write = false;

        public NomObjectsWithRemarks(NomenclatureObject nomObject, string remark)
        {
            NomObject = nomObject;
            Remark = remark;
            Write = false;
        }
    }

    public class OrderStructure
    {
        public NomenclatureObject Order;
        public List<DSEWithStdItems> ListDSE = new List<DSEWithStdItems>();
        public List<StdItemsWithDSE> ListSTD = new List<StdItemsWithDSE>();

        public OrderStructure(NomenclatureObject order, List<DSEWithStdItems> listDSE, bool bySTD)
        {
            Order = order;
            if (bySTD)
            {
                foreach (var dse in listDSE)
                {
                    foreach (var std in dse.StandartItems)
                    {
                        if (!std.Write)
                        {
                            if (ListSTD.Count(t => t.StandartItem.NomObject.SystemFields.Id == std.NomObject.SystemFields.Id) == 0)
                            {
                                StdItemsWithDSE stdwithDSE = new StdItemsWithDSE(dse.DSE, std);
                                foreach (var dse2 in listDSE)
                                {
                                    foreach (var std2 in dse.StandartItems)
                                    {
                                        if (std.NomObject.SystemFields.Id == std2.NomObject.SystemFields.Id &&
                                            dse != dse2 && dse2.StandartItems.Contains(std))
                                        {
                                            stdwithDSE.DSE.Add(dse2.DSE);
                                            //  std2.Write = true;
                                        }
                                    }
                                }
                                stdwithDSE.DistinctDSE();
                                ListSTD.Add(stdwithDSE);
                            }
                        }
                    }
                }
            }
            else
            {
               
                ListDSE.AddRange(listDSE);
            }
        }

       


    }

    public class DSEWithStdItems
    {
        public NomenclatureObject DSE;
        public List<NomObjectsWithRemarks> StandartItems = new List<NomObjectsWithRemarks>(); 

        public DSEWithStdItems(NomenclatureObject dse, NomObjectsWithRemarks stdItem)
        {
            DSE = dse;
            StandartItems.Add(stdItem);
        }
    }

    public class StdItemsWithDSE
    {
        public List<NomenclatureObject> DSE = new List<NomenclatureObject>();
        public NomObjectsWithRemarks StandartItem;

        public StdItemsWithDSE(NomenclatureObject dse, NomObjectsWithRemarks stdItem)
        {
            DSE.Add(dse);
            StandartItem = stdItem;
        }

        public void DistinctDSE()
        {

            for (int mainID = 0; mainID < DSE.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < DSE.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (DSE[mainID].SystemFields.Id == DSE[slaveID].SystemFields.Id)
                    {
                        DSE.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
        }
    }
}
