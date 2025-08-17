using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;

namespace TermsPeriodicTestsReport
{
    public class PTDocument
    {
        public string Name;                                      //Номер документа ПИ
        public DateTime Date;                                    //Дата документа ПИ
        public string LastNotice;                                //Последнее извещение
        public ReferenceObject NomenclatureObject;              //Связанный номенклатурный объект
        public List<string> Executers = new List<string>();      //Список исполнителей
        public DateTime TermFollowingTest;                       //Срок следущего периодического испытания
        public int Count;
        public bool ObjectNotFound = false;

        public PTDocument(ReferenceObject refObject, ReferenceObject nomObject)
        {
            int errCode = 0;

            try
            {
                Name = refObject[Guids.PeriodicTestsDocuments.NamePI].Value.ToString();
                errCode++;
                Date = Convert.ToDateTime(refObject[Guids.PeriodicTestsDocuments.DatePI].Value);
                errCode++;
                LastNotice = refObject[Guids.PeriodicTestsDocuments.LastNotice].Value.ToString();
                errCode++;
                NomenclatureObject = nomObject;
                errCode++;
                Executers.AddRange(refObject.GetObjects(Guids.PeriodicTestsDocuments.ListObjectsExecuters).ToList().Select(t => t.ToString()).ToList()); //??????
                errCode++;
                TermFollowingTest = Convert.ToDateTime(refObject[Guids.PeriodicTestsDocuments.TermFollowingTest].Value.ToString());
                ObjectNotFound = false;
            }
            catch
            {
               // MessageBox.Show("Код ошибки " + errCode);
            }
        }

        public PTDocument(ReferenceObject refObject)
        {
            NomenclatureObject = refObject;
            ObjectNotFound = true; 
        }
    }
}
