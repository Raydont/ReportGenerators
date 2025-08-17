using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Nomenclature;

namespace GroupSpecVarB
{
    public class TFDDocument
    {
        // Порядок полей в запросе чтения документа

        public int ObjectID;
        public int ParentID;
        public int Class;
        public string FirstUse;
        public double Amount;
        public string Zone;
        public int Position;
        public string Format;
        public string Naimenovanie;
        public string Denotation;
        public string Remarks;
        public string BomSection = string.Empty;
        public string MeasureUnit;
        public TFDDocument rootDocument;
        public int ActualVersion;
        public int ClientViewID;
        //private List<TFDDocument> _childDocuments;

        public TFDDocument()
        {
        }

        public TFDDocument(NomenclatureObject nomenclatureObject)
        {
            
            ObjectID = Convert.ToInt32(nomenclatureObject[TFDGuids.ParInfo.ObjectId].Value);
            Class = Convert.ToInt32(nomenclatureObject[TFDGuids.ParInfo.Class].Value);
            ParentID = Convert.ToInt32(nomenclatureObject.Parent[TFDGuids.ParInfo.ObjectId]);
            Naimenovanie = nomenclatureObject[TFDGuids.ParInfo.Name].Value.ToString();
            Denotation = nomenclatureObject[TFDGuids.ParInfo.Denotation].Value.ToString();
            Format = nomenclatureObject[TFDGuids.ParInfo.Format].Value.ToString();



        }

        public TFDDocument(int documentID)
        {
            NomenclatureReference refNomenclature = new NomenclatureReference();
            // TFlex.DOCs.Model.Structure.ParameterInfo classInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
            TFlex.DOCs.Model.Structure.ParameterInfo nameInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
            TFlex.DOCs.Model.Structure.ParameterInfo denotationInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);
            //TFlex.DOCs.Model.Structure.ParameterInfo bomSectionInfo = refNomenclature.ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.);

            rootDocument = null;
            ParentID = 0;
            ObjectID = documentID;
            NomenclatureObject nomenclatureObject = (NomenclatureObject)refNomenclature.Find(documentID);
            // Class = nomenclatureObject[classInfo].Value.ToString();
            Amount = 1;
            Naimenovanie = (string)nomenclatureObject[nameInfo].Value;
            Denotation = (string)nomenclatureObject[denotationInfo].Value;
            // Вырезаем из обозначения последние два символа (ПИ) по требованию Баскаковой
            string pi = Denotation.Trim().Remove(Denotation.Length - 2, 2);

            if (Denotation.Contains("ПИ") && !pi.Contains("ПИ"))
            {
                Denotation = Denotation.Trim().Remove(Denotation.Length - 2, 2);
            }


            if (Denotation.Contains("ПИ") && !(Denotation.Trim().Remove(Denotation.Length - 5, 2)).Contains("ПИ"))
            {
                Denotation = Denotation.Trim().Remove(Denotation.Length - 5, 2);

            }
            if (Denotation.Contains("ПИ") && !(Denotation.Trim().Remove(Denotation.Length - 6, 2)).Contains("ПИ"))
            {
                Denotation = Denotation.Trim().Remove(Denotation.Length - 6, 2);

            }

            Remarks = String.Empty;
        }
    }
}
