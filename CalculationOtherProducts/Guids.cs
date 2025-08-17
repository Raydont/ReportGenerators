using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;

namespace CalculationOtherProducts
{
    public class Guids
    {
        public class NomenclatureParameters
        {
            public static readonly Guid Name = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            public static readonly Guid Denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");

        }

        public class NomenclatureTypes
        {
            public static readonly Guid OtherItems = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
            public static readonly Guid StandartItems = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
            public static readonly Guid Materials = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");
            public static readonly Guid Details = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
            public static readonly Guid Assembly = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
            public static readonly Guid Complex = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");
            public static readonly Guid Komplekt = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");
        }

        public class CategoryOtherItems
        {
            public static readonly Guid Name = new Guid("59aced04-fb3c-4516-9662-5a49a3b87ba9");
            public static readonly Guid LinkOtherItems = new Guid("5d02eeab-e8fb-4958-973d-6b3ef2038605");    //Связь со справочником прочие изделия

            public static readonly Guid FolderTypes = new Guid("b585a7d3-4545-461d-9479-18d8146d7958");       // тип папки
            public static readonly Guid ObjectTypes = new Guid("0a7209da-929e-4c85-b6a4-d74f734cc585");       // Тип объекты

            public static readonly Guid SerialIndex = new Guid("cddfb839-edf6-46a7-bbee-edf1bdb1ac4c");       // Порядковый номер
            public static readonly Guid RegularExpression = new Guid("d45903ef-f8c8-40dc-8f79-eb355b26eb2e"); // Регулярное выражение
            public static readonly Guid OtherItemsNames = new Guid("7f4c936e-dba2-46a5-8fce-3b4a01dc6937");  // Наименования вхождений прочих изделий
        }

        public class OtherItems
        {
            public static readonly Guid Name = new Guid("4d3a600e-0dd2-4741-b045-c37b78fe88ae");
            public static readonly Guid Denotation = new Guid("27fe7568-2650-453e-a34c-a431ea0f5f00");
            public static readonly Guid Remark = new Guid("b796b9f5-1a43-4a1e-b958-f8a505b2dcf7");

            public static readonly Guid FolderTypes = new Guid("04cd41fb-02e3-44c5-853f-691dcc2404c9");
            public static readonly Guid ObjectTypes = new Guid("bd33ab80-c510-45ed-83c9-db4e62117a57");
        }

        public class References
        {
            public static readonly Guid Nomenclature = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static readonly Guid CategoriesOtherItems = new Guid("914f2ef7-c8eb-4476-a385-efa037a9c203");

        }
        public readonly static Guid GAVGuid = new Guid("712890e3-558b-4429-a24f-16dbca4c6a20");


    }

    public class ParametersInfo
    {
        // Наименование
        public static readonly ParameterInfo Name = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Name);
        // Обозначение
        public static readonly ParameterInfo Denotation = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(NomenclatureReferenceObject.FieldKeys.Denotation);
        //Примечание
        public static readonly ParameterInfo Comment = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce"));
        // Класс объекта
        public static readonly ParameterInfo Class = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(SystemParameterType.Class);
        // ID объекта
        public static readonly ParameterInfo ObjectId = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(SystemParameterType.ObjectId);
        // Формат объекта 
        public static readonly ParameterInfo Format = new NomenclatureReference().ParameterGroup.OneToOneParameters.Find(new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c"));
    }


    public class NomHierarchyLink
    {
        public static class Fields
        {
            public static readonly Guid Pos = new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e");
            public static readonly Guid Count = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");
            public static readonly Guid Comment = new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce");
            public static readonly Guid BomSection = new Guid("7e2425f7-15ea-4921-be03-b60db93fbe28");
            public static readonly Guid Zone = new Guid("1367dda9-7850-4c15-b123-636ab692034c");
            public static readonly Guid MeasureUnit = new Guid("530922fa-8490-49c8-b93a-5e604edc1d7d");
        }
    }
}
