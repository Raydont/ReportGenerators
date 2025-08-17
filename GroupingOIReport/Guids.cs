using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Structure;

namespace GroupingOIReport
{
    public class Guids
    {
        public class NomenclatureParameters
        {
            public static readonly Guid NameGuid = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            public static readonly Guid DenotationGuid = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            public static readonly Guid CodeGuid = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
            public static readonly Guid LinkStoreObject = new Guid("1afd85bf-35cf-450e-93b6-82a1240110db");
        }

        public class References
        {
            public static readonly Guid NomenclatureGuid = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");            //Справочник Номенклатура и изделия
            public static readonly Guid GroupOtherAndStdItem = new Guid("914f2ef7-c8eb-4476-a385-efa037a9c203");        //Справочник группы прочих и стандартных изделий
            public static readonly Guid RegistrationCard = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");            //Справочник Регистрационно-контрольные карточки
        }

        public class NomenclatureTypes
        {
            public static readonly Guid Assembly = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");         //Сборочная единица
            public static readonly Guid Detail = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");           //Деталь
            public static readonly Guid Document = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");         //Документ
            public static readonly Guid Komplekt = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");         //Комплект
            public static readonly Guid Complex = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");          //Комплекс
            public static readonly Guid OtherItem = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");        //Прочее изделие
            public static readonly Guid StandartItem = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");     //Стандартное изделие
            public static readonly Guid Material = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");         //Материал
            public static readonly Guid ComponentProgram = new Guid("a1862b2c-032c-48af-9c9f-ab7ace0d5b2f"); //Компонент(Программы)
            public static readonly Guid ComplexProgram = new Guid("b7f7df88-eefa-4d73-a4dc-c08c46d584d1");   //Комплекс(Программы)
        }

        public class NomStore
        {
            public static readonly Guid NameObjectStore = new Guid("2adc34c8-484f-4de3-b159-1367fde71ed1");
            public static readonly Guid ArticleObjectStore = new Guid("60d321b9-3279-446f-9373-d8e9055c9d42");
            public static readonly Guid AmountObjectStore = new Guid("3d717bd5-42d0-4757-a28d-58e2d902ac42");
            public static readonly Guid TypeStoreObject = new Guid("99e1f846-18a3-4025-8f8f-e81f2b0296f1");
            public static readonly Guid TypeStoreFolderObject = new Guid("7de89103-7203-48f8-bdee-3879fcf0285a");
            public static readonly Guid StatusStore = new Guid("e09f4e6c-ee20-45d8-a87e-981bfd46172c");
        }

        public class Invoices
        {
            public static readonly Guid Number = new Guid("8e5e5101-e4a3-4461-9939-9b51510d2b35");
            public static readonly Guid Date = new Guid("ffffec91-a79c-4671-ad1f-ebcb27d8cc8b");
            public static readonly Guid ContractorName = new Guid("b4487c25-8491-4bff-97d3-4d338d9342b1");
            public static readonly Guid LinkInvoicesToContractor = new Guid("31f2de84-4864-403c-8c77-6fa3b2f02049");
            public static readonly Guid Price = new Guid("bf97aab3-37bc-4d03-a541-cd9a5246cc89");
            public static readonly Guid NDS = new Guid("507c1a7f-4355-4428-989f-b4dbf36a0c80");
            public static readonly Guid Store = new Guid("ba1a5479-0d79-45f2-b08a-ff18c9cfed2f");
            public static readonly Guid ListObjectDeliveredItems = new Guid("0002262d-8462-4355-a198-d89265d50624");
            public static readonly Guid TypeInvoices = new Guid("ace37d05-43f7-4882-85a9-d34662ff066b");
        }

        public class RegistartionCards
        {
            public static readonly Guid ТипНарядЗаказ = new Guid("b0c914bd-a920-40e8-bebd-a538c0e54436");
            public static readonly Guid ПараметрРегНомер = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
            public static readonly Guid ПараметрНомерЗаказа = new Guid("050b0e4f-ebc1-44cf-8c83-b76e21057399");
            public static readonly Guid СвязьДокументы = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5");

        }

        public class GroupOtherAndStdItemsParameters
        {
            public static readonly Guid Name = new Guid("59aced04-fb3c-4516-9662-5a49a3b87ba9");            //Наименование
            public static readonly Guid Assembly = new Guid("db0a3575-6282-4c70-b2fb-9fada8f58289");        //Сборочная единица
            public static readonly Guid Engineer = new Guid("125ea42c-181e-42a9-8c73-d88440933009");        //Инженер
        }

        public class GroupOtherAndStdItemsTypes
        {
            public static readonly Guid Item = new Guid("0a7209da-929e-4c85-b6a4-d74f734cc585");            //Изделие
            public static readonly Guid Folder = new Guid("b585a7d3-4545-461d-9479-18d8146d7958");          //Папка
        }

        public class DeliveredItems
        {
            public static readonly Guid Name = new Guid("65b23e9f-709a-4d44-8019-66fd7ea4ee53");
            public static readonly Guid Count = new Guid("8081d2e8-c352-4e66-bed5-e4a8f995f2eb");
            public static readonly Guid Price = new Guid("f9264d6b-e5e6-4555-8332-fe6634954bbf");
            public static readonly Guid LinkNomStore = new Guid("3539d7ec-5c46-4b67-893c-c9af891a341e");
            public static readonly Guid TypeDeliveredItems = new Guid("b2844fe6-2dd9-4a39-ad2f-1e1a8f22b370");
        }

        public class GroupAndUsers
        {
            public static readonly Guid РольИнженерОМТС = new Guid("617289f3-916e-47d8-a6eb-4e6b8cb89f90");
        }
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
