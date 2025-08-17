using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PreciosMetalsReport
{
    public static class Guids
    {
        public static readonly Guid StandartItemsReferenceGuid = new Guid("eb5ca11e-0e19-4cbe-af10-9cac390b3d24");
        public static readonly Guid OtherItemsReferenceGuid = new Guid("5cb5931a-3329-466a-ad46-13f07c4b3d7a");
        public static readonly Guid MaterialsReferenceGuid = new Guid("5685f416-21b0-4706-9baa-d896fe0279da");
        public static readonly Guid NomenclatureReferenceGuid = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
        public static readonly Guid NomenclatureStoreGuid = new Guid("020d64f8-5ccb-41c2-ada1-e4e551c1bbf1");   //Справочник Номенклатура складов

        public static class NomenclatureType
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
            public static readonly Guid GoldContains = new Guid("e81ae161-295c-4fff-bbfc-493a3d958851");
            public static readonly Guid SilverContains = new Guid("06ed64f7-275b-4317-927c-5402890745e4");
            public static readonly Guid PlatinaContains = new Guid("6b159a33-bbd0-4b47-8a35-675cc4d3ff3a");
        }

        public static class NomenclatureParameters
        {
            public static readonly Guid NameGuid = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
            public static readonly Guid DenotationGuid = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            public static readonly Guid CodeGuid = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
        }

        public static class SpecNomParameters
        {
            public static readonly Guid NameOIGuid = new Guid("4d3a600e-0dd2-4741-b045-c37b78fe88ae");
            public static readonly Guid DenotationOIGuid = new Guid("27fe7568-2650-453e-a34c-a431ea0f5f00");
            public static readonly Guid RemarkOIGuid = new Guid("b796b9f5-1a43-4a1e-b958-f8a505b2dcf7");
            public static readonly Guid NameSIGuid = new Guid("e09101f5-8cf6-4c67-9d26-f1cd752f1631");
            public static readonly Guid DenotationSIGuid = new Guid("f4286168-e7cf-48da-8b9e-4f739f9f8d6f");
            public static readonly Guid RemarkSIGuid = new Guid("951f65fa-3ac8-4354-9d1c-20ba543bf986");
            public static readonly Guid NameMaterialsGuid = new Guid("00f6ac65-d327-4cda-967f-5cca20282bf3");
            public static readonly Guid DenotationMaterialsGuid = new Guid("3677c893-e813-4a89-a6d0-55b28c1fb18c");
            public static readonly Guid RemarkMaterialGuid = new Guid("b180c06f-4ed6-4bba-a89b-5b5bb6145a34");
        }
        public static class SpecNomRefType
        {
            public static readonly Guid StandartItemsGuid = new Guid("29f20589-0082-4bea-a221-f735ad1e2d2f");
            public static readonly Guid OtherItemsGuid = new Guid("bd33ab80-c510-45ed-83c9-db4e62117a57");
            public static readonly Guid MaterialGuid = new Guid("06cb2e2c-fc9d-497a-984e-d42da9cc733d");
        }

        public static class PeriodicTestsParameters
        {
            public static readonly Guid LinkPItoNomenclature = new Guid("3641eb13-6f84-4805-8ea9-2c4d7c76f1da");
            public static readonly Guid NamePIDevice = new Guid("d71f5d75-ef94-4148-b0de-c67991acea11");
            public static readonly Guid NumberPIDocument = new Guid("4095723b-006a-4a25-b977-4519f5379163");
            public static readonly Guid DatePIDocument = new Guid("f414cfc7-95ef-4d45-af89-3a5abc747608");
            public static readonly Guid ListObjectMaker = new Guid("d1905131-0e35-4dd1-b76e-82934fabed46");
            public static readonly Guid TypePI = new Guid("ffbb1351-5959-48bc-9000-17edda16ae68");
            public static readonly Guid NameMaker = new Guid("f7e1b0d0-9f1e-4405-bea7-eeb177d556e1");
        }
    }
}
