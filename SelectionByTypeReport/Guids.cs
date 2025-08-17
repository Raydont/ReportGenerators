using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SelectionByTypeReport
{
    class Guids
    {
        public class DSEKooperation
        {
            public static readonly Guid ReferenceId = new Guid("4cd39f28-27cb-4c65-b445-2711509ba163");                // справочник ДСЕ по кооперации
            public static readonly Guid GuidParameter = new Guid("265b02c1-6b9c-453a-b251-9e6005644159");              // параметр Уникальный идентификатор
            public static readonly Guid Name = new Guid("410d732c-1a99-4c36-a89b-4094610a4627");                       // параметр Наименование
        }

        public class NomencaltureNumber
        {
            public static readonly Guid ReferenceID = new Guid("c1b7c5cb-fe73-49c9-8b97-d17a0334a609");                //Справочник Номенклатурные номера
            public static readonly Guid LinkToNomenclatureObject = new Guid("3ea22b4d-3e4c-430a-91c2-ebe827a589de");   //Связь с объектом Номенклатуры
            public static readonly Guid Number = new Guid("813f7d45-0e31-4d71-a6ca-7d5eda65fe76");                     //Номенклатурный номер
        }

        public class Nomenclature
        {
            public static readonly Guid ReferenceId = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");                // Справочник Номенклатура

            public class Types
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

            public class Fields
            {
                public static readonly Guid Name = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
                public static readonly Guid Denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
                public static readonly Guid Litera = new Guid("8d297c19-5d8f-4dc0-9113-c758e43fc5ee"); // расширенная
                public static readonly Guid Code = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
                public static readonly Guid CodeRCP = new Guid("79598f4d-22e4-4cff-b062-0ca77242d8e0");
                public static readonly Guid Format = new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c");
                public static readonly Guid FirstUse = new Guid("06cb710f-e74a-43ce-944c-39e2b78f53d1");
                public static readonly Guid НеСоответствуетКлассификатору = new Guid("929af8a0-ef9b-44a4-927c-436e3fcd70d8");
                public static readonly Guid ГражданскаяПродукция = new Guid("0d5aa79f-d34c-4e8f-b908-9157fd051625");
            }
        }

        public static class NomHierarchyLink
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

        public static class ГруппыИПользователи
        {
            internal static readonly Guid RefId = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");

            //public static int Users_RKBGlobusID = 87;
            //public static int Users_RKBGlobusTopManagersID = 687;

            public static class Типы
            {
                public static readonly Guid ПроизводственнаяОрганизация = new Guid("dd7c723e-3e7f-4198-861b-97cbe727f678");

                public static readonly Guid ПроизводственноеПодразделение = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");

                public static readonly Guid Бюро = new Guid("c71cc322-7cc0-4822-a625-d72f9428a40d");
                public static readonly Guid Участок = new Guid("6f16140a-6baa-4a11-b612-597df1098094");
                public static readonly Guid Цех = new Guid("05aca34a-0e80-4e38-8367-7d7bcbf66e69");
                public static readonly Guid Лаборатория = new Guid("2c1edef0-8f15-4ffe-952d-98adeb74a4aa");
                public static readonly Guid Отдел = new Guid("1dad31c3-e89e-403f-9454-ee70fc1347c7");
                public static readonly Guid Бригада = new Guid("e6819f01-cf5f-4a91-8e3a-f56f612bcece");
                public static readonly Guid Подразделение = new Guid("9ed09a3b-8a54-403a-81a8-e9a182231da2");
                public static readonly Guid Пользователь = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");
            }

            public static class Поля
            {
                public static readonly Guid Наименование = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
                public static readonly Guid СокращенноеНаименование = new Guid("d834e736-141a-42fa-b2d8-c14f58505686");
                public static readonly Guid Номер = new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420");

                public static readonly Guid UserName = new Guid("0f3b4b68-72d6-4376-8258-f41a512cac73");
                public static readonly Guid UserSurename = new Guid("76a97c36-f2d6-49ad-abb0-f5fdc91c071b");
                public static readonly Guid UserPatronymic = new Guid("6d5a1ce6-59fc-4d1c-9b1f-67980695edc0");

                public static readonly Guid КороткоеИмя = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
            }

            public static class Связи
            {
                public static readonly Guid Руководитель = new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2");
            }
        }
    }
}
