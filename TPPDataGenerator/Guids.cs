using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TPPDataGenerator
{
    public static class Guids
    {
        public static class Номенклатура
        {
            public static Guid id = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static class Типы
            {
                public static Guid комплекс = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");
                public static Guid комплект = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");
                public static Guid сборочнаяЕдиница = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
                public static Guid деталь = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
                public static Guid документ = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");
                public static Guid указаниеДоработка = new Guid("9e5ca733-cbe0-4aa9-ad1e-fb8a7afb82be");
            }
            public static class Поля
            {
                public static Guid тпНеактуален = new Guid("92e9c178-a0b9-4ec4-9a64-84f4a654e0c1");
                public static Guid наименование = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
                public static Guid обозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
            }
            public static class Связи
            {
                public static Guid техПроцессы = new Guid("e1e8fa07-6598-444d-8f57-3cfd1a3f4360");
            }
        }
        public static class ТехПроцессы
        {
            public static Guid id = new Guid("353a49ac-569e-477c-8d65-e2c49e25dfeb");
            public static class Поля
            {
                public static Guid актуальностьТП = new Guid("37924b0f-631e-494b-87b5-d2bec270d001");
            }
        }
        public static class Извещения
        {
            public static Guid id = new Guid("4853c5ce-6b8a-48ac-94bf-e80cdbcb4c1b");
            public static class Поля
            {
                public static Guid обозначение = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");
            }
            public static class Связи
            {
                public static Guid изменяемыеОбъекты = new Guid("67e9ecca-8db6-4089-8897-29757c5ae46e");
            }
        }
        public static class СтруктурыЗаказов
        {
            public static Guid id = new Guid("f61d0185-8db1-4154-b03d-5ba17b6dc62f");

            public static class Типы
            {
                public static Guid структура = new Guid("a21911bf-7137-4a48-a4fd-71d70ab57fcc");
                public static Guid прочееИзделие = new Guid("c1e4f9a3-7fe4-4908-8e68-9c1c9afb571f");
                public static Guid стандартноеИзделие = new Guid("c622a19e-3f8f-4a4c-8c00-9b9c06e190fb");
                public static Guid дсеДляРазработки = new Guid("f3ec739d-7427-4a93-bfd5-ceea59a241dd");
                public static Guid вариантЭСИ = new Guid("5593a390-67b2-450e-9331-fb567d8f5c85");
                public static Guid описание = new Guid("5e8a90cb-4a2c-41d9-a5f3-479c800683f0");
                public static Guid документ = new Guid("0e00f8b3-bf26-43f5-ba61-d2eabf84b678");
                public static Guid указаниеДоработка = new Guid("e7979f37-0240-4dc7-88b7-9f5348c4f8e9");
            }
            public static class Поля
            {
                public static Guid номерЗаказа = new Guid("0e0e527e-058b-4364-ad7c-1c35cee6b1c4");
                public static Guid обозначение = new Guid("ad9da063-9be6-46e9-b01a-088e9dbc6530");
                public static Guid тпАктуальны = new Guid("fb0c8d0d-23d3-4419-9e4a-e4bc25b6803e");
                public static Guid выгрузитьв1С = new Guid("4ef8c660-167b-4eaa-ad15-553396588145");
                public static Guid кдИзменена = new Guid("ca3ff76d-0b7b-46be-9692-c2543d79d7d1");
            }
            public static class Связи
            {
                public static Guid номенклатура = new Guid("622e1301-c16d-4714-a479-d82f6a7da5c7");
                public static Guid номерЗаказа = new Guid("02168ac1-92c6-47a0-86e8-00c5f893e6b9");
            }
        }
        public static class ЛогВыгрузки
        {
            public static Guid id = new Guid("a705b9df-4e8e-4702-aced-bf3d54c9101e");
            public static class Типы
            {
                public static Guid лог = new Guid("7ae7f700-62b6-4bda-ae57-f0cab5770f52");
            }
            public static class Поля
            {
                public static Guid инициатор = new Guid("d6209c37-6f6d-438c-97d2-5fbb464c48e5");
                public static Guid флагВыгрузитьв1С = new Guid("6716a3ef-3185-4d16-aab0-9b170af56224");
            }
        }
        public static class НомераЗаказов
        {
            public static Guid id = new Guid("1c2d5a9d-f1dc-4c38-bbb8-939a974f63b5");
            public static class Поля
            {
                public static Guid номерЗаказа = new Guid("ae85bf35-5941-40d8-a96f-4ba5368cbb43");
                public static Guid заказЗакрыт = new Guid("65110d10-aca7-404e-9cd5-a66918cdfc9c");
            }
        }
        public static class Стадии
        {
            public static class ТП
            {
                public static Guid тпУтвержден = new Guid("e7688986-4488-4858-bab4-b80979f16063");
                public static Guid тпОткрыт = new Guid("52471213-2818-4a54-a590-55e2d5b6b0fc");
            }
            public static class КД
            {
                public static Guid установкаТПАктуальны = new Guid("060e1f8a-52bb-4c76-9b63-c79120ec4af7");
            }
            public static class ЭСЗ
            {
                public static Guid установкаТПАктуальны = new Guid("b231674a-c7be-4147-8750-0141f9b89b06");
            }
        }
        public static class ГруппыПользователи
        {
            public static Guid id = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
            public static class Поля
            {
                public static Guid короткоеИмя = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
            }
            public static class Объекты
            { 
                public static Guid отдел60 = new Guid("afc508b5-c5cf-42ba-b04c-5c238e30a62a");
                public static Guid начальник60 = new Guid("b7404e9f-17dc-4871-9700-46a63402cfb5");
                public static Guid бюро603 = new Guid("fbbc339b-2085-477c-938e-ae3ee101b32e");
                public static Guid начальник603 = new Guid("2d4d8989-bd08-41a5-987b-e3f2dea97fe4");
                public static Guid отдел911 = new Guid("94c47ae8-83e4-486c-a9dc-661e5ae3fd92");
                public static Guid начальник911 = new Guid("d0cfa8fa-9384-4ddf-8a55-c9abd8cb73f6");
            }
        }
        public static class СтопЛистДляТПП
        {
            public static Guid id = new Guid("40b818d1-8a23-4957-96e9-07c93710676b");
            public static class Поля
            {
                public static Guid наименование = new Guid("59843f4d-db7f-422e-9320-de7dc165a711");
                public static Guid обозначение = new Guid("df411c6f-3e3c-40e1-a5bd-45a38a7b8d74");
            }
        }
        public static class ЛогЗапусков
        {
            public static Guid id = new Guid("df7be9e2-c76a-4c72-954b-3958b452455a");
            public static class Типы
            {
                public static Guid лог = new Guid("f3d20df1-4312-44a7-9d97-d4cee94eedf7");
            }
            public static class Поля
            {
                public static Guid обозначение = new Guid("902862de-56f8-4ac3-9d28-401bb87e2019");
                public static Guid этоИзвещение = new Guid("22e69bfe-33f7-4a7f-a5aa-70c205184e2b");
            }
        }
        public static class ЭСИсУтвержденнымиТП
        {
            public static Guid id = new Guid("e44bf13e-8134-43c4-9093-2f22d787eb26");
            public static class Поля
            {
                public static Guid наименование = new Guid("c839e9f5-93f4-4a1e-bb8d-d6663998946c");
                public static Guid обозначение = new Guid("ee98df2b-1c56-4cce-a3ce-6f3d92624ad0");
                public static Guid ID_номенклатуры = new Guid("25c593d2-961d-430b-a8f9-0616178bdc24");
                public static Guid учтен = new Guid("ac016db5-6a88-4328-93af-605fccef1015");
            }
        }
        public static Guid строкаПодключения = new Guid("f0d42460-398e-46fa-840a-10bfb8f3bd83");
    }
}
