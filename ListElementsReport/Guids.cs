using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListElementsReport
{
    public class Guids
    {
        public class SpecReferences
        {
            public static readonly Guid StandartItems = new Guid("eb5ca11e-0e19-4cbe-af10-9cac390b3d24");
            public static readonly Guid OtherItems = new Guid("5cb5931a-3329-466a-ad46-13f07c4b3d7a");
            public static readonly Guid Materials = new Guid("5685f416-21b0-4706-9baa-d896fe0279da");

        }

        public class SpecReferencesTypes
        {
            public static readonly Guid StandartItems = new Guid("29f20589-0082-4bea-a221-f735ad1e2d2f");
            public static readonly Guid OtherItems = new Guid("bd33ab80-c510-45ed-83c9-db4e62117a57");
            public static readonly Guid Materials = new Guid("06cb2e2c-fc9d-497a-984e-d42da9cc733d");

        }

        public class OtherItemsParameters
        {
            public static readonly Guid Name = new Guid("4d3a600e-0dd2-4741-b045-c37b78fe88ae");
            public static readonly Guid Denotation = new Guid("27fe7568-2650-453e-a34c-a431ea0f5f00");
            public static readonly Guid НаименованиеЭталонное = new Guid("ce35c292-0d7d-4338-8cb8-4dffd1b0aec5");
            public static readonly Guid ОбозначениеЭталонное = new Guid("2bb27851-006a-406a-9437-35434f82ade8");
            public static readonly Guid Нормализован = new Guid("9183734f-4c23-4666-b315-30c33b9d3665");
        }

        public class StandartItemsParameters
        {
            public static readonly Guid Name = new Guid("e09101f5-8cf6-4c67-9d26-f1cd752f1631");
            public static readonly Guid Denotation = new Guid("f4286168-e7cf-48da-8b9e-4f739f9f8d6f");
            public static readonly Guid НаименованиеЭталонное = new Guid("758ee145-4b49-4719-934d-50048a1bf8a1");
            public static readonly Guid ОбозначениеЭталонное = new Guid("db9f1944-d1df-42a7-b738-79ea4c1e1827");
            public static readonly Guid Нормализован = new Guid("2e7ede67-b9e3-443e-99db-9e98b732881d");
        }

        public class MaterialsParameters
        {
            public static readonly Guid Name = new Guid("00f6ac65-d327-4cda-967f-5cca20282bf3");
            public static readonly Guid Denotation = new Guid("3677c893-e813-4a89-a6d0-55b28c1fb18c");
        }

        public static class НоменклатураИИзделия
        {
            public static readonly Guid Id = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static class Параметры
            {
                public static readonly Guid Обозначение = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");
                public static readonly Guid ОтображатьВДокументеСтруктурыЗаказа = new Guid("4dcaa944-1aed-47b3-af1e-1b69798a5cfc");
                public static readonly Guid Наименование = new Guid("45e0d244-55f3-4091-869c-fcf0bb643765");
                public static readonly Guid Назначение = new Guid("36e92fff-79d5-4c44-b121-8aa17eafd36f");
                public static readonly Guid Формат = new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c");
                public static readonly Guid Литера = new Guid("8d297c19-5d8f-4dc0-9113-c758e43fc5ee");
                public static readonly Guid Шифр = new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6");
                public static readonly Guid Подразделение = new Guid("d5d6ebff-bd57-483e-a238-540d30e3a723");
            }
            public static class Связи
            {
                public static readonly Guid ИзвещенияОбИзменениях = new Guid("67e9ecca-8db6-4089-8897-29757c5ae46e");
            }
            public static class Типы
            {
                public static readonly Guid Комплекс = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");
                public static readonly Guid СборочнаяЕдиница = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
                public static readonly Guid Комплект = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");
                public static readonly Guid ПрочееИзделие = new Guid("f50df957-b532-480f-8777-f5cb00d541b5");
                public static readonly Guid СтандартноеИзделие = new Guid("87078af0-d5a1-433a-afba-0aaeab7271b5");
                public static readonly Guid Материал = new Guid("b91daab4-88e0-4f3e-a332-cec9d5972987");
                public static readonly Guid Документ = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");
                public static readonly Guid Деталь = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
                public static readonly Guid ТехнологическоеОборудование = new Guid("19cac458-770d-421b-9a81-7b392f92ad51");
                public static readonly Guid УказаниеПоДоработке = new Guid("9e5ca733-cbe0-4aa9-ad1e-fb8a7afb82be");
            }
            public static class ПараметрыИерархии
            {
                public static readonly Guid Количество = new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818");
                public static readonly Guid Примечание = new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce");
            }
        }
    }
}
