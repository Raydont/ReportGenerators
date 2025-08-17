using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TFlex.DOCs.Model.References.Files;

namespace AccountingCardReport
{
    public static class Guids
    {
        public static class Nomenclature
        {
            public static class Types
            {
                public static readonly Guid complex = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");                 // Комплекс
                public static readonly Guid complement = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");              // Комплект
            }
            public static class Fields
            {
                public static readonly Guid denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");              // Обозначение объекта номенклатуры
                public static readonly Guid format = new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c");                  // Формат
                public static readonly Guid firstUse = new Guid("06cb710f-e74a-43ce-944c-39e2b78f53d1");                // Первичная применяемость               
            }
            public static class BTDGroup                                                                                // Вкладка "Карточка учета БТД"
            {
                public static class Fields
                {
                    public static readonly Guid name = new Guid("cc6b4e09-3b45-4223-a36e-2e3f7bcd3bc2");                // Наименование (по БТД)
                    public static readonly Guid doctype = new Guid("bc0c4e69-a000-4e00-b0b1-79b0b2b94d31");             // Вип документа (бумага, электрография)
                    public static readonly Guid department = new Guid("d5d6ebff-bd57-483e-a238-540d30e3a723");          // Подразделение
                    public static readonly Guid invnum = new Guid("f255f1ae-e2b3-4b09-9439-01c662b1d3a2");              // Инвентарный номер
                    public static readonly Guid manufactoryCode = new Guid("cac5571f-a581-4514-9193-8ac2044bfab3");     // Код предприятия
                    public static readonly Guid sheetsCount = new Guid("f254a25d-7487-42ea-aae5-76fbf0bcb67c");         // Количество листов
                    public static readonly Guid incomingDate = new Guid("a8a43986-7c4d-4cca-8445-aa067c2b6423");        // Дата поступления
                }
            }
            public static class Links
            {
                public static readonly Guid baseMake = new Guid("dd112a77-29b1-4914-802a-0ece94346ea0");                // Базовое исполнение
                public static readonly Guid documentMods = new Guid("67e9ecca-8db6-4089-8897-29757c5ae46e");            // Связь с извещениями об изменениях 
                public static readonly Guid externCustomers = new Guid("a47892ba-d922-4600-9f71-ab1ca82374dd");         // Внешние абоненты - список
                public static readonly Guid internCustomers = new Guid("a16ba9c3-2a90-4c3f-9298-f83bbf7d71da");         // Внутренние абоненты - список
            }
            public static class ExternCustomer                                                                              // Внешний абонент
            {
                public static class Fields
                {
                    public static readonly Guid customer = new Guid("9884de27-dca0-4d3c-bb0c-e2e07f761814");                // Абонент
                    public static readonly Guid date = new Guid("1d60bffa-5527-4745-a913-ca754f6006a9");                    // Дата
                    public static readonly Guid count = new Guid("e7fd9236-8ab4-45b0-bc0b-16df5be0615b");                   // Кол-во экзмепляров
                    public static readonly Guid based = new Guid("d5434d37-1522-4f76-95a0-0aa46e6fc812");                   // Основание
                    public static readonly Guid cancelled = new Guid("6e519f5a-63c2-429d-a1bb-2ed40c33db5c");               // Списано - да / нет
                }
            }
            public static class InternCustomer                                                                              // Внутренний абонент
            {
                public static class Fields
                {
                    public static readonly Guid customer = new Guid("d8b3a1b2-db9f-497d-aa2f-07fd8057e457");                // Абонент
                    public static readonly Guid date = new Guid("9a86f431-ef66-457a-829d-7e9b24e1c316");                    // Дата
                    public static readonly Guid count = new Guid("b41ea7c7-e3f6-4238-a91e-ded8820b20c1");                   // Кол-во экзмепляров
                    public static readonly Guid based = new Guid("b2fca417-8fcc-4f0e-845a-b862ddfe578c");                   // Основание
                    public static readonly Guid cancelled = new Guid("98ef0e1f-1312-41da-aa1a-7e0c1bd7fe25");               // Списано - да / нет
                }
            }
        }
        public static class DocumentMods                                                                                // Извещение об изменении
        {
            public static readonly Guid id = new Guid("4853c5ce-6b8a-48ac-94bf-e80cdbcb4c1b");                          // Guid справочника
            public static class Fields
            {
                public static readonly Guid date = new Guid("fab79790-f88c-416f-bc5d-bc520eb89101");                    // Дата выпуска
                public static readonly Guid denotation = new Guid("b03c9129-7ac3-46f5-bf7d-fdd88ef1ff9a");              // Обозначение документа
            }
            public static class Links
            {
                public static readonly Guid actions = new Guid("6ff6c470-9073-4a4d-ab17-f19c2755aef0");                 // Действия - список
            }
        }
        public static class Actions                                                                                     // Действие
        {
            public static readonly Guid number = new Guid("2e9e535e-3392-4f4e-b2af-5417d9120ba9");                      // Номер изменения
            public static readonly Guid actionObject = new Guid("024b0eee-a812-4abf-a278-3804bb40bd99");                // Связь - Объект изменения
            public static readonly Guid content = new Guid("fa8ee369-2df1-4a39-a7bf-819dee3c632c");                     // Содержание изменения
        }

        public static class STO
        {
            public static class Fields
            {
                public static readonly Guid name = new Guid("01dcb8e4-7ea4-4d64-b4ee-5d2d0e015623");
                public static readonly Guid denotation = new Guid("5d3a1c0a-5fc1-4b60-94e2-e3114a43d88c");
            }
            public static class Links
            {
                public static readonly Guid documents = new Guid("520f80a4-04da-4a1a-9bdb-cf32b3fe910b");               // Связь - документы
                public static readonly Guid externCustomers = new Guid("49c7d876-4a20-486c-b4de-b0bf7cf4702b");         // Внешние абоненты - список
                public static readonly Guid internCustomers = new Guid("0d6f6bb8-cd48-4a74-9546-9ca6b655cd7c");         // Внутренние абоненты - список
            }
            public static class BTDGroup                                                                                // Вкладка "Карточка учета БТД"
            {
                public static class Fields
                {
                    public static readonly Guid department = new Guid("d70c0375-618d-4e1d-aaae-d56da54dce88");          // Подразделение
                    public static readonly Guid invnum = new Guid("30a56b98-e633-4ea6-8d77-1b189fa07b7a");              // Инвентарный номер
                    public static readonly Guid manufactoryCode = new Guid("dcfa55ea-053f-4c7e-87a8-351748abb7cf");     // Код предприятия
                    public static readonly Guid sheetsCount = new Guid("b86076d3-3c47-4a5a-855c-fb9f68bd6ac5");         // Количество листов
                    public static readonly Guid incomingDate = new Guid("b869145d-04ae-49e7-8b8f-ad34e8cacee8");        // Дата поступления
                }
            }
            public static class ExternCustomer                                                                              // Внешний абонент
            {
                public static class Fields
                {
                    public static readonly Guid customer = new Guid("62e511b4-5dbc-4285-a77c-ce6d81b918c3");                // Абонент
                    public static readonly Guid date = new Guid("273d6b66-ff13-4aa4-aa4a-5dd2932caf7d");                    // Дата
                    public static readonly Guid count = new Guid("7c97dadd-3dad-4b53-834b-74449cb5ff30");                   // Кол-во экзмепляров
                    public static readonly Guid based = new Guid("96cb8166-6132-44ff-b303-12c868c36035");                   // Основание
                    public static readonly Guid cancelled = new Guid("2828c3c6-c38f-4cd2-9f80-eb2de9b4f2d6");               // Списано - да / нет
                }
            }
            public static class InternCustomer                                                                              // Внутренний абонент
            {
                public static class Fields
                {
                    public static readonly Guid customer = new Guid("7835b5ca-3a64-4672-8049-949db0ba426c");                // Абонент
                    public static readonly Guid date = new Guid("c88484cf-6df4-4494-b27c-0b4b3f5f5433");                    // Дата
                    public static readonly Guid count = new Guid("ef6b1ef0-3c34-49fe-81aa-1aa3ed140c01");                   // Кол-во экзмепляров
                    public static readonly Guid based = new Guid("7923878b-2e71-4242-a020-24f1af2b9e11");                   // Основание
                    public static readonly Guid cancelled = new Guid("268728a6-d727-4fc3-a121-8d60469eda05");               // Списано - да / нет
                }
            }
        }

        public static class Files
        {
            public static class Fields
            {
                public static readonly Guid name = new Guid("63aa0058-4a37-4754-8973-ffbc1b88f576");
                public static readonly Guid modNumber = new Guid("270062ab-f810-4d7a-b518-6a593a812d21");
                public static readonly Guid modDenotation = new Guid("9bc9fd10-a644-406b-9602-79c856305530");
            }
        }
    }
    public static class Settings
    {
        public static readonly string PathToTemplates = @"Служебные файлы\Шаблоны отчётов\Карточка учета КД\";
        public static readonly string Form2Gost2_501_2013 = "Карточка учета Форма 2 ГОСТ 2.501-2013_Шаблон.grb";
        public static readonly string Form2aGost2_501_2013 = "Карточка учета Форма 2а ГОСТ 2.501-2013_Шаблон.grb";

        public const int nomID = 399;                                                                                       // ID справочника Номенклатура
        public const int stoID = 864;                                                                                       // ID справочника СТО
    }
    public static class RegexPatterns
    {
        public static readonly Regex PatternESKD = new Regex(@"(?:[А-Я]{4})\.(?:[0-9]{6})\.(?:[0-9]{3})(?:-(?:[0-9]{2,3}))?(?:\.(?:[0-9]{2}))?(?:(\x20)?(?:[А-Я0-9]{0,4}([.-][А-Я0-9]{0,4})?))?(?:-(?:ЛУ)|-(?:УД([0-9])?))?(?:\x20)?(?:\([А-Я0-9]{2,4}\))?");
        public static readonly Regex digits = new Regex(@"\d+");                                                            // Только цифры
        public static readonly Regex docPattern = new Regex(@"(БФМИ)\.?\s*?(\d+)\s*?-\s*?(\d+)");                           // ОБозначение извещения
        public static readonly Regex modNumber = new Regex(@"((?i)изм\s*?\.?)\s*?(\d+)?");                                  // Номер извещения
        public static readonly Regex partDoc = new Regex(@"((Часть|Приложение)\s\d{1,})?(Часть|Приложение)\s\d{1,}");       // Приложение к документу
        public static readonly Regex makeNumber = new Regex(@"(?<=-)([0-9]{2,3})$");                                        // Номер исполнения
    }
}
