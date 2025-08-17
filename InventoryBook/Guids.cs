using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace InventoryBook
{
    public static class Guids
    {
        public static class Nomenclature
        {
            public static List<Guid> typesList = new List<Guid> { Types.detail, Types.assembly, Types.complex, Types.complement, 
                                                           Types.complexProgram, Types.componentProgram, Types.document };

            public static readonly Guid refInfo = new Guid("853d0f07-9632-42dd-bc7a-d91eae4b8e83");
            public static class Types
            {
                public static readonly Guid detail = new Guid("08309a17-4bee-47a5-b3c7-57a1850d55ea");
                public static readonly Guid assembly = new Guid("1cee5551-3a68-45de-9f33-2b4afdbf4a5c");
                public static readonly Guid complex = new Guid("62fe5c3f-aa31-4624-9820-7e58b8a21b74");
                public static readonly Guid complement = new Guid("3ac93dc0-8551-43a9-8f20-202d0967a340");
                public static readonly Guid complexProgram = new Guid("b7f7df88-eefa-4d73-a4dc-c08c46d584d1");
                public static readonly Guid componentProgram = new Guid("a1862b2c-032c-48af-9c9f-ab7ace0d5b2f");
                public static readonly Guid document = new Guid("7494c74f-bcb4-4ff9-9d39-21dd10058114");
            }
            public static class Fields
            {
                public static readonly Guid denotation = new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb");              // Обозначение объекта номенклатуры
                public static readonly Guid format = new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c");                  // Формат
            }
            public static class BTDGroup                                                                                // Вкладка "Карточка учета БТД"
            {
                public static class Fields
                {
                    public static readonly Guid name = new Guid("cc6b4e09-3b45-4223-a36e-2e3f7bcd3bc2");                // Наименование (по БТД)
                    public static readonly Guid department = new Guid("d5d6ebff-bd57-483e-a238-540d30e3a723");          // Подразделение
                    public static readonly Guid invnum = new Guid("f255f1ae-e2b3-4b09-9439-01c662b1d3a2");              // Инвентарный номер
                    public static readonly Guid sheetsCount = new Guid("f254a25d-7487-42ea-aae5-76fbf0bcb67c");         // Количество листов
                    public static readonly Guid incomingDate = new Guid("a8a43986-7c4d-4cca-8445-aa067c2b6423");        // Дата поступления
                    public static readonly Guid comment = new Guid("4698b7f2-8f4f-4669-baeb-3426db729cd1");             // Примечание
                }
            }
        }
        public static class STO
        {
            public static readonly Guid refInfo = new Guid("bef8bafc-4636-4c64-9f71-a28f9e3f5a9c");
            public static readonly Guid type = new Guid("4908b0a9-f6b4-4f7d-9bca-3e600aeaa36a");                         // Тип "СТО"

            public static class Fields
            {
                public static readonly Guid name = new Guid("01dcb8e4-7ea4-4d64-b4ee-5d2d0e015623");                     // Наименование
                public static readonly Guid denotation = new Guid("5d3a1c0a-5fc1-4b60-94e2-e3114a43d88c");               // Обозначение
            }
            public static class BTDGroup
            {
                public static class Fields
                {
                    public static readonly Guid department = new Guid("d70c0375-618d-4e1d-aaae-d56da54dce88");           // Подразделение
                    public static readonly Guid invnum = new Guid("30a56b98-e633-4ea6-8d77-1b189fa07b7a");               // Инвентарный номер
                    public static readonly Guid sheetsCount = new Guid("b86076d3-3c47-4a5a-855c-fb9f68bd6ac5");          // Количество листов
                    public static readonly Guid incomingDate = new Guid("b869145d-04ae-49e7-8b8f-ad34e8cacee8");         // Дата поступления
                    public static readonly Guid comment = new Guid("820bfc68-9258-44cc-96da-5e5bbb064aa7");              // Примечание
                }
            }
        }
    }
    public static class Settings
    {
        public static readonly string PathToTemplates = @"Служебные файлы\Шаблоны отчётов\Карточка учета КД\";
        public static readonly string Form1Gost2_501_2013 = "Инвентарная книга Форма 1 ГОСТ 2.501-2013_Шаблон.grb";
    }
    public static class RegexPatterns
    {
        public static readonly Regex PatternESKD = new Regex(@"(?:[А-Я]{4})\.(?:[0-9]{6})\.(?:[0-9]{3})(?:-(?:[0-9]{2,3}))?(?:\.(?:[0-9]{2}))?(?:(\x20)?(?:[А-Я0-9]{0,4}([.-][А-Я0-9]{0,4})?))?(?:-(?:ЛУ)|-(?:УД([0-9])?))?(?:\x20)?(?:\([А-Я0-9]{2,4}\))?");
        public static readonly Regex digits = new Regex(@"\d+");
    }
}
