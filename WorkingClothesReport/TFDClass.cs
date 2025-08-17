using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WorkingClothesReport
{
    class TFDClass
    {
        // Guid справочника "Группы и пользователи"
        public static readonly Guid usersGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        // Guid типа "Производственное подразделение"
        public static readonly Guid manufactoryDepartment = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");

        // Guid справочника "Закупки"
        public static readonly Guid globPurchasesGuid = new Guid("e2d7d9f6-7ec6-49f3-85f9-d765e2befcbe");
        // Guid типа "Закупка"
        public static readonly Guid purchaseTypeGuid = new Guid("7ed22847-90b3-4772-a860-0fb2be6956dd");
        // Guid параметра "Наименование" типа "Закупка"
        public static readonly Guid purchaseNameGuid = new Guid("e8867f88-0593-4922-a2c1-43cf4dac01e5");
        // Guid связи "Заявки" типа "Закупка"
        public static readonly Guid claimsLinkGuid = new Guid("663fddfa-906c-480c-a407-39ab6d45c025");

        // Guid справочника "Заявки на спецодежду"
        public static readonly Guid globClaimsGuid = new Guid("d4acba3f-c056-4b7f-8d13-f5ad3c6ae484");
        // Guid параметра "Размер" типа "Заявка"
        public static readonly Guid clothesSizeGuid = new Guid("97fba244-efdc-48cd-b388-86b9fc50a1a8");
        // Guid параметра "Рост" типа "Заявка"
        public static readonly Guid clothesHeightGuid = new Guid("86967290-f053-48fc-836c-c542bfe6a186");
        // Guid параметра "Материал" типа "Заявка"
        public static readonly Guid clothesMaterialGuid = new Guid("763e8900-9a57-45c4-8400-a17daa4483a9");
        // Guid параметра "Цвет" типа "Заявка"
        public static readonly Guid clothesColorGuid = new Guid("6e0b36b5-48a3-4c83-b006-581f1c9f45fa");
        // Guid параметра "Технические характеристики" типа "Заявка"
        public static readonly Guid clothesTechParamsGuid = new Guid("a36b04f0-e2b6-4edf-8e88-5ba7692997bc");
        // Guid параметра "Количество" типа "Заявка"
        public static readonly Guid clothesCountGuid = new Guid("615696de-0454-4345-83f3-6e9226e14125");
        // Guid параметра "Дата получения" типа "Заявка"
        public static readonly Guid clothesDateGuid = new Guid("77774b7f-01a7-4f49-8590-fec2e42c6b6b");
        // Guid параметра "Примечание" типа "Заявка"
        public static readonly Guid commentGuid = new Guid("5f13093c-d325-4158-bdfb-fc3749e4a171");

        // Guid связи "Получатель спецодежды" типа "Заявка"
        public static readonly Guid reqUserLinkGuid = new Guid("a1581296-e8bf-4047-9354-78d2e066c6fc");
        // Guid параметра "ФИО" типа "Получатель спецодежды"
        public static readonly Guid reqUserFIOGuid = new Guid("2ea91551-803f-4fe3-a581-afa22450acf4");
        // Guid параметра "Профессия" типа "Получатель спецодежды"
        public static readonly Guid reqUserProfessionGuid = new Guid("5c875a64-cc9a-41a3-861a-38dde9cbd352");
        // Guid параметра "Пол" типа "Получатель спецодежды"
        public static readonly Guid reqUserGenderGuid = new Guid("9ede7dc7-eacb-446b-a640-2a424fc5b015");
        // Guid связи "Подразделение" типа "Получатель спецодежды"
        public static readonly Guid departmentLinkGuid = new Guid("731a53ee-163a-485a-aaf0-2a9e6fff81b5");

        // Guid связи "Норматив из справочника" типа "Заявка"
        public static readonly Guid clothesTypeLinkGuid = new Guid("77663a83-4c61-4775-a6d8-23438761645f");
        // Guid параметра "Вид продукции" типа "Заявка"
        public static readonly Guid clothesNameGuid = new Guid("dfe4d36b-029a-4053-8f46-3164cc194579");
        // Guid параметра "Единица измерения" типа "Норматив спецодежды"
        public static readonly Guid measureUnitGuid = new Guid("95806d20-75d0-475d-bb46-db4210e8c451");
    }
}
