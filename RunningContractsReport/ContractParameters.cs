using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RunningContractsReport
{
    public class ContractParameters
    {
        public static Guid RKKReference = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");  //Guid Справочника РКК
        public static readonly Guid Users = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea"); //Guid Справочника Группы и пользователи
        public static readonly Guid Department908 = new Guid("f452318c-ff6c-457e-bbcb-e9359115ef21"); //Guid Отдел 908

        // типы справочника "Регистрационно-контрольные карточки"
        public static Guid DODType = new Guid("a9d4a8c0-da13-451c-9464-1e5c4c995fa6");   // тип Договор основной деятельности
        public static Guid DZUType = new Guid("1614e166-81d1-492a-942e-4b03970fcf8a");   // тип Договор закупки услуги
        public static Guid KZType = new Guid("e5fafaff-1925-4612-947a-b84f55687f3b");    // тип Карточка заказа 

        // Guids для ДОД
        public static class DODGuids
        {
            public static Guid RegNumber = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");      // Guid параметра "Регистрационный номер" справочника "Регистрационно-контрольные карточки"
            public static Guid From = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");           // Guid параметра "От" справочника "Регистрационно-контрольные карточки"
            public static Guid Contractor = new Guid("a612e412-e461-46f5-8cda-53797513ea3b");     // Guid параметра "Контрагент"
            public static Guid CostWithNDS = new Guid("27581738-efa8-44a7-af7a-de3545f29d0a");    // Guid Цена с НДС (сумма)
            public static Guid CostNDS = new Guid("7530f237-5627-4c19-8d53-2a1e1ee2849d");        // Guid в т.ч. НДС 
            public static Guid Currency = new Guid("c5eed221-622f-4266-b085-edb3ffe1a560");       // Guid валюта договора
            public static Guid Content = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");        // Guid предмет договора
            public static Guid DateClose = new Guid("e5a4e797-19b9-4ef7-9ee4-8942e5fb02fa");      // Guid дата приказа на закрытие
            public static Guid NumberOrderClose = new Guid("513dc775-9c30-45ff-aa67-663ade0d453f"); // Guid номер приказа на закрытие
            public static Guid NumberOrder = new Guid("33ef2a9f-07b1-4809-9a42-0b25f06a7ff8");      // Guid номер заказа(этапа) 
            public static Guid AdvancePaymentPlan = new Guid("d566db53-0811-4572-b67c-cb27677e03d7");  // Guid Планируемая величина аванса (берется из КЗ)
            public static Guid NumberOrderKZ = new Guid("c062de33-b890-4ffe-925b-6a88e068639f");  // Guid Номер заказа(этапа) для КЗ
            public static Guid DateEndKZ = new Guid("f9182505-c932-44f9-b9de-88ddcf69d2c0");    // Guid Дата окончания заказа(этапа) для КЗ
            public static Guid RealNumber = new Guid("12600c0e-d487-49a8-b77d-cecb060e3f59");    //Номер оригинального договора ДОД
        }

        // Guids для ДЗУ
        public static class DZUGuids
        {
            public static Guid RegNumber = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");      // Guid параметра "Регистрационный номер" справочника "Регистрационно-контрольные карточки"
            public static Guid From = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");           // Guid параметра "От" справочника "Регистрационно-контрольные карточки"
            public static Guid Contractor = new Guid("a2d8a1f1-0e4f-41f0-ad7d-512fe751d157");     // Guid параметра "Контрагент"
            public static Guid CostWithNDS = new Guid("5f357c2c-58bb-456e-86bd-38b1abd7e42c");    // Guid параметра Цена с НДС (сумма)
            public static Guid CostNDS = new Guid("a429e396-2779-4e82-859e-c4894a3a1651");        // Guid в т.ч. НДС
            public static Guid Currency = new Guid("4bef6151-10d5-4f43-b6b8-45c54c3ba02d");       // Guid валюта договора
            public static Guid Content = new Guid("bc62cf57-65a5-47e2-a3e2-7dc38c47c25a");        // Guid Предмет договора
            public static Guid RealDate = new Guid("89a50a7f-2db1-48d7-9b84-e24afa2a5da7");        //Дата ДЗУ
            public static Guid RealNumber = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4");      //Номер оригинального договора ДЗУ

        }

        public static Guid DocumentLinks = new Guid("f5d17767-10c8-4df7-9670-cfbc6b2590f7");  // Guid связи "Документы" справочника "Регистрационно-контрольные карточки"

    }
}
