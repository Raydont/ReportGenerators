using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentsListReport
{
    public class TFDClass
    {
        // Guid справочника РКК
        public static readonly Guid regCardsGuid = new Guid("80831DC7-9740-437C-B461-8151C9F85E59");

        // Guid типа "Исходящий документ"
        public static readonly Guid inputDocGuid = new Guid("885d0647-fbcc-4670-88f0-8a71cc3b19e1");
        // Guid типа "Приказ по предприятию"
        public static readonly Guid orderDocGuid = new Guid("ac00966a-3fc5-497b-b6ed-7fd7c49f434e");
        // Guid типа "Служебная записка"
        public static readonly Guid serviceGuid = new Guid("a006bf79-78b2-4656-a00d-9db1194f0089");
        // Guid типа "Распоряжение по предприятию"
        public static readonly Guid decreeDocGuid = new Guid("6608f368-9d7e-45b8-90d2-a9e4731a2545");
        // Guid типа "Докладная записка"
        public static readonly Guid reportDocGuid = new Guid("f6114524-df5a-477a-9d54-b03181c283ac");
        // Guid типа "Договор закупки/услуги"
        public static readonly Guid contractDocGuid = new Guid("1614e166-81d1-492a-942e-4b03970fcf8a");

        /* Общие параметры РКК*/
        // Guid параметра "От"
        public static readonly Guid regDateGuid = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
        // Guid параметра "Номер"
        public static readonly Guid regNumberGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
        // Guid связи "Подписал / Кому"
        public static readonly Guid RCUsersLinkGuid = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce");
        // Guid связи "Исполнитель"
        public static readonly Guid executorLink = new Guid("17330c2c-589f-4838-b6d1-8a9214666c4e");
        // Guid параметра "Получатели"
        public static readonly Guid receiversList = new Guid("a5a76a9d-0b32-4e3b-9124-93c193c2330f");
        // Guid связи "Документы"
        public static readonly Guid RCDocLinkGuid = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5");

        /* Параметры исходящего документа*/
        // Guid параметра "Краткое содержание" исходящего документа
        public static readonly Guid summaryGuid = new Guid("963e7d68-fc9a-4892-9518-86b892b854f4");
        // Guid параметра "Порядковый номер" исходящего документа
        public static readonly Guid orderNumberGuid = new Guid("a61a19d6-03f5-4b88-b092-54288686b2f8");
        // Guid списка "Список получателей" исходящего документа
        public static readonly Guid receiversLink = new Guid("26dc1c05-7f4c-46ff-a95e-7a0620c8f6c1");
        // Guid параметра "Адресат" списка получателей исходящего
        public static readonly Guid receicerName = new Guid("8d05c0a6-7f61-4cd1-ab97-7b88036e1c8d");
        // Guid параметра "Город" списка получателей исходящего
        public static readonly Guid receiverCity = new Guid("822061bb-b793-44ee-9010-d7f4f4dae540");

        /* Параметры приказа*/
        // Guid параметра "Содержание" приказа по предприятию
        public static readonly Guid contentGuid = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
        // Guid параметра "Подписал"
        public static readonly Guid signGuid = new Guid("b4de3b67-5ca7-42d7-97d9-c52b2e780dab");

        /* Параметры служебной записки*/
        // Guid параметра "Краткое содержание" служебной записки
        public static readonly Guid contentSZGuid = new Guid("8919ca23-78b3-49af-8c47-373454ac2fbc");
        // Guid параметра "Исполнитель" служебной записки
        public static readonly Guid executorSZ = new Guid("654ee95c-b576-4732-8acb-2349653be414");       

        /* Параметры докладной записки*/
        // Guid параметра "Номер подразделения"
        public static readonly Guid codeDepartmentDZ = new Guid("2c527c39-7fcc-47f1-847d-c79efacbf1d8");
        // Guid параметра "Исполнитель докладной записки"
        public static readonly Guid executorDZ = new Guid("f6ea3be0-0457-4cbf-b496-b8ff68cccafd");
        // Guid параметра "Краткое содержание" докладной записки
        public static readonly Guid contentDZGuid = new Guid("903352f2-3d7f-413e-998f-ff496d7f41c1");
        // Guid параметра "Подписал" докладной записки
        public static readonly Guid DZSenderGuid = new Guid("ced9424f-05cf-4720-b574-19ce4a6b4407");

        /* Параметры договора закупки/услуги*/
        // Guid параметра "Рамочный договор"
        public static readonly Guid frameContractDZUGuid = new Guid("85e45613-ec2d-45b3-89b5-5bfddf2b6a5d");
        // Guid параметра "Краткое содержание ДЗУ"
        public static readonly Guid contentDZUGuid = new Guid("bc62cf57-65a5-47e2-a3e2-7dc38c47c25a");
        // Guid параметра "Номер договора"
        public static readonly Guid orderNumDZUGuid = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4");
        // Guid параметра "Контрагент"
        public static readonly Guid contractorDZUGuid = new Guid("a2d8a1f1-0e4f-41f0-ad7d-512fe751d157");
        // Guid параметра "Сумма с НДС"
        public static readonly Guid sumWithNDSDZUGuid = new Guid("5f357c2c-58bb-456e-86bd-38b1abd7e42c");
        // Guid связи "Рамочный договор"
        public static readonly Guid frameContractLink = new Guid("97c6c17d-fd80-4573-96db-e6dd1775dfb9");

        // Guid справочника "Группы и пользователи"
        public static readonly Guid usersGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        // Guid свойства Login
        public static readonly Guid loginGuid = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");
        // Guid роли "Канцелярия предприятия"
        public static readonly Guid chancelloryGuid = new Guid("39bfdc6b-6526-478e-82cc-95b559f72e45");
        // Guid должности "Администратор"
        public static readonly Guid adminGuid = new Guid("921CE1A4-5DBA-4742-8D3F-27D13A78B0C7");
        // Guid типа "Производственное подразделение"
        public static readonly Guid manufactoryDepartment = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");
        // Guid параметра "Код подразделения"
        public static readonly Guid codeDepartmentGuid = new Guid("6d58df61-2a43-45b2-b8f6-34f3507107fd");
        // Guid параметра "Наименование подразделения"
        public static readonly Guid nameDepartmentGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");

        // Guid стадии "Закрыто" у документа
        public static readonly Guid closeStage = new Guid("8e12d323-25e1-49a1-8f7c-f874326f6ad2");
        // Guid стадии "Согласуется" у документа
        public static readonly Guid processStage = new Guid("437647c1-028a-4b99-8616-bee98e045ff5");
        // Guid стадии "Доработка" у документа
        public static readonly Guid returnStage = new Guid("9fe58f4a-90b2-43e9-bc23-21addab0bf5b");
        // Guid стадии "Отклонено" у документа
        public static readonly Guid cancelStage = new Guid("95c7129c-34c8-4da6-91d3-ebecf0eb83c1");               
    }
}
