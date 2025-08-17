using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentsListReport
{
    public class Guids
    {
        public static class RegCardsReference
        {
            // Guid справочника РКК
            public static readonly Guid refGuid = new Guid("80831DC7-9740-437C-B461-8151C9F85E59");
            public static class Fields
            {
                // Guid параметра "От"
                public static readonly Guid regDate = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
                // Guid параметра "Номер"
                public static readonly Guid regNumber = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
                // Guid параметра "Содержание"
                public static readonly Guid content = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
                // Guid параметра "Номер подразделения"
                public static readonly Guid departmentNumber = new Guid("788dbd76-6ffd-41a9-9df9-82b80bd0a8a3");
            }
            public static class Links
            {
                // Guid связи "Получатели"
                public static readonly Guid receiversLink = new Guid("a5a76a9d-0b32-4e3b-9124-93c193c2330f");
                // Guid связи "Подписал / Кому"
                public static readonly Guid rCUsersLink = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce");
                // Guid связи "Исполнитель"
                public static readonly Guid executorLink = new Guid("17330c2c-589f-4838-b6d1-8a9214666c4e");
                // Guid связи "Документы"
                public static readonly Guid docLink = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5");
            }
            public static class InputDoc
            {
                // Guid типа "Исходящий документ"
                public static readonly Guid type = new Guid("885d0647-fbcc-4670-88f0-8a71cc3b19e1");
                public static class Fields
                {
                    // Guid параметра "Краткое содержание" исходящего документа
                    public static readonly Guid summary = new Guid("963e7d68-fc9a-4892-9518-86b892b854f4");
                    // Guid параметра "Порядковый номер" исходящего документа
                    public static readonly Guid orderNumber = new Guid("a61a19d6-03f5-4b88-b092-54288686b2f8");
                }
                public static class Links
                {
                    // Guid списка "Список получателей" исходящего документа
                    public static readonly Guid receiversLink = new Guid("26dc1c05-7f4c-46ff-a95e-7a0620c8f6c1");
                }
                public static class Receivers
                {
                    // Guid параметра "Адресат" списка получателей исходящего
                    public static readonly Guid receicerName = new Guid("8d05c0a6-7f61-4cd1-ab97-7b88036e1c8d");
                    // Guid параметра "Город" списка получателей исходящего
                    public static readonly Guid receiverCity = new Guid("822061bb-b793-44ee-9010-d7f4f4dae540");
                }
            }
            public static class OrderDoc
            {
                // Guid типа "Приказ по предприятию"
                public static readonly Guid type = new Guid("ac00966a-3fc5-497b-b6ed-7fd7c49f434e");
                public static class Fields
                {
                    // Guid параметра "Содержание" приказа по предприятию
                    public static readonly Guid content = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
                    // Guid параметра "Подписал"
                    public static readonly Guid signed = new Guid("b4de3b67-5ca7-42d7-97d9-c52b2e780dab");
                }
            }
            public static class ServiceDoc
            {
                // Guid типа "Служебная записка"
                public static readonly Guid type = new Guid("a006bf79-78b2-4656-a00d-9db1194f0089");
                public static class Fields
                {
                    // Guid параметра "Краткое содержание" служебной записки
                    public static readonly Guid content = new Guid("8919ca23-78b3-49af-8c47-373454ac2fbc");
                    // Guid параметра "Исполнитель" служебной записки
                    public static readonly Guid executor = new Guid("654ee95c-b576-4732-8acb-2349653be414");
                }
            }
            public static class DecreeDoc
            {
                // Guid типа "Распоряжение по предприятию"
                public static readonly Guid type = new Guid("6608f368-9d7e-45b8-90d2-a9e4731a2545");
            }
            public static class ReportDoc
            {
                // Guid типа "Докладная записка"
                public static readonly Guid type = new Guid("f6114524-df5a-477a-9d54-b03181c283ac");
                public static class Fields
                {
                    // Guid параметра "Исполнитель докладной записки"
                    public static readonly Guid executor = new Guid("f6ea3be0-0457-4cbf-b496-b8ff68cccafd");
                    // Guid параметра "Краткое содержание" докладной записки
                    public static readonly Guid content = new Guid("903352f2-3d7f-413e-998f-ff496d7f41c1");
                    // Guid параметра "Подписал" докладной записки
                    public static readonly Guid sender = new Guid("ced9424f-05cf-4720-b574-19ce4a6b4407");
                }
            }
            public static class ContractDoc
            {
                // Guid типа "Договор закупки/услуги"
                public static readonly Guid type = new Guid("1614e166-81d1-492a-942e-4b03970fcf8a");
                public static class Fields
                {
                    // Guid параметра "Рамочный договор" (Да / Нет)
                    public static readonly Guid frameContract = new Guid("85e45613-ec2d-45b3-89b5-5bfddf2b6a5d");
                    // Guid параметра "Краткое содержание ДЗУ"
                    public static readonly Guid content = new Guid("bc62cf57-65a5-47e2-a3e2-7dc38c47c25a");
                    // Guid параметра "Номер договора"
                    public static readonly Guid contractNum = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4");
                    // Guid параметра "Контрагент"
                    public static readonly Guid contractor = new Guid("a2d8a1f1-0e4f-41f0-ad7d-512fe751d157");
                    // Guid параметра "Сумма с НДС"
                    public static readonly Guid sumWithNDS = new Guid("5f357c2c-58bb-456e-86bd-38b1abd7e42c");
                    // Guid параметра "Номер заказа"
                    public static readonly Guid orderNum = new Guid("103ef275-bedc-4751-a7d0-3eaf74a341f3");
                }
                public static class Links
                {
                    // Guid связи "Рамочный договор"
                    public static readonly Guid frameContractLink = new Guid("97c6c17d-fd80-4573-96db-e6dd1775dfb9");
                }
            }
            public static class PaymentServiceDoc
            {
                // Guid типа "Служебная записка на оплату"
                public static readonly Guid type = new Guid("2a3e0993-8cdf-4ebb-bf13-0c187dcdbc9e");
                public static class Fields
                {
                    // Guid параметра "Оплачено"
                    public static readonly Guid paid = new Guid("094b0823-a3cb-447d-92ed-a0a1873b6637");
                    // Guid параметра "Контрагент"
                    public static readonly Guid contractor = new Guid("ed34fb5e-d8a4-4561-8f81-480edb85d11b");
                    // Guid параметра "Наименование товара, услуги"
                    public static readonly Guid productName = new Guid("d6be5cb3-b4df-40e7-9b0e-fde1c6a2ce56");
                    // Guid параметра "Сумма к оплате"
                    public static readonly Guid paymentSum = new Guid("f962d417-de7d-4723-ba9c-ed736271ab22");
                    // Guid параметра "Номер договора"
                    public static readonly Guid contractNumber = new Guid("f430baef-49e8-4ead-a375-5edcf2f31dd7");
                }
                public static class Links
                {
                    public static readonly Guid contractLink = new Guid("a3b057b8-6072-42c3-badc-1623dbf02c16");
                    // Guid списка статей бюджета
                    public static readonly Guid budgetArticlesListLink = new Guid("64545b55-eab5-4824-8527-cbd4c89419ce");
                }
                public static class BudgetArticlesList
                {
                    // Guid параметра "Код статьи БДДС"
                    public static readonly Guid budgetArticleCode = new Guid("b10ce6fa-00b5-47fc-a6de-de2186369c4e");
                    // Guid параметра "Сумма"
                    public static readonly Guid paymentSum = new Guid("1d735473-459e-4d11-a2c5-569c08a817e3");
                }
            }
        }
        public static class UserReference
        {
            // Guid справочника "Группы и пользователи"
            public static readonly Guid refGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
            public class Types
            {
                // Guid типа "Производственное подразделение"
                public static readonly Guid department = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");
                // Guid типа "Администратор"
                public static readonly Guid admin = new Guid("0bc8a939-2914-4e79-9621-3e03c39b9c31");
            }
            public class Users
            {
                public static class Fields
                {
                    // Guid свойства Login
                    public static readonly Guid login = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");
                }
            }
            public class Departments
            {
                public static class Fields
                {
                    // Guid параметра "Код подразделения"
                    public static readonly Guid code = new Guid("6d58df61-2a43-45b2-b8f6-34f3507107fd");
                    // Guid параметра "Наименование подразделения"
                    public static readonly Guid name = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
                }
                public static class Objects
                {
                    // Guid АО "РКБ "Глобус"
                    public static readonly Guid globus = new Guid("d54c3bcb-9c4d-4b73-b214-c341e7a8ea95");
                    // Guid подразделения "Руководство предприятия"
                    public static readonly Guid chiefs = new Guid("8285a299-6f49-4139-adad-54ffee622c09");
                    // Guid роли "Канцелярия предприятия"
                    public static readonly Guid chancellory = new Guid("39bfdc6b-6526-478e-82cc-95b559f72e45");
                    // Guid должности "Администратор"
                    //public static readonly Guid admin = new Guid("921CE1A4-5DBA-4742-8D3F-27D13A78B0C7");
                }
            }
        }
        public static class DocumentStages
        {
            // Guid стадии "Закрыто" у документа
            public static readonly Guid closeStage = new Guid("8e12d323-25e1-49a1-8f7c-f874326f6ad2");
            // Guid стадии "Согласуется" у документа
            public static readonly Guid processStage = new Guid("437647c1-028a-4b99-8616-bee98e045ff5");
            // Guid стадии "Доработка" у документа
            public static readonly Guid returnStage = new Guid("9fe58f4a-90b2-43e9-bc23-21addab0bf5b");
            // Guid стадии "Отклонено" у документа
            public static readonly Guid cancelStage = new Guid("95c7129c-34c8-4da6-91d3-ebecf0eb83c1");
            // Guid стадии "Утверждено" для платежного документа
            public static readonly Guid signedStage = new Guid("A5EA2E1C-D441-42FD-8F92-49840351D6C1");
            // Guid стадии "В архиве (СЗНО)"
            public static readonly Guid inArchiveStage = new Guid("554f7f22-9c5b-4f5f-84e4-068360678ea7");
        }
    }
}
