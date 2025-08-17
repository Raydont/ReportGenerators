using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PaymentsReport
{
    public static class Guids
    {
        public static class RegCardsReference
        {
            // Guid справочника РКК
            public static readonly Guid refGuid = new Guid("80831DC7-9740-437C-B461-8151C9F85E59");
            public static class Fields
            {
                // Guid параметра "Регистрационный Номер" справочника РКК
                public static readonly Guid registrationNumberGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
                // Guid параметра "От" (дата регистрации документа)
                public static readonly Guid registrationDateGuid = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
                // Guid параметра "Примечание"
                public static readonly Guid commentGuid = new Guid("d66a5469-d1a6-4820-92dd-7a52299c7c03");
            }
            public static class Links
            {
                // Guid связи "Документы"
                public static readonly Guid DocumentLinks = new Guid("aab1cd65-fe2d-4fc0-8db2-4ff2b3d215d5");
            }
            // РКК типа "Служебная записка на оплату" 
            public static class SZNO
            {
                public static readonly Guid typeGuid = new Guid("2A3E0993-8CDF-4EBB-BF13-0C187DCDBC9E");
                // Параметры объектов типа СЗНО
                public static class Fields
                {
                    // Guid параметра "Номер счета" типа СЗНО
                    public static readonly Guid invoiceNumberGuid = new Guid("6e527664-a3b3-4b70-b602-3421c3e05abd");
                    // Guid параметра "Контрагент" типа СЗНО
                    public static readonly Guid contractorGuid = new Guid("ed34fb5e-d8a4-4561-8f81-480edb85d11b");
                    // Guid параметра "Наименование товара, услуги"->"Содержание" типа СЗНО
                    public static readonly Guid contentGuid = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
                    // Guid параметра "Назначение платежа" типа СЗНО
                    public static readonly Guid paymentFunctionGuid = new Guid("1f1b0d95-b598-4f6d-8ae4-539c3c7fdca8");
                    // Guid параметра "Вид оплаты" типа СЗНО
                    public static readonly Guid paymentKindGuid = new Guid("6480abc4-a8cd-4f85-bf12-416382d9ff96");
                    // Guid параметра "Процент" типа СЗНО
                    public static readonly Guid percentGuid = new Guid("77e7b8b7-9f49-4392-915d-50d40d103a5e");
                    // Guid параметра "Сумма к оплате" типа СЗНО
                    public static readonly Guid paymentSumGuid = new Guid("f962d417-de7d-4723-ba9c-ed736271ab22");
                    // Guid параметра "Номер заказа" типа СЗНО
                    public static readonly Guid orderNumberGuid = new Guid("89dc03b6-14da-4a37-90e2-a45310caee29");
                    // Guid параметра "Регистрационный номер договора" типа СЗНО
                    public static readonly Guid contractRegNumGuid = new Guid("4b2ad646-909b-4886-8479-004d591f50bf");
                    // Guid параметра "Номер договора" типа СЗНО
                    public static readonly Guid contractNumberGuid = new Guid("f430baef-49e8-4ead-a375-5edcf2f31dd7");
                    // Guid параметра "Пункт ПТП" типа СЗНО
                    public static readonly Guid ptpPointGuid = new Guid("ecde8aa6-f96e-4bd2-8fba-daa3345b5608");
                    // Guid параметра "Номер платежного поручения"
                    public static readonly Guid draftNumberGuid = new Guid("82987a91-04b5-45e8-a3a5-925cef8a1c46");
                    // Guid параметра "Дата платежного поручения"
                    public static readonly Guid draftDateGuid = new Guid("1c8ab3ba-e5e9-41d2-8cfe-100ca505c59c");
                    // Guid параметра "Дата утверждения"
                    public static readonly Guid confirmDateGuid = new Guid("fe8e8c1e-91f6-4e82-85ff-f12274c5aa20");
                    // Guid параметра "Оплачено"
                    public static readonly Guid paidGuid = new Guid("094b0823-a3cb-447d-92ed-a0a1873b6637");                     // Для рабочей базы

                }
                public static class Links
                {
                    // Guid связи "Договора закупки-услуги"
                    public static readonly Guid DZULinks = new Guid("a3b057b8-6072-42c3-badc-1623dbf02c16");
                    // Guid связи справочника со справочником "Статьи бюджета"
                    public static readonly Guid budgetArticlesLink = new Guid("4a33eea5-696b-44c5-a5ee-76a440842440");
                    // Guid списка статей бюджета
                    public static readonly Guid budgetArticlesListLink = new Guid("64545b55-eab5-4824-8527-cbd4c89419ce");        // Для рабочей базы
                }
                public static class BudgetArticlesList
                {
                    public static readonly Guid budgetArticleCode = new Guid("b10ce6fa-00b5-47fc-a6de-de2186369c4e");             // Для рабочей базы
                    public static readonly Guid paymentSum = new Guid("1d735473-459e-4d11-a2c5-569c08a817e3");                    // Для рабочей базы
                }
            }
            // РКК типа "Договор закупки-услуги" 
            public static class DZU
            {
                // Параметры объектов типа ДЗУ
                public static class Fields
                {
                    // Guid параметра "Регистрационный номер ДЗУ"
                    public static readonly Guid regNumberGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
                    // Guid параметра "Номер договора"
                    public static readonly Guid numberGuid = new Guid("40164d61-5172-4e85-9ab8-51ac1d27f1e4");
                }
            }
            // РКК типа "Реестр платежей и поступлений"
            public static class RPP
            {
                public static class Fields
                {
                    // Guid параметра "Начало периода"
                    public static readonly Guid startDateGuid = new Guid("745a085f-fc17-476e-a434-fa26dc1a9fc6");
                    // Guid параметра "Конец периода"
                    public static readonly Guid finishDateGuid = new Guid("f70829fb-27c0-404f-a5e1-4093fd50c40a");
                }
            }
            // Стадии прохождения объектов справочника РКК
            public static class Stages
            {
                // Guid стадии "Утверждено" для платежного документа
                public static readonly Guid stageSignedGuid = new Guid("A5EA2E1C-D441-42FD-8F92-49840351D6C1");
                // Guid стадии "В архиве (СЗНО)"
                public static readonly Guid stageInArchiveGuid = new Guid("554f7f22-9c5b-4f5f-84e4-068360678ea7");
            }
        }
        public static class BudgetArticles
        {
            // Guid справочника "Статьи бюджета"
            public static readonly Guid refGuid = new Guid("dee585c0-368b-4e23-a50b-f3b86857e8ed");
            // Параметры объектов типа "Статья бюджета"
            public static class Fields
            {
                // Guid типа "Статья БДДС"
                public static readonly Guid budgetArticleGuid = new Guid("267c52d0-d6bc-4a07-8c50-83eb605f6ca1");
                // Guid параметра "Код статьи БДДС" справочника "Статьи бюджета"
                public static readonly Guid articleCodeGuid = new Guid("3c69827a-6b75-46e2-b1d0-fe1ba3dba4f1");
            }
        }
        public static class Users
        {
            // Guid справочника "Группы и пользователи"
            public static readonly Guid refGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
            // Параметры объектов типа "Пользователь"
            public static class Fields
            {
                // Guid параметра "Наименование"
                public static readonly Guid userNameGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
                // Guid параметра "Короткое имя"
                public static readonly Guid userShortNameGuid = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
                // Guid параметра "Образец подписи"
                public static readonly Guid signatureImageGuid = new Guid("4af0e216-71de-4040-8622-99281d8054bb");
                // Guid свойства Login
                public static readonly Guid loginGuid = new Guid("42c81c2b-7354-46aa-9547-0f1a93e9d4e1");
            }
            // Связи объекта типа "Пользователь", "Подразделение"
            public static class Links
            {
                // Guid параметра "Руководитель"
                public static readonly Guid depChiefGuid = new Guid("fdb41549-2adb-40b0-8bb5-d8be4a6a8cd2");
            }
            // Объекты справочника "Группы и пользователи"
            public static class Objects
            {
                // Guid должности "Зам.ген.директора по экономическим вопросам"
                public static readonly Guid chiefEconomistGuid = new Guid("317d6239-fa91-4319-a8b9-a919043ae415");
                // Guid должности "Специалист ЕРКЦ"
                public static readonly Guid specERKCGuid = new Guid("6aa682b1-f17c-45ac-a6c5-8c7ff4239c4e");
                // Guid должности "Главный бухгалтер"
                public static readonly Guid chiefAccountGuid = new Guid("b080563f-e38c-417a-9316-eed64c10555e");
                // Guid типа "Администратор"
                public static readonly Guid adminTypeGuid = new Guid("0bc8a939-2914-4e79-9621-3e03c39b9c31");
                // Guid отдела 101
                public static readonly Guid dep101Guid = new Guid("7fd1729a-5e8c-4794-9401-8f1333ad0fd2");
            }
        }
        public static class Files
        {
            // Guid справочника "Файлы"
            public static readonly Guid refGuid = new Guid("A0FCD27D-E0F2-4C5A-BBA3-8A452508E6B3");
            public static class Paths
            {
                // Путь до папки "Реестры платежей и поступлений"
                public static readonly string rppFolderPath = @"Хранилище ОРД\Реестры платежей и поступлений";
                // Путь до файла шаблона
                public static readonly string patternPath = @"Служебные файлы\Шаблоны отчётов\Отчеты Excel\Excel.xlsx";
            }
        }
    }
}
