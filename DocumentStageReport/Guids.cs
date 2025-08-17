using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentStagesReport
{
    public static class Guids
    {
        public static class RegCardsReference
        {
            // Guid справочника РКК
            public static readonly Guid RefGuid = new Guid("80831DC7-9740-437C-B461-8151C9F85E59");
            public static class Fields
            {
                // Guid параметра "От"
                public static readonly Guid RegDate = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
                // Guid связи "Процесс"
                public static readonly Guid Process = new Guid("22805070-d9ab-4929-ae72-918f3641afbf");
            }
        }
        public static class ProcedureReference
        {
            //Справочник процедуры
            public static Guid RefGuid = new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3");
            public static class Fields
            {
                //Доступ на запуск процедуры
                public static readonly Guid RunPremission = new Guid("cdee3402-fc51-48ae-855b-4505b8391cd5");
                //Условие для процедуры
                public static readonly Guid ConditionReferenceProcedure = new Guid("6e40f04c-ca99-43d3-b353-0a1e240abbe8");
                //Список объектов справочники в справочнике Процедуры
                public static readonly Guid ListObjectsReference = new Guid("e54a1ab2-eb94-4b79-be8d-3722cb1d5e45");
                //Список пользователей с правом на запуск
                public static readonly Guid ListUsersAccess = new Guid("84579eed-d8ea-404f-a8b7-a1b63c3db722");
                //Параметр справочника в справочнике Процедуры
                public static readonly Guid ParReferenceProcedure = new Guid("71fb1911-694f-4106-8641-7b27516d2c3b");
                //Список объектов Состояния
                public static Guid ListObjectsState = new Guid("de0cdb34-aee3-459a-bf97-74094270cef9");
                //Наименование процедуры
                public static readonly Guid NameProcedure = new Guid("5e99d9c1-4d74-4d62-8806-aea35c4c792e");
            }
            public static class Object
            {
                // Согласование ДЗУ
                public static readonly Guid ApprovalDZU = new Guid("5a6687f4-a0dd-4f11-a4bf-ebb6b0f38ecd");
                // Служебная записка на оплату
                public static readonly Guid ApprovalSZNO = new Guid("1de319ec-36b5-439a-b905-b35806401551");
                // Служебная записка на оплату (налоги и зарплата)
                public static readonly Guid TaxesSZNO = new Guid("1de319ec-36b5-439a-b905-b35806401551");
                // Служебная записка
                public static readonly Guid ApprovalSZ = new Guid("0ef0d60a-abc2-4fcb-8541-02b12b73ef0d");
                // Заявка на проведение закупочной процедуры
                public static readonly Guid ApprovalZPZP = new Guid("b762060e-afc5-4666-b227-78ce5b975589");
                // Служебная записка на поступление денежных средств
                public static readonly Guid ApprovalSZNPDS = new Guid("18291c38-d71a-40c7-aecd-e7baa29cc6bc");
                // Запись об изменениях КД
                public static readonly Guid ApprovalZICD = new Guid("4b4d412f-5743-46b1-ab6a-336a51dd77a2");
                // Наряд-заказ
                public static readonly Guid ProductOrder = new Guid("bb4d3cc6-8fe9-4e07-8332-cb54043f74eb");
            }
        }
        public static class ProcessReference
        {
            //Справочник процессы
            public static readonly Guid RefGuid = new Guid("e0c70b3c-bee1-4321-87b0-c44f3d8b5f68");
            public static class Fields
            {
                //Статус процесса
                public static readonly Guid ProcessStatus = new Guid("cc06378b-4987-4322-a69d-6bcbb295cdf7");
                //Запущенные объекты по процессу
                public static readonly Guid ProcessStartObjects = new Guid("cf72b950-52d9-4089-a07b-efad09ed613c");
                //Связь процесса с процедурой
                public static readonly Guid LinkProcessToProcedure = new Guid("44c3a825-ef4a-4ecb-a291-1048db42d5cb");
                // связь процессов и справочника действия
                public static Guid LinkProcessToAction = new Guid("56cf584b-7206-47d1-9aa8-81faf7484dc8");
                //Текущие действия процесса
                public static readonly Guid ActiveActionsLinkGuid = new Guid("721f2a2a-3dc2-4578-b3c5-4224f8d8a1e0");
            }
        }
        public static class ProcessActionsReference
        {
            public static readonly Guid RefGuid = new Guid("60e48c95-9d85-4a04-8bed-795bc65b61a4");
            public static class Types
            {
                //Тип Состояние процесса
                public static Guid StateActionType = new Guid("5505cf81-c0c3-46c1-9116-38b7a1b2d441");
                //Тип Пользовательское состояние процесса
                public static Guid UserStateActionType = new Guid("ed731de4-2568-4148-94c8-a15634a0303b");
            }
            public static class Fields
            {
                //идентификатор состояния этапа процесса для текущего действия
                public static readonly Guid IdentificatorStepGuid = new Guid("a3c276dc-ca53-4ada-9e36-fb2c49efe8ae");
                //параметры состояния процесса
                public static Guid StateParameters = new Guid("9cb2da09-f671-436d-9797-2e5d0eaf868e");
                //Идентификатор состояния
                public static Guid IdenticatorState = new Guid("9045a6ad-b3d6-4420-967a-29e7b69289f9");
                //Время начала
                public static readonly Guid StartTime = new Guid("2b2321d5-c4bc-4a8a-8820-7c31d1b2564a");
                //Время завершения
                public static readonly Guid EndTime = new Guid("eef93b06-8ec7-472f-acfc-cc1b2ee7a12f");
                //Исполнители действия процесса
                public static readonly Guid MakersProcessActions = new Guid("3f93b7c3-fdf2-4411-b342-dba16f927576");
                //Связь исполнители
                public static Guid LinkMakerToUser = new Guid("07d428d2-a222-4b2f-b4b9-a936b437cc80");
            }
        }
        public static class UsersReference
        {
            public static class Objects
            {
                public static readonly Guid DeadlineControl = new Guid("685ae049-13bf-401a-a147-e4dfe41ddb9c");
            }
        }
    }
}
