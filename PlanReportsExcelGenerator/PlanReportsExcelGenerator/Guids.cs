using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Globus.PlanReportsExcel
{
    /// <summary>
    /// Guid-ы справочника "Планы и работы"
    /// </summary>
    public static class PlanReferenceGuids
    {
        /// <summary>
        /// guid справочника
        /// </summary>
        public static Guid guid = new Guid("4217e186-6656-40b2-9d4d-f8720ce191ac");
        /// <summary>
        /// параметр "Наименование"
        /// </summary>
        public static Guid NameParameter = new Guid("acb70580-7c8b-48c0-b718-7927dde85e75");
        /// <summary>
        /// параметр "Номер пункта"
        /// </summary>
        public static Guid PointNumberParameter = new Guid("f4895b0b-9576-4d71-ab3c-e8125795d602");       

        /// <summary>
        /// типы планов
        /// </summary>
        public static class Plan
        {
            /// <summary>
            /// тип "Простой план" 
            /// </summary>
            public static Guid SimplePlan = new Guid("c8e18eb0-9b39-4299-81b9-775eb64d73c7");
            /// <summary>
            /// тип "График изготовления"
            /// </summary>
            public static Guid ProductionPlan = new Guid("be1497c7-5db1-4fc1-b894-3b98e9fc9895");
            /// <summary>
            /// тип "График предварительных испытаний"
            /// </summary>
            public static Guid PreliminaryTestPlan = new Guid("c855c0fc-8df7-46c7-bb93-a96d0f18d9ea");
            /// <summary>
            /// тип "График разработки документации"
            /// </summary>
            public static Guid DocDevelopPlan = new Guid("77cf32d2-e458-4fd3-8caa-8a3b4bc2d391");
            /// <summary>
            /// тип "График периодических испытаний"
            /// </summary>
            public static Guid PeriodicTestPlan = new Guid("73303651-e232-4843-9490-2c5e188822c6");

            /// <summary>
            /// список объектов "Блоки подписей"
            /// </summary>
            public static class SignBlocks
            {
                public static Guid guid = new Guid("25c3a064-7559-436a-bd9e-aebf6aedcfac");

                public static Guid typeGuid = new Guid("42a4b4be-b70d-416e-994c-43347177f41f");       // guid типа списка объектов

                // --------- ПАРАМЕТРЫ --------
                public static Guid FIO = new Guid("5b0c8f38-e2ee-4c92-8552-df842a46471f");              // ФИО
                public static Guid SignBlockType = new Guid("6940cbaf-5e10-42fb-a119-0c9ba3a34053");    // Тип блока подписи
                public static Guid SignBlockPos = new Guid("aa044355-252c-405c-98b3-804de64c37e9");     // Позиция блока подписи в таблице
                public static Guid SignBlockPost = new Guid("8c671a7c-7fcb-494d-aa1c-234ed0c727a7");    // Должность блока подписи в таблице
            }

            // ---------------------- ПАРАМЕТРЫ -----------------------------------------------------
            /// <summary>
            /// параметр "Колонок в таблице подписей"
            /// </summary>
            public static Guid ColInSignBlockTable = new Guid("b3cc05b7-002c-4e5f-8f1e-982d3c3ebe45");
            /// <summary>
            /// параметр "Дата утверждения"
            /// </summary>
            public static Guid ConfirmDateParameter = new Guid("02ef5697-9316-408c-a522-c34e6eebc281");
            /// <summary>
            /// параметр "Подразделение-составитель"
            /// </summary>
            public static Guid DepartmentAuthor = new Guid("b580c477-e0e2-4c00-8d2c-e2dcfea0f283");      // !!!!!
        }


        /// <summary>
        /// типы работ
        /// </summary>
        public static class Work
        {
            /// <summary>
            /// тип "Работа" 
            /// </summary>
            public static Guid guid = new Guid("097af6a5-f945-460e-a9c8-5d08f8f2dc6c");

            /// <summary>
            /// параметр "Код"
            /// </summary>
            public static readonly Guid WorkCode = new Guid("990e1ea5-47b9-4e85-827a-3d5e39b68bba");
            /// <summary>
            /// параметр "Номер работы"
            /// </summary>
            public static readonly Guid WorkNumber = new Guid("f4895b0b-9576-4d71-ab3c-e8125795d602");     // !!!!!
            /// <summary>
            /// параметр "Наименование работы"
            /// </summary>
            public static readonly Guid WorkName = new Guid("acb70580-7c8b-48c0-b718-7927dde85e75");       // !!!!!
            /// <summary>
            /// параметр "Срок"
            /// </summary>
            public static readonly Guid Deadline = new Guid("e93b904a-246e-4ef6-92f9-5e82295bc2c7");       // !!!!!           
            /// <summary>
            /// параметр "Примечание"
            /// </summary>
            public static readonly Guid Comment = new Guid("5ac366b8-a718-47cf-b817-a45f043f68be");
            /// <summary>
/*            /// параметр "Код" (параметры работы)
            /// </summary>
            public static readonly Guid WorkParametersCode = new Guid("990e1ea5-47b9-4e85-827a-3d5e39b68bba");*/
            /// <summary>
            /// связь со списком объектов "Подработы"
            /// </summary>
            public static readonly Guid SubWorksLink = new Guid("2cbe96c7-bdaf-48d1-882e-fc06ec5618e3");
            /// <summary>
            /// связь "Исполнители" (со справочником "Группы и пользователи") - для планов типа Период.И, РД
            /// </summary>
            public static Guid ExecutorsLink = new Guid("829597ed-024f-492b-8952-8d0a73741718");               // !!!!!
            /// <summary>
            /// связь "Исполнитель" для планов типа ГИ
            /// </summary>
            public static Guid ExecutorPGLink = new Guid("ac5541cc-c8ac-4970-ac88-51f6921cce4e");              // !!!!!
            /// <summary>
            /// параметр "Исполнители"
            /// </summary>
            public static Guid ExecutorsString = new Guid("83d53569-97f4-4b37-9b81-59ceda1ebb22");          

            /// <summary>
            /// связь "Исполнение задания"
            /// </summary>
            public static class WorkExecutionLink
            {
                public static Guid guid = new Guid("61940561-1371-43ad-b281-0ef457be7e0f");
                public static Guid ExecutorDepartmentLink = new Guid("b2713ac2-f533-498c-a1ac-f39a3f36a52e");  // ?????
            }

            /// <summary>
            /// список объектов "Перенос сроков"
            /// </summary>
            public static class SlipdateObjects
            {
                public static Guid guid = new Guid("b96fbc98-e6f1-4919-bd0f-f24a62dbb994");                    // !!!!!
                // guid типа списка объектов                
                public static Guid typeGuid = new Guid("8df0b482-0cde-4c5b-a269-0915247dd8d5");                // ?????
            }


            /// <summary>
            /// тип "Работа пред. исп-й" 
            /// </summary>
            public static class PreliminaryTestWork
            {
                public static Guid guid = new Guid("b3a13644-a908-4fdd-b57b-cce639639373");
            }

            /// <summary>
            /// тип "Работа РД" 
            /// </summary>
            public static class DocDevelopWork
            {
                public static Guid guid = new Guid("e501e51e-10e1-4ee2-98cd-c7154449445d");
            }

            /// <summary>
            /// тип "Работа период.исп-й" 
            /// </summary>
            public static class PeriodicTestWork
            {
                public static Guid guid = new Guid("2124db2d-ef0d-41ff-88e4-c54bb3afd657");
                /// <summary>
                /// параметр "Акт ПИ"
                /// </summary>
                public static Guid ActPT = new Guid("57d8d71a-1ccf-4ee4-aca4-cc64a0bb0d85");
                /// <summary>
                /// параметр "Дата окончания действия акта ПИ"
                /// </summary>
                public static Guid ActFinishDate = new Guid("7ce5009f-7fd5-4d96-85e6-3724e33b23fd");     
                /// <summary>
                /// параметр "График изготовления"
                /// </summary>
                public static Guid ProductionGraph = new Guid("90aede82-adf2-40db-b293-916b99324bd9");                                   
            }

            /// <summary>
            /// тип "Работа ГИ" 
            /// </summary>
            public static class ProductionWork
            {
                public static Guid guid = new Guid("9430bb2b-300c-4b5b-992f-e81576651cd6");
                /// <summary>
                /// параметр "Децимальный номер"
                /// </summary>
                public static Guid DecimalNum = new Guid("dc0c507e-ae4f-4325-ac15-1259df86e868");
                /// <summary>
                /// параметр "Количество"
                /// </summary>
                public static Guid Amount = new Guid("6817f609-7ec9-408a-b092-9edaa23a6abd");
                /// <summary>
                /// связь "Исполнитель" -> "Разработчик"
                /// </summary>
                public static Guid Executor = new Guid("ac5541cc-c8ac-4970-ac88-51f6921cce4e");
                /// <summary>
                /// параметр "Номер наряд-заказа"
                /// </summary>
                public static Guid DocNumber = new Guid("8c519421-b9ba-4487-aab2-4544359f78ef");
            }

            /// <summary>
            /// тип "Подработа" 
            /// </summary>
            public static class SubWork
            {
                /// <summary>
                /// параметр "Тип"
                /// </summary>
                public static Guid SubWorkType = new Guid("cb830105-06d2-4624-ba07-516b033783a5");
                /// <summary>
                /// параметр "Срок"
                /// </summary>
                public static Guid Deadline = new Guid("e93b904a-246e-4ef6-92f9-5e82295bc2c7");
                /// <summary>
                /// параметр "Исполнитель"
                /// </summary>
                public static Guid ExecutorsLink = new Guid("35c9aeff-3827-48b9-b735-8c50ef337e18");
            }
        }

        /// <summary>
        /// типы заголовков
        /// </summary>
        public static class Header
        {
            /// <summary>
            /// абстрактный тип "Заголовок" 
            /// </summary>
            public static Guid ParentGuid = new Guid("c7a5ad15-d6ec-4ac0-a595-9c0aa81caa06");
            /// <summary>
            /// тип "Заголовок простого плана" 
            /// </summary>
            public static Guid SimplePlanHeader = new Guid("7d617171-42bf-4119-980a-152c3e8a319e");
            /// <summary>
            /// тип "Заголовок ГИ" 
            /// </summary>
            public static Guid ProductionPlanHeader = new Guid("953f9f08-a0a8-4e0f-92ab-db3ef0c2a694");
            /// <summary>
            /// тип "Заголовок ПрИ" 
            /// </summary>
            public static Guid PreliminaryTestPlanHeader = new Guid("b782fe4b-a504-41c8-bfa3-f686a88632ea");
            /// <summary>
            /// тип "Заголовок РД" 
            /// </summary>
            public static Guid DocDevelopPlanHeader = new Guid("5afe8075-8758-4df7-8797-f96ef6b5ccb5");
            /// <summary>
            /// тип "Заголовок ПИ" 
            /// </summary>
            public static Guid PeriodicTestPlanHeader = new Guid("3172e728-8ae9-4949-ab5a-28ac821dcca0");
            /// <summary>
            /// параметр "Название заголовка" 
            /// </summary>
            public static Guid WorkName = new Guid("acb70580-7c8b-48c0-b718-7927dde85e75");
            /// <summary>
            /// параметр "Номер заголовка" 
            /// </summary>
            public static Guid WorkNumber = new Guid("f4895b0b-9576-4d71-ab3c-e8125795d602");
        }

    }
    
    /// <summary>
    /// Guid-ы справочника "Исполнение заданий"
    /// </summary>
    public static class TaskReferenceGuids
    {
        /// <summary>
        /// guid справочника
        /// </summary>
        public static Guid guid = new Guid("3c36ce26-d541-45c0-8351-86fcd73cda04");
        
        /// <summary>
        /// guid типа "Исполнение заданий"
        /// </summary>
        public static Guid taskType = new Guid("3f135d72-3874-4e16-9742-44f0e69b2a7e");

        /// <summary>
        /// связь "Подразделение" (исполнителя)
        /// </summary>
        public static class ExecutorDepartmentLink
        {
            /// <summary>
            /// guid связи
            /// </summary>
            public static Guid guid = new Guid("b2713ac2-f533-498c-a1ac-f39a3f36a52e");
        }

    }

    public static class GroupsAndUsersReferenceGuids
    {
        /// <summary>
        /// параметр "Наименование"
        /// </summary>
        public static Guid Name = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
        /// <summary>
        /// параметр "Краткое имя подразделения"
        /// </summary>
        public static Guid ShortName = new Guid("d834e736-141a-42fa-b2d8-c14f58505686");
        /// <summary>
        /// параметр "Краткое имя пользователя"
        /// </summary>
        public static Guid UserShortName = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");     
        /// <summary>
        /// параметр "Код подразделения"
        /// </summary>
        public static Guid Code = new Guid("6d58df61-2a43-45b2-b8f6-34f3507107fd");
                /// <summary>
/*        /// параметр "Название подразделения"
        /// </summary>
        public static Guid UserOrDepartmentNameGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");*/
    }

    /// <summary>
    /// Должность лица в блоке подписи
    /// </summary>
    public enum SignBlockPost
    {
        /// <summary>
        /// Генеральный директор
        /// </summary>
        GeneralDirector = 0,
        /// <summary>
        /// Главный конструктор
        /// </summary>
        ChiefDesigner = 1,
        /// <summary>
        /// Главный инженер
        /// </summary>
        ChiefEngineer = 2,
        /// <summary>
        /// Зам.генерального директора
        /// </summary>
        DeputyGeneralDirector = 3,
        /// <summary>
        /// Начальник ВП МО
        /// </summary>
        VPMO = 4
    }

    /// <summary>
    /// Тип блока подписи
    /// </summary>
    public enum SignBlockType
    {
        /// <summary>
        /// Согласовано
        /// </summary>
        Agreed = 0,
        /// <summary>
        /// Утверждено
        /// </summary>
        Confirmed = 1
    }

}
