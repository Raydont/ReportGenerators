using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Stages;

namespace LogisticsControlReport
{
    public static class Guids
    {
        //Справочники
        public static class References
        {
            public static readonly Guid GroupsAndUsers = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");   //Группы и пользователи
            public static readonly Guid OMTS = new Guid("78c9616a-e62b-4221-9112-b61874afd07d");             // Журнал контроля исполнения документов ОМТС
        }

        //Параметры справочника Группы и пользователи
        public static class UsersParameters
        {
            public static readonly Guid ShortName = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");       // Короткое имя
            public static readonly Guid FullName = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");        // Полное имя
        }

        //Типы справочника Группы и пользователи
        public static class UserType
        {
            public static readonly Guid Employee = new Guid("41ca5dde-4b17-4144-b12f-447eacefd7e8");        //Сотрудник
        }

        // Объекты справочника Группы и пользователи
        public static class UserObject
        {
            public static readonly Guid Department914 = new Guid("6a04a9d7-5a8e-4fb7-9996-fd87575e8254");     //Отдел 914
            public static readonly Guid Chief914 = new Guid("0e651676-614a-4de9-bdc7-bae148517972");          //Начальник отдела 914
            public static readonly Guid Store1 = new Guid("998c4cee-7fe1-4651-941e-f154dd10c2ed");            //Склад 1
            public static readonly Guid Store4 = new Guid("2cadc16e-9deb-4cf4-bb4f-8a028b362b4c");            //Склад 4
        }

        //Параметры справочника ОМТС
        public static class OMTSParameters
        {
            public static readonly Guid RegNumber = new Guid("28c0b86d-d591-4274-be44-fe4c407be10f");               //Регистрационный номер
            public static readonly Guid InDocLink = new Guid("be24c108-95a3-4084-81d1-f46919395fa8");               //Связь с входящим дркументов
            public static readonly Guid Contractor = new Guid("31bf7f24-e6a6-47a7-b831-4c962231098b");              //Контрагент
            public static readonly Guid Executer = new Guid("e695e689-6aca-4b65-98ba-c9dffd9a2089");                //Ответственный исполнитель
            public static readonly Guid RegData= new Guid("062dd2ac-612b-4743-8375-b23292fe05a0");                  //Дата регистрации
            public static readonly Guid PlanExecuteDate = new Guid("067eb3b9-c8c9-4f1c-9e1d-c0cdfe0a2dfd");         //Срок исполнения
            public static readonly Guid ExecuteData = new Guid("25a52b65-14ea-4f06-8948-7bc01079e917");             //Дата исполнения
            public static readonly Guid FailsListObjects = new Guid("a74cb0f0-8c52-4e4b-9dd3-99a7d7c25b30");        //Список объектов незачеты
            public static readonly Guid Fail = new Guid("59ee7af3-d294-4163-9e00-e70775d501ef");                    //Незачет
            public static readonly Guid ChangeDateHistoryList = new Guid("0a7fc775-f960-4809-b757-3805f0fc7ca8");   //Список объектов история дат переносов
            public static readonly Guid ChangeDate = new Guid("8bdb0dbd-cf2e-4c6e-9552-cba08388cab6");              //Дата переноса
            public static readonly Guid Content = new Guid("749672a6-a37b-4e79-962c-d40db38d27cc");                 //Краткое содержание
        }

        // Параметры справочника РКК
        public static class RKKParametetrs
        {
            public static readonly Guid RegNumber = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");               //Регистрационный номер
            public static readonly Guid RegData = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");                 //Дата регистрации
        }

        //Стадии справочника ОМТС
        public class StagesOMTS
        {
            public static readonly Guid Open = new Guid("38bbddb0-0726-46e8-9431-d03fcafd1099");
            public static readonly Guid InWork = new Guid("b6bf4233-d775-4b67-8965-1a554b95cc38");
            public static readonly Guid ReWork = new Guid("bfd6885d-2b1a-488f-81c1-e3e253aed083");
            public static readonly Guid Close = new Guid("0b185d13-e282-478b-a75b-5c1676e832e1");
        }
    }

}
