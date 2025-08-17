using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TransitTimeProcess
{
    public class Guids
    {
        // связь процессов и справочника действия
        public static Guid LinkProcessToAction = new Guid("56cf584b-7206-47d1-9aa8-81faf7484dc8");
        //Исполнители действия процесса
        public static Guid MakersProcessActions = new Guid("3f93b7c3-fdf2-4411-b342-dba16f927576");
        //Список объектов подписи блока установка подписей
        public static Guid SignatureListObject = new Guid("feedf347-d014-4ded-93fd-cf0c94e85ede");
        //Тип объекта блока Установка подписи
        public static Guid PasteSignatureType = new Guid("ecf6b2bf-a3dc-430e-b189-d6e242529c02");
        //Связь исполнители
        public static Guid LinkMakerToUser = new Guid("07d428d2-a222-4b2f-b4b9-a936b437cc80");
        //Справочник процедуры
        public static Guid ProcedureReference = new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3");
        //параметры состояния процесса
        public static Guid StateParameters = new Guid("9cb2da09-f671-436d-9797-2e5d0eaf868e");
        //Идентификатор состояния
        public static Guid IdenticatorState = new Guid("9045a6ad-b3d6-4420-967a-29e7b69289f9");
        //Список объектов Состояния
        public static Guid ListObjectsSate = new Guid("de0cdb34-aee3-459a-bf97-74094270cef9");
        //Тип подписи
        public static Guid TypeSignature = new Guid("39c141fd-9784-45d6-b191-b31e202256d5");
        //Тип состояния Установка подписи
        public static Guid TypeSigns = new Guid("ecf6b2bf-a3dc-430e-b189-d6e242529c02");
        //Гуид блока исполнителя
        public static Guid ExecuterStage = new Guid("307590d1-e516-472d-bbca-762151c98a7b");
        //Тип Состояние процесса
        public static Guid StateActionType = new Guid("5505cf81-c0c3-46c1-9116-38b7a1b2d441");
        //Принятое решения
        public static Guid SolutionName = new Guid("eb8e57a2-426c-4d81-beff-beaa8f42d125");
        //Время принятия решения
        public static Guid SolutionEndTime = new Guid("e726e796-84e3-4770-a835-123453c2533b");
        //Плановая дительность процесса
        public static Guid DurationPlanState = new Guid("983552bd-080d-4285-b48b-255dbeaff1c3");
        //Короткое имя пользователя
        public static Guid UserShortNameGuid = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
        //Абстрактный тип пользователь
        public static Guid UserTypeGuid = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");
        //Гуид списка объектов комментарии исполнителей
        public static readonly Guid CommentsMakersGuid = new Guid("bb1db87b-54d1-4f13-bec8-efdc612b6c0d");
        //Текст комментария
        public static readonly Guid TextCommentGuid = new Guid("339f97b6-742e-47f7-b7ee-7becc08143c8");
        //Тип состояния запуск подпроцесса
        public static readonly Guid TypeStateProcedureGuid = new Guid("48e4864f-717e-4479-84d9-14e3923e0411");
        //Тип состояния Работа
        public static readonly Guid TypeStateWorkGuid = new Guid("bd103ee7-8824-41f7-9a78-5d96f0cbf9ff");
        //Тип состояния Согласование
        public static readonly Guid TypeStateCoordinationGuid = new Guid("f4208140-87d9-4adb-8e98-90cb50d8e774");
        //Тип состояния Сумматор
        public static readonly Guid TypeStateSummatorGuid = new Guid("74e0d49d-14d8-4955-b9b8-20c0a00354ba");
        //Тип состояния Делитель
        public static readonly Guid TypeStateDelitelGuid = new Guid("e5f7cc95-06d3-4156-9bfd-3ca282798d68");
        //Переход для состояния
        public static readonly Guid ProcessTransitions = new Guid("4f26d906-cdce-4cdd-aa1d-e621cfa8b6f5");
        //Тип переход
        public static readonly Guid TypeTransitionGuid = new Guid("0f200b3e-07f6-44d2-8c55-72209a623af8");
        //Целевое состояние перехода
        public static readonly Guid TargetProcessTransition = new Guid("4adf5127-edf9-4184-8cab-5fd70dd2160b");
        //Наименование состояния
        public static readonly Guid NameStateTransition = new Guid("a94b5d95-4d79-49b2-b510-45aaa863ede3");

        public static readonly Guid NameState = new Guid("8c2e2c13-9a94-4d2f-a084-17791a312aeb");
        //Исполнители состояния
        public static readonly Guid MakerStateGuid = new Guid("702580d5-27b4-4e45-8eec-8949e3db5a3d");
        //Планируемое время завершения 
        public static readonly Guid PlanEndTimeState = new Guid("35c1e484-2df5-4aea-b721-759abe87b85e");
        //Статус процесса
        public static readonly Guid ProcessStatusGuid = new Guid("cc06378b-4987-4322-a69d-6bcbb295cdf7");
        //Текущие действия процесса
        public static readonly Guid ActiveActionsLinkGuid = new Guid("721f2a2a-3dc2-4578-b3c5-4224f8d8a1e0");
        //Наименование текущего действия 
        public static readonly Guid NameActiveActionGuid = new Guid("60f68ed7-0f6e-4ff2-ab2a-1823f1bbd969");
        //идентификатор состояния этапа процесса для текущего действия
        public static readonly Guid IdentificatorStepGuid = new Guid("a3c276dc-ca53-4ada-9e36-fb2c49efe8ae");
        //Родительский объект действия процессов
        public static readonly Guid ParentActiveActionGuid = new Guid("64eb9bfc-eacb-4d79-b435-8281e5d1e6b7");
        //Время начала
        public static readonly Guid StartTime = new Guid("2b2321d5-c4bc-4a8a-8820-7c31d1b2564a");
        //Время завершения
        public static readonly Guid EndTime = new Guid("eef93b06-8ec7-472f-acfc-cc1b2ee7a12f");

    }
}
