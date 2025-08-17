using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentSendingSummaryReport
{
    public static class TFDClass
    {
        // Guid справочника РКК
        public static readonly Guid regCardsGuid = new Guid("80831DC7-9740-437C-B461-8151C9F85E59");
        // Guid типа "Входящий документ"
        public static readonly Guid inputDocGuid = new Guid("8784c913-ae09-4bce-b27c-0522f6587223");
        // Guid типа "Приказ по предприятию"
        public static readonly Guid orderDocGuid = new Guid("ac00966a-3fc5-497b-b6ed-7fd7c49f434e");
        // Guid типа "Распоряжение по предприятию"
        public static readonly Guid raspDocGuid = new Guid("6608f368-9d7e-45b8-90d2-a9e4731a2545");
        // Guid параметра "Регистрационный номер"
        public static readonly Guid regNumGuid = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
        // Guid параметра "От"
        public static readonly Guid regDateGuid = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
        // Guid параметра "Отправитель"
        public static readonly Guid organisationGuid = new Guid("73198485-e9f9-4dd4-98cd-ecf097f30144");
        // Guid параметра "Исходящий номер"
        public static readonly Guid incomingIndexGuid = new Guid("b430a493-1a7f-45b7-bf44-1d44147a6b2b");
        // Guid параметра "Дата исходящего"
        public static readonly Guid incomingDateGuid = new Guid("d03b88b0-9cc0-4dc2-9262-440f43c91863");
        // Guid параметра "Краткое содержание" типа "Входящий документ"
        public static readonly Guid summaryIndex = new Guid("99b568a7-b92a-4b23-95ec-b8b4947c227a");
        // Guid параметра "Город"
        public static readonly Guid cityGuid = new Guid("4018ca39-8a30-448b-80cc-f5a1f8ac4904");
        // Guid параметра "Содержание"
        public static readonly Guid contentGuid = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
        // Guid параметра "Подписал / Кому"
        public static readonly Guid RCUsersLinkGuid = new Guid("a3746fbd-7bda-4326-8e4b-f6bbefc2ecce");

        // Guid справочника "Списки рассылки"
        public static readonly Guid distributionListGuid = new Guid("202e89ac-2607-464d-a1f5-27cc89675b0c");
        // Guid списка "Объекты рассылки"
        public static readonly Guid distributionObjectsListGuid = new Guid("da40b3a9-ac53-4012-9f65-961e50e3355f");
        // Guid типа "Объект рассылки"
        public static readonly Guid distributionObjectClassGuid = new Guid("6e87454b-f81f-4d8b-a16e-64a673da4b65");
        // Guid параметра "Комментарий" объекта рассылки
        public static readonly Guid distributionObjectCommentGuid = new Guid("7084d5de-83b5-4689-b4a3-88bb80a8ee90");
        // Guid параметра "Срок" объекта рассылки
        public static readonly Guid distributionObjectDeadlineGuid = new Guid("dfe31e86-f563-423b-9939-dd767c2edf37");
        // Guid связи "Получатели" объекта рассылки
        public static readonly Guid distributionObjectReceiversLinkGuid = new Guid("5e02e743-dd63-4674-82fc-ace67c4272ce");
        // Guid параметра "Связанная РКК"
        public static readonly Guid connectedRKKGuid = new Guid("799ee12f-68fa-4419-9383-b5c546909166");
        // Guid связи "Документы" типа "Список рассылки"
        public static readonly Guid distributionListDocumentsLinkGuid = new Guid("291b0cde-eec5-41f1-a603-9fee6ff6522f");
        // Guid параметра "Кто разослал" типа "Список рассылки"
        public static readonly Guid senderDocumentGuid = new Guid("08329b1a-6877-46a4-8155-2f128c854847");

        // Guid справочника "Группы и пользователи"
        public static readonly Guid userReferenceGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        // Guid типа "Производственное подразделение"
        public static readonly Guid departmentClassGuid = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");
        // Guid параметра Наименование справочника Группы и пользователи
        public static readonly Guid usersNameGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
        // Guid подразделения "Канцелярия предприятия"
        public static readonly Guid chancelloryGuid = new Guid("ebf64024-69dc-4119-84eb-d4c283e5dad9");
        // Новый Guid подразделения "Канцелярия предприятия"
        public static readonly Guid chancellory920Guid = new Guid("39bfdc6b-6526-478e-82cc-95b559f72e45");
        
    }
}
