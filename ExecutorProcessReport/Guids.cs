using System;

namespace ExecutorProcessReport
{
    public class Guids
    {
        public static class ДействияПроцессов
        {
            public static readonly Guid Id = new Guid("60e48c95-9d85-4a04-8bed-795bc65b61a4");

            public static class Поля
            {
                public static readonly Guid ИдентификаторСостояния = new Guid("9045a6ad-b3d6-4420-967a-29e7b69289f9");
                public static readonly Guid ПриянтоеРешение = new Guid("eb8e57a2-426c-4d81-beff-beaa8f42d125");
                public static readonly Guid ВремяПринятияРешения = new Guid("e726e796-84e3-4770-a835-123453c2533b");
            }
            public static class СписокОбъектов
            {
                public static class Исполнители
                {
                    public static readonly Guid Id = new Guid("3f93b7c3-fdf2-4411-b342-dba16f927576");
                    public static class Связи
                    {
                        public static readonly Guid Пользователь = new Guid("07d428d2-a222-4b2f-b4b9-a936b437cc80");
                    }

                    public static class СпискиОбъектов
                    {
                        public static class Комментарии
                        {
                            public static readonly Guid Id = new Guid("bb1db87b-54d1-4f13-bec8-efdc612b6c0d");
                            public static class Поля
                            {
                                public static readonly Guid Текст = new Guid("339f97b6-742e-47f7-b7ee-7becc08143c8");
                            }
                        }
                    }
                }
            }
            public static class Типы
            {
                public static readonly Guid Состояния = new Guid("5505cf81-c0c3-46c1-9116-38b7a1b2d441");
            }
        }
        public static class ТекущиеДействияПроцессов
        {
            public static readonly Guid Id = new Guid("f6cc44ca-886b-404f-89e2-593e1754d278");
            public static class Поля
            {
                public static readonly Guid ИдентификаторЭтапаПроцесса = new Guid("a3c276dc-ca53-4ada-9e36-fb2c49efe8ae");
            }
        }
        public static class Процессы
        {
            public static readonly Guid Id = new Guid("e0c70b3c-bee1-4321-87b0-c44f3d8b5f68"); 
            public static class Связи
            {
                public static readonly Guid Действия = new Guid("56cf584b-7206-47d1-9aa8-81faf7484dc8");
                public static readonly Guid ТекукщиеДействия = new Guid("721f2a2a-3dc2-4578-b3c5-4224f8d8a1e0");
                public static readonly Guid Процедура = new Guid("44c3a825-ef4a-4ecb-a291-1048db42d5cb");
                public static readonly Guid ЗапущенныеОбъекты = new Guid("cf72b950-52d9-4089-a07b-efad09ed613c");
            }
            public static class Поля
            {
                public static readonly Guid Статус = new Guid("cc06378b-4987-4322-a69d-6bcbb295cdf7");
            }
        }    
        public static class Процедуры
        {
            public static readonly Guid id = new Guid("61d922d0-4d60-49f1-a6a0-9b6d22d3d8b3");
            public static class Поля
            {
                public static readonly Guid ДоступНаЗапуск = new Guid("cdee3402-fc51-48ae-855b-4505b8391cd5");
                public static readonly Guid Наименование = new Guid("5e99d9c1-4d74-4d62-8806-aea35c4c792e");
            }
            public static class Связи
            {
                public static readonly Guid ПравоНаЗапускБп = new Guid("84579eed-d8ea-404f-a8b7-a1b63c3db722");
            }
            public static class СписокОбъектов
            {
                public static class Состояния
                {
                    public static readonly Guid Id = new Guid("de0cdb34-aee3-459a-bf97-74094270cef9");
                    public static class Поля
                    {
                        public static readonly Guid ВремяНачала = new Guid("2b2321d5-c4bc-4a8a-8820-7c31d1b2564a");
                        public static readonly Guid ВремяЗавершения = new Guid("eef93b06-8ec7-472f-acfc-cc1b2ee7a12f");
                        public static readonly Guid Наименование = new Guid("9184d80e-7707-4c62-aed2-f2fdda779b57");
                    }
                    public static class Типы
                    {
                        public static readonly Guid Работа = new Guid("bd103ee7-8824-41f7-9a78-5d96f0cbf9ff");
                        public static readonly Guid Согласование = new Guid("f4208140-87d9-4adb-8e98-90cb50d8e774");
                        public static readonly Guid Делитель = new Guid("e5f7cc95-06d3-4156-9bfd-3ca282798d68");
                        public static readonly Guid Сумматор = new Guid("74e0d49d-14d8-4955-b9b8-20c0a00354ba");
                    }   
                    public static class СписокОбъектов
                    {
                        public static class Переходы
                        {
                            public static readonly Guid Id = new Guid("4f26d906-cdce-4cdd-aa1d-e621cfa8b6f5");
                            public static class Типы
                            {
                                public static readonly Guid Переход = new Guid("0f200b3e-07f6-44d2-8c55-72209a623af8");
                            }
                        }
                    }
                    public static class Связи
                    {
                        public static readonly Guid Исполнители = new Guid("702580d5-27b4-4e45-8eec-8949e3db5a3d");
                    }
                }
                public static class Справочники
                {
                    public static readonly Guid Id = new Guid("e54a1ab2-eb94-4b79-be8d-3722cb1d5e45");
                    public static class Поля
                    {
                        public static readonly Guid Справочник = new Guid("71fb1911-694f-4106-8641-7b27516d2c3b");
                        public static readonly Guid Условие = new Guid("6e40f04c-ca99-43d3-b353-0a1e240abbe8");
                    }
                }
            }      
        }
        public static class ГруппыИПользователи
        {
            public static readonly Guid Id = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
            public static class Типы
            {
                public static readonly Guid ПроизводственнаяОрганизация = new Guid("dd7c723e-3e7f-4198-861b-97cbe727f678");
                public static readonly Guid ПроизводственноеПодразделение = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");
                public static readonly Guid Бюро = new Guid("c71cc322-7cc0-4822-a625-d72f9428a40d");
                public static readonly Guid Участок = new Guid("6f16140a-6baa-4a11-b612-597df1098094");
                public static readonly Guid Цех = new Guid("05aca34a-0e80-4e38-8367-7d7bcbf66e69");
                public static readonly Guid Лаборатория = new Guid("2c1edef0-8f15-4ffe-952d-98adeb74a4aa");
                public static readonly Guid Отдел = new Guid("1dad31c3-e89e-403f-9454-ee70fc1347c7");
                public static readonly Guid Бригада = new Guid("e6819f01-cf5f-4a91-8e3a-f56f612bcece");
                public static readonly Guid Подразделение = new Guid("9ed09a3b-8a54-403a-81a8-e9a182231da2");
                public static readonly Guid Пользователь = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");
                public static readonly Guid Администратор = new Guid("0bc8a939-2914-4e79-9621-3e03c39b9c31");
                public static readonly Guid Сотрудник = new Guid("41ca5dde-4b17-4144-b12f-447eacefd7e8");
            }
            public static class Поля
            {
                public static readonly Guid Наименование = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
                public static readonly Guid СокращенноеНаименование = new Guid("d834e736-141a-42fa-b2d8-c14f58505686");
                public static readonly Guid Номер = new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420");
                public static readonly Guid UserName = new Guid("0f3b4b68-72d6-4376-8258-f41a512cac73");
                public static readonly Guid UserSurename = new Guid("76a97c36-f2d6-49ad-abb0-f5fdc91c071b");
                public static readonly Guid UserPatronymic = new Guid("6d5a1ce6-59fc-4d1c-9b1f-67980695edc0");
                public static readonly Guid КороткоеИмя = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");
                public static readonly Guid Код = new Guid("6d58df61-2a43-45b2-b8f6-34f3507107fd");
            }
        }
        public static class Ркк
        {
            internal static readonly Guid Id = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");
            public static class Типы
            {
                public static readonly Guid НарядЗаказ = new Guid("b0c914bd-a920-40e8-bebd-a538c0e54436");
                public static readonly Guid Сзно = new Guid("2a3e0993-8cdf-4ebb-bf13-0c187dcdbc9e");
                public static readonly Guid Зпзп = new Guid("e35cbc95-2bc4-442d-a530-73075eaf4ed9");
                public static readonly Guid РегистрационноКонтрольныеКарточки = new Guid("ec785091-5483-4d5a-9cf5-188b15cd6cac");
            }
            public static class Поля
            {
                public static readonly Guid РегистрационныйНомер = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
                public static readonly Guid От = new Guid("9930335a-8703-47a1-8dc8-49e17a959d20");
                public static readonly Guid Заказчик = new Guid("5f506d0e-503f-4b2b-932b-bf436935ee4d");
                public static readonly Guid ШифрЗаказа = new Guid("050b0e4f-ebc1-44cf-8c83-b76e21057399");
                public static readonly Guid СуммаЗпзп = new Guid("71841bbe-ab66-49b7-82e2-c219e5d9fc58");
                public static readonly Guid СодержаниеЗпзп = new Guid("29c9b844-7825-4597-86ca-f6d789e4b2d6");
                public static readonly Guid КонтрагентЗпзп = new Guid("5ec1cf13-54ed-45c4-84ae-a6e93e2ed76c");
                public static readonly Guid СуммаСзно = new Guid("f962d417-de7d-4723-ba9c-ed736271ab22");
                public static readonly Guid СодержаниеСзно = new Guid("b85553c9-2e2d-4f27-96fa-9ebb8e365747");
                public static readonly Guid КонтрагентСзно = new Guid("ed34fb5e-d8a4-4561-8f81-480edb85d11b");
            }
            public static class Связи
            {
                public static readonly Guid Подразделение = new Guid("e316a9f3-08f5-4cf8-9e78-ba82b70b204a");
            }
        }
    }
}
