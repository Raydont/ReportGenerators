using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalSheetReport
{
	public static class Guids
	{
		public static class ПроектыДокументов
		{
			public static readonly Guid Id = new Guid("ff45096e-efdc-44d1-a6a9-cd7ecb9dbe8d");

			public static class Параметры
			{
				public static readonly Guid РегистрационныйНомер = new Guid("f25f6039-a77b-4388-b67d-45d4c7ed6d50");
				public static readonly Guid ДатаРегистрации = new Guid("f3090b33-c9df-49a2-aa1e-a8ab12ac1632");
				public static readonly Guid ВидДокумента = new Guid("999f3566-4d9a-4d48-9802-f6c27dcc64bb");
				public static readonly Guid СрокСогласования = new Guid("3be932e4-f944-4d87-be99-c546c0474847");
				public static readonly Guid Закрыто = new Guid("35b2695b-cb3e-4e5b-9ae1-94d467932e8a");
				public static readonly Guid НаименованиеКраткоеСодержание = new Guid("49fdd272-c3c7-4b16-9c89-427176966f1d");
			}

			public static class Связи
			{
				public static readonly Guid РешенияОСогласовании = new Guid("6b99734c-3958-4472-8c93-1c603d14d016");
				public static readonly Guid СогласующиеЛица = new Guid("f3f3dda7-2eaa-4977-9a21-f6c555e9b739");
				public static readonly Guid ФайлыПроектаДокумента = new Guid("80fb4efe-299f-4448-96ff-b8bf369d5190");
			}
			
		}

		public static class РешенияОСогласовании
		{
			public static readonly Guid Id = new Guid("96b35376-bbd5-48d8-bb47-04fec6671c9f");

			public static class Параметры
			{
				public static readonly Guid СрокСогласования = new Guid("4add0f53-b956-4680-bce0-068d399d8e17");
				public static readonly Guid ДатаСогласования = new Guid("7ab41f21-8a37-48c2-9bcb-f3628c6ccb6e");
				public static readonly Guid ВерсияФайла = new Guid("e771f687-eb3e-4cfb-8db0-7da47207b93c");
			}

			public static class Связи
			{
				public static readonly Guid СогласующееЛицо = new Guid("539d5255-68a8-415d-95a8-3e88dd81d9e3");
			}

			public static class СпискиОбъектов
			{
				public static class ЗамечанияПредложения
				{
					public static readonly Guid Id = new Guid("9272b921-5c74-493d-86aa-170f4c9a315c");
					public static class Параметры
					{
						public static readonly Guid ДатаРешения = new Guid("5d4d3358-fcac-425b-88f2-aa50a4c68570");
						public static readonly Guid ЗамечанияПредложения = new Guid("f52dbee6-0602-41d0-9f59-b03c51ecfdc6");
						public static readonly Guid ВерсияФайла = new Guid("cc196260-a830-4333-8c17-6772fa2e51f0");
						public static readonly Guid СтадияРешения = new Guid("4fb9d9a5-9d22-46ff-83f0-b2f1c205628a");
					}
				}
				
			}
		}
	}
}
