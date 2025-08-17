using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccountingREPReport
{
	public static class Guids
	{
		public static class РКК
		{
			public static readonly Guid Id = new Guid("80831dc7-9740-437c-b461-8151c9f85e59");

			public static class Параметры
			{
				public static readonly Guid РегистрационныйНомер = new Guid("ac452c8a-b17a-4753-be3f-195d76480d52");
				public static readonly Guid ДатаУтверждения = new Guid("fe8e8c1e-91f6-4e82-85ff-f12274c5aa20");
				public static readonly Guid НомерПлатежногоПоручения = new Guid("82987a91-04b5-45e8-a3a5-925cef8a1c46");
				public static readonly Guid ДатаПлатежногоПоручения = new Guid("1c8ab3ba-e5e9-41d2-8cfe-100ca505c59c");
			}

			public static class Типы
			{
				public static readonly Guid СзноЗпзп = new Guid("2a3e0993-8cdf-4ebb-bf13-0c187dcdbc9e");
			}

			public static class Стадии
			{
				public static readonly Guid Утверждено = new Guid("a5ea2e1c-d441-42fd-8f92-49840351d6c1");
				public static readonly Guid ВАрхивеСзно = new Guid("554f7f22-9c5b-4f5f-84e4-068360678ea7");
			}

			public static class СпискиОбъектов
			{
				public static class РэпДляВнутреннегоПотребления
				{
					public static readonly Guid Id = new Guid("c2b6bfdd-d4df-4a63-b61b-ac7fed4a9e70");

					public static class Параметры
					{
						public static readonly Guid КодОКПД2 = new Guid("ac1645f5-a3bf-4467-b0bb-307b7c399058");
						public static readonly Guid Количество = new Guid("ca3c2136-ab0c-4b96-8c89-8d87a3740b02");
						public static readonly Guid НаименованиеРЭП = new Guid("6a8b40f0-99ef-4a63-899c-3d7da384b2e5");
						public static readonly Guid ОтечественнаяРЭП = new Guid("978f9e28-1a43-470c-b784-8cf9eaf8a979");
						public static readonly Guid СтоимостьРуб = new Guid("d47b6bce-1be4-4e7f-ac48-f1fd02811eb9");
					}
				}
				
			}
			
		}
	}
}
