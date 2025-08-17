using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.References;

namespace ApprovalSheetReport
{
    // данные проекта документа
    internal class TableData
    {        
        public string РегистрационныйНомер;
        public DateTime ДатаРегистрации;
        public bool Закрыто;
        public string ВидДокумента;
        public DateTime СрокCогласования;
        public string НаименованиеКраткоеСодержание;
        public string ПоследняяВерсияФайла;
        public List<РешениеОСогласовании> РешенияОСогласовании;
    }

    // приенятые решения о согласовании проекта документа
    internal class РешениеОСогласовании
	{
        public string СогласующееЛицо;
        public string ВерсияФайла;
        public DateTime ДатаCогласования;
        public DateTime СрокCогласования;
        public string ТекущаяСтадия;
        public List<ЗамечаниеПредложение> ЗамечанияПредложения;        
    }

    // замечения/предложения решений о согласовании проекта документа
    internal class ЗамечаниеПредложение
	{
        public string ВерсияФайла;
        public DateTime ДатаРешения;
        public string ЗамечанияПредложения;
        public string СтадияРешения;
    }
}