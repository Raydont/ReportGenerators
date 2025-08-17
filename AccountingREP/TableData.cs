using System;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.References;

namespace AccountionREPReport
{
    internal class TableData
    {
        public string НомерСзноЗпзп;
        public string КодОКПД2;
        public string НаименованиеРЭП;
        public int Количество;
        public double СтоимостьРуб;
        public bool ОтечественнаяРЭП;
        public string НомерПлатежногоПоручения;
        public DateTime ДатаПлатежногоПоручения;
    }

    internal class TableDataTotal
    {
        public string КодОКПД2;
        public int Количество;
        public double СтоимостьРуб;
        public bool ОтечественнаяРЭП;
        public string ПлатежноеПоручение;

        public TableDataTotal()
        {
        }

        public TableDataTotal(List<TableDataTotal> сверткаПоОКПД2иПП)
		{
            КодОКПД2 = сверткаПоОКПД2иПП.First().КодОКПД2;
            ОтечественнаяРЭП = сверткаПоОКПД2иПП.First().ОтечественнаяРЭП;
            ПлатежноеПоручение = сверткаПоОКПД2иПП.First().ПлатежноеПоручение;
            Количество = сверткаПоОКПД2иПП.Sum(x => x.Количество);
            СтоимостьРуб = сверткаПоОКПД2иПП.Sum(x => x.СтоимостьРуб);           
        }

        public string Ключ()
		{
            return КодОКПД2 + " " + ПлатежноеПоручение + " " + ОтечественнаяРЭП;
        }
            
    }


}