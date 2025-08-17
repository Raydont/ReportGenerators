using System.Collections.Generic;
using TFlex.DOCs.Model.References;

namespace ChargesReport
{
    internal class TableData
    {
        public string FIO;                          //ФИО
        public string TimeBoardNumber;                //Табельный номер
        public string Profession;                   //Профессия
        public string Category;                     //Разряд
        public double WageRate;                     //Тарифная ставка (оклад)
        public double FulfilledHours;               //Отработанное время
        public double OvertimeFirst2Hours;          //Сверхурочные первые два часа
        public double OvertimeNextHours;            //Сверхурочные последующие часы
        public double SverkhurochnyeGOZ;            //Сверхурочные ГОЗ
        //public double ActualFulfilledHours;         //Фатически отработанное время
        public double SurChargesFulfilledHours;     //Отработанное время (Доплаты)
        public double ActualFulfilledHours;       //Отработанное время (Сверхурочно)
        public double DaysOffFulfilledHours;        //Отработанное время (Выходные)
        public double PercentForSurcharges;         //Процент занятости для доплат
        public double SalaryUnderTarif;             //Сумма заработной платы по тарифу
        public double Surcharges;                   //Надбавки
        public double PersentForExtracharges;       //Процент для надбавок
        public double PercentBasicAwards;           //Основная премия
        public double PercentAdditionalAwards;      //Дополнительная премия
        public double PercentAdditionalAwards2;     //Дополнительная премия для професси по совместительству
        public double SumAwards;                    //Сумма премии
        public string Comment;                      //Примечание
        public double SumSalaryPlanOverfulfillment; //Сумма зарплаты за перевыполнение плана
        public double NormoHours; //отработанные нормочасы
        public ReferenceObject Worker;              //Рабочий
        public bool Pupil;                          //Ученик
        public bool Dismissed;                      //Уволен
        public double InterdigitDifferenceSum;           //Межразрядная разница
        public bool CalcNormByOtherProf;            //Выячисление сверхномр по остальным профессиям
        public bool PartTimeWorker;                 //Совместитель
                                                    // public List<double> CountHoursWorks = new List<double>();  //Список по количеству часов каждой работы
                                                    // public List<string> ProfessionWork = new List<string>(); //профессия (станочники или остальные рабочие, в соответствии с работой)
        public List<ВыполняемаяРабота> ВыполняемыеРаботы = new List<ВыполняемаяРабота>();
    }
    class ВыполняемаяРабота
    {
        public bool ОсновнаяПрофессия;
        public bool ПрофессияПоСовместительству;
        public double ОтработанныеЧасы;
        public double ПроцентОсновнойПремии;
        public double СуммаЗаработнойПлаты;
        public ReferenceObject СтоимостьЧасовСверхНорм;

        public ВыполняемаяРабота(ReferenceObject стоимостьЧасовСверхНорм, bool основнаяПрофессия, bool профессияПоСовместительству, double отработанныеЧасы, double процентОсновнойПремии, double суммаЗаработнойПлаты)
        {
            СтоимостьЧасовСверхНорм = стоимостьЧасовСверхНорм;
            ОсновнаяПрофессия = основнаяПрофессия;
            ПрофессияПоСовместительству = профессияПоСовместительству;
            ОтработанныеЧасы = отработанныеЧасы;
            ПроцентОсновнойПремии = процентОсновнойПремии;
            СуммаЗаработнойПлаты = суммаЗаработнойПлаты;
        }
    }
}