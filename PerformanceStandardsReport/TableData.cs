
using TFlex.DOCs.Model.References;

namespace PerformanceStandardsReport
{
    class TableData
    {
        public string FIO;                          //ФИО
        public string Profession;                   //Профессия
        public string NumberTab;                    //Табельный номер
        public string Category;                     //Разряд
        public double FulfilledHours;               //Отработанное время
        public string Payment;                      //Оплата труда
        public double OverTimeFulfilledHours;       //Отработанное время (Сверхурочно)
        public double DaysOffFulfilledHours;        //Отработанное время (Выходные)
        public double ActualFulFilledHours;         //Фактически отработанное время    
        public double DevelopedStandartHours;       //Выработано нормо-часов
        public bool PartTimeWorker;                 //Совместитель
        public bool Pupil;                          //Ученик      

        public ReferenceObject Worker;              //Рабочий
        
    }
}
