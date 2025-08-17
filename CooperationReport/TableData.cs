using System;
using System.Collections.Generic;

namespace CooperationReport
{
    class TableData
    {
        public string Order;
        public string Device;
        public string IncomeDevice;
        public string Denotation;
        public string IncomeType;
        public double Cost;
        public string Contractor;
        public string Contract;
        public List<IncomeExpend> ArrivalExpenses = new List<IncomeExpend>();
    }

    class IncomeExpend
    {
        public string DocumentNumber;
        public int Count;
        public double Labor800;
        public double Labor850;
        public double Labor750;
        public double SumLabor;
        public double Sum;
        public DateTime Date;

        public IncomeExpend(int count, double labor800, double labor850, double labor750, double sumLabor, double sum, DateTime date, string documentNumber)
        {
            Count = count;
            DocumentNumber = documentNumber;
            Labor800 = Math.Round(labor800,2) ;
            Labor850 = Math.Round(labor850, 2);
            Labor750 = Math.Round(labor750, 2);
            SumLabor = Math.Round(sumLabor, 2);
            Sum = Math.Round(sum, 2); ;
            Date = date;
        }
        public IncomeExpend()
        {
            Count = 0;
            Labor800 = 0;
            Labor850 = 0;
            Labor750 = 0;
            SumLabor = 0;
            Sum = 0;
            Date = DateTime.MinValue;
        }
    }
}
