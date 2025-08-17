using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CooperationReport
{
    public class ReportParameters
    {
       
        public bool MonthReport = false;
        public bool DateReport = false;
        public bool ExpenseData = true;
        public bool SumOnDevices = true;
        public int Month;
        public string MonthString;
        public int Year;
        public DateTime Date;
        public bool PeriodReport;
        public DateTime BeginDate;
        public DateTime EndDate;
    }
}
