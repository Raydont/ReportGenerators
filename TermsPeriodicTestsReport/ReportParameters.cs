using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;

namespace TermsPeriodicTestsReport
{
    public class ReportParameters
    {
        public NomenclatureObject OrderName;          //Наименование 
                                                   //   public DateTime TermFollowingTest;         //Срок следующего периодического испытания
                                                   //  public bool DefaultPeriodicTests;          //Срок следущего испытания по умолчанию (через 3 года)
        public DateTime TermDeliveryOrder;         // Срок сдачи заказа


    }
}
