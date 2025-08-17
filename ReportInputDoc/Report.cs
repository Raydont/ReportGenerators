using System.Windows.Forms;
using TFlex.Reporting;

namespace ReportIncomingDoc
{
    public class Report
    {
        //Точка входа для запуск библиотеки из справочника Номенклатура и изделия T-FLEX DOCs
        public void Start(IReportGenerationContext context)
        {
            var generateExcel = new GenerateExcel();
            var formSelection = new FormSelection();
            if (formSelection.ShowDialog() == DialogResult.OK)
            {
                using (var formIndication = new FormIndication())
                {
                    formIndication.Visible = true; //Отображаем форму с логом
                    generateExcel.ExportToExcel(formSelection, formIndication, context);
                }
            }
        }
    }
}