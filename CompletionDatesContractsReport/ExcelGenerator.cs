using System.Windows.Forms;
using TFlex.Reporting;

namespace CompletionDatesContractsReport
{
    public class ExcelGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xls";
            }
        }

        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionForm selectionForm = new SelectionForm();

        public void Generate(IReportGenerationContext context)
        {
            using (new WaitCursorHelper(false))
            {
                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    selectionForm.Init(context);
                }
            }
        }
    }
}
