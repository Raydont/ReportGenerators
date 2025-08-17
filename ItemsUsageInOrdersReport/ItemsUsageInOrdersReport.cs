using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.Reporting;

namespace ItemsUsageInOrdersReport
{
    // отчет "Ведомость применяемости ПКИ"
    class ItemsUsageInOrdersReport : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xlsx";
            }
        }

        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }

        public void Generate(IReportGenerationContext context)
        {
            if (File.Exists(context.ReportFilePath))
            {
                try
                {
                    File.Delete(context.ReportFilePath);
                }
                catch
                {
                    MessageBox.Show("Закройте файл " + context.ReportFilePath + " перед формированием отчета!");
                    return;
                }
            }

            var dialog = new SearchMainForm(context.ReportFilePath);
            if (dialog.ShowDialog(Form.ActiveForm) != DialogResult.OK)
            {
                return;
            }          
        }
    }
}
