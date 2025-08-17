using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.Reporting;

namespace SearchObjectsNotLinkFilesReport
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

        /// <summary>
        /// Определяет редактор параметров отчета ("Параметры шаблона" в контекстном меню Отчета)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="context"></param>
        /// <param name="owner"></param>
        /// <returns></returns>
        public bool EditTemplateProperties(ref byte[] data, IReportGenerationContext context, System.Windows.Forms.IWin32Window owner)
        {
            return false;
        }


        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {          

            try
            {

                var report = new MakeReport();
                report.Make(context);

                //string str = @" taskkill /IM TFlex.DOCs.FilePreview.Provider.exe /F";
                //ExecuteCommand(str);

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                //if (document != null)
                // document.CloseDocument();
            }
        }
      
    }

    // ТЕКСТ МАКРОСА ===================================

    public delegate void LogDelegate(string line);
    }
