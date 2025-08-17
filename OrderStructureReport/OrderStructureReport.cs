using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.Reporting;
using ReportHelpers;

namespace OrderStructureReport
{
    public class ExcelGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
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
                OrderStructureReport report = new OrderStructureReport();
                report.Make(context);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
    public class OrderStructureReport : IDisposable
    {
        public void Dispose()
        {
        }

        public void Make(IReportGenerationContext context)
        {
            // Получение ID объекта, на который получаем отчет
            int baseDocumentID = Initialize(context);
            if (baseDocumentID == -1) return;

            DatabaseReader reader = new DatabaseReader();
            // Считывание параметров корневого объекта
            ComplexParameters parameters = reader.getComplexParameters(baseDocumentID);
            // Считыввание параметров дерева корневого объекта
            List<DocumentParameters> docParams = reader.getDBData(baseDocumentID);

            bool isCatch = false;
            // Создание копии шаблона
            context.CopyTemplateFile();
            // Открытие нового документа
            Xls xls = new Xls();

            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath);
                ExcelReportGenerator generator = new ExcelReportGenerator();
                // Создание титульного листа
                generator.CreateTitulList(parameters.Name, parameters.Denotation, parameters.Order, xls);
                // Создание классификатора
                generator.CreateClassificator(new string[] {"Покупное", "ЗИП"},xls);
                // Создание заголовка отчета
                generator.CreateHeader(xls);
                // Создание тела отчета
                generator.CreateReport(docParams, xls);
                // Сохранение документа
                xls.Save();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                    xls.Quit(false);
                else
                {
                    xls.Quit(true);
                    // Открытие файла
                    System.Diagnostics.Process.Start(context.ReportFilePath);
                }
            }
        }

        private int Initialize(IReportGenerationContext context)
        {
            int objectID = 0;

            // Получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) objectID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            return objectID;
        }
    }
}
