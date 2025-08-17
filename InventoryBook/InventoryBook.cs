using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex;
using System.Windows.Forms;
using System.Drawing;

namespace InventoryBook
{
    public class ReportGenerator : IReportGenerator
    {
        /// <summary>
        /// Расширение файла отчета
        /// </summary>
        public string DefaultReportFileExtension
        {
            get
            {
                return ".grb";
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
            ReportForm rForm = new ReportForm();
            rForm.ShowDialog();
            if (rForm.DialogResult == DialogResult.OK)
            {
                try
                {
                    var reader = new DataReader(rForm.reportParameters);
                    var docParams = reader.ReadData();
                    if (docParams.Count > 0)
                    {
                        // Инициализация CAD APi
                        var cadDocument = GetCadDocument(context);
                        if (cadDocument == null) MessageBox.Show("Нет документа");
                        var maker = new ReportMaker(cadDocument);
                        maker.FillDocument(docParams);
                        cadDocument.Save();
                        cadDocument.Close();
                        System.Diagnostics.Process.Start(context.ReportFilePath);
                    }
                    else
                        MessageBox.Show("За выбранный период документов не зарегистрировано");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Ошибка!");
                }
            }
        }

        /// <summary>
        /// Инициализация CAD APi
        /// </summary>
        public TFlex.Model.Document GetCadDocument(IReportGenerationContext context)
        {
            var setup = new ApplicationSessionSetup
            {
                Enable3D = true,
                ReadOnly = false,
                PromptToSaveModifiedDocuments = false,
                EnableMacros = true,
                EnableDOCs = true,
                DOCsAPIVersion = ApplicationSessionSetup.DOCsVersion.Version13,
                ProtectionLicense = ApplicationSessionSetup.License.TFlexDOCs
            };
            bool result = TFlex.Application.InitSession(setup);

            if (!result)
                throw new InvalidOperationException("Ошибка подключения к Cad Api");

            TFlex.Application.FileLinksAutoRefresh = TFlex.Application.FileLinksRefreshMode.DoNotRefresh;     //Инициализация API CAD              
            context.CopyTemplateFile();
            return TFlex.Application.OpenDocument(context.ReportFilePath);
        }
    }

    public class DocumentParameters
    {
        public string invnum;                             // Инвентарный номер (текущий)
        public string incomingDate;                       // Дата поступления
        public string denotation;                         // Обозначение
        public string sheetsCount;                        // Количество листов
        public string format;                             // Формат документа
        public string name;                               // Наименование
        public string department;                         // Подразделение
        public string comment;                            // Примечание
    }

    public class ReportParameters
    {
        public bool iNumber;                              // Отбор по дате либо инвентарному номеру
        public DateTime startDate;                        // Начальная дата
        public DateTime endDate;                          // Конечная дата
        public int startNumber;                           // Начальный номер
        public int endNumber;                             // Конечный номер
        public TextFormatter tForm;                       // Форматировщик текста
    }
}