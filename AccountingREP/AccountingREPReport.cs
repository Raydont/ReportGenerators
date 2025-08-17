using AccountionREPReport;
using ReportHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.Reporting;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using TFlex.DOCs.Model.Stages;

namespace AccountingREPReport
{
    public class ExcelGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".xlsx";
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

        //Создание экземпляра формы выбора параметров формирования отчета
        public SelectionForm selectionForm = new SelectionForm();

        /// <summary>
        /// Основной метод, выполняется при вызове отчета 
        /// </summary>
        /// <param name="context">Контекст отчета, в нем содержится вся информация об отчете, выбранных объектах и т.д...</param>
        public void Generate(IReportGenerationContext context)
        {
            try
            {
                selectionForm.Init(context);

                if (selectionForm.ShowDialog() == DialogResult.OK)
                {
                    AccountingREPReport.MakeReport(context, selectionForm.reportParameters);
                }

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e + "", "Exception!", System.Windows.Forms.MessageBoxButtons.OK,
                                                     System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {

            }
        }
    }

    public delegate void LogDelegate(string line);

    public class AccountingREPReport
    {
        public IndicationForm m_form = new IndicationForm();

        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context, НастройкиОтчета reportParameters)
        {
            AccountingREPReport report = new AccountingREPReport();
            //LogDelegate m_writeToLog;

            report.Make(context, reportParameters);
        }

        public void Make(IReportGenerationContext context, НастройкиОтчета reportParameters)
        {
            var m_writeToLog = new LogDelegate(m_form.WriteToLog);
            context.CopyTemplateFile();    // Создаем копию шаблона

            m_form.TopMost = true;
            m_form.Visible = true;
            m_form.Activate();
            m_writeToLog = new LogDelegate(m_form.WriteToLog);

            m_form.SetStage(IndicationForm.Stages.Initialization);
            m_writeToLog("=== Инициализация ===");

            Xls xls = new Xls();
            bool isCatch = false;

            AccountingREPReport report = new AccountingREPReport();

            m_form.SetStage(IndicationForm.Stages.DataAcquisition);

            List<TableData> reportData = GetDetailData(reportParameters, m_writeToLog, m_form.progressBar);

            if (reportData == null)
            {
                m_form.Close();
                return;
            }

            try
            {
                xls.Application.DisplayAlerts = false;
                //  xls.Open(context.ReportFilePath);

                m_form.SetStage(IndicationForm.Stages.ReportGenerating);
                m_writeToLog("=== Формирование отчета ===");

                //Формирование отчета
                MakeAccountingREPReport(xls, reportData, reportParameters, m_writeToLog, m_form.progressBar);

                m_form.SetStage(IndicationForm.Stages.Done);
                m_writeToLog("=== Завершение работы ===");

                //  xls.Save();
                xls.Application.DisplayAlerts = true;

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e + "", "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                xls.Visible = true;
                //if (isCatch)
                //{
                //   // xls.Quit(false);
                //}
                //else
                //{

                //    xls.Quit(true);
                //    // Открытие файла
                //    Process.Start(context.ReportFilePath);
                //}
            }

            m_form.Dispose();
        }

        private List<TableData> GetDetailData(НастройкиОтчета reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            logDelegate("=== Загрузка данных ===");

            List<TableData> reportData = new List<TableData>();

            var connection = ServerGateway.Connection;

            var referenceRKKInfo = connection.ReferenceCatalog.Find(Guids.РКК.Id);
            var referenceRKK = referenceRKKInfo.CreateReference();
            var classSZNO_ZPZP = referenceRKKInfo.Classes.Find(Guids.РКК.Типы.СзноЗпзп);

            // все СЗНО/ЗПЗП с заполненной таблицей РЭП за указанный период
            var filterSznoZpzp = new TFlex.DOCs.Model.Search.Filter(referenceRKKInfo);
            filterSznoZpzp.Terms.AddTerm(referenceRKK.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classSZNO_ZPZP);
            filterSznoZpzp.Terms.AddTerm(referenceRKK.ParameterGroup[SystemParameterType.Stage], ComparisonOperator.IsOneOf, new List<Stage> { Stage.Find(Guids.РКК.Стадии.Утверждено), Stage.Find(Guids.РКК.Стадии.ВАрхивеСзно) } );
            filterSznoZpzp.Terms.AddTerm("[" + Guids.РКК.Параметры.ДатаУтверждения + "]", ComparisonOperator.GreaterThanOrEqual, reportParameters.НачалоПериода);
            filterSznoZpzp.Terms.AddTerm("[" + Guids.РКК.Параметры.ДатаУтверждения + "]", ComparisonOperator.LessThanOrEqual, reportParameters.КонецПериода);
            filterSznoZpzp.Terms.AddTerm("[" + Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Id + "]", ComparisonOperator.IsNotNull, null);
            List<ReferenceObject> СзноЗпзп = referenceRKK.Find(filterSznoZpzp);

            // формирование детальных данных
            foreach (var оплата in СзноЗпзп)
            {
                var списокРЭП = оплата.GetObjects(Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Id);
                foreach (var рэп in списокРЭП)
                {
                    var позицияРЭП = new TableData();
                    позицияРЭП.НомерСзноЗпзп = оплата[Guids.РКК.Параметры.РегистрационныйНомер].GetString();
                    позицияРЭП.КодОКПД2 = рэп[Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Параметры.КодОКПД2].GetString();
                    позицияРЭП.НаименованиеРЭП = рэп[Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Параметры.НаименованиеРЭП].GetString();
                    позицияРЭП.Количество = рэп[Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Параметры.Количество].GetInt32();
                    позицияРЭП.СтоимостьРуб = рэп[Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Параметры.СтоимостьРуб].GetDouble();
                    позицияРЭП.ОтечественнаяРЭП = рэп[Guids.РКК.СпискиОбъектов.РэпДляВнутреннегоПотребления.Параметры.ОтечественнаяРЭП].GetBoolean();
                    позицияРЭП.НомерПлатежногоПоручения = оплата[Guids.РКК.Параметры.НомерПлатежногоПоручения].GetString();
                    позицияРЭП.ДатаПлатежногоПоручения = оплата[Guids.РКК.Параметры.ДатаПлатежногоПоручения].GetDateTime();

                    reportData.Add(позицияРЭП);
                }
            }

            foreach (var record in reportData)
            {
                MessageBox.Show(record.НомерСзноЗпзп + " " + record.КодОКПД2 + " " + record.НаименованиеРЭП + " " + 
                    record.Количество + " " + record.СтоимостьРуб + " " + record.ОтечественнаяРЭП + " " + 
                    record.НомерПлатежногоПоручения + " " + record.ДатаПлатежногоПоручения + " ");
            }

            return reportData;
        }

        private List<TableDataTotal> GetTotalData(List<TableData> reportData)
		{
            // преобразование данных для отчета
            var reportDataTotal = reportData
                .Select(x => new TableDataTotal
                {
                    КодОКПД2 = x.КодОКПД2,
                    Количество = x.Количество,
                    СтоимостьРуб = x.СтоимостьРуб,
                    ОтечественнаяРЭП = x.ОтечественнаяРЭП,
                    ПлатежноеПоручение = "№ " + x.НомерПлатежногоПоручения + " от " + x.ДатаПлатежногоПоручения.ToString("dd.MM.yyyy")
                })
                .GroupBy(x => x.Ключ())
                .Select(x => new TableDataTotal(x.Select(t => t).ToList()))
                .ToList();

            foreach (var record in reportDataTotal)
            {
                MessageBox.Show(record.КодОКПД2 + " " + record.Количество + " " + record.СтоимостьРуб + " " + record.ОтечественнаяРЭП + " " + record.ПлатежноеПоручение);
            }

            return reportDataTotal;
        }

        private void MakeAccountingREPReport(Xls xls, List<TableData> reportData, НастройкиОтчета reportParameters, LogDelegate logDelegate, ProgressBar progressBar)
        {
            //Задание начальных условий для progressBar
            var progressStep = reportData.Count / 50;
            int step = 0;

            #region Вывод детальных записей

            xls.Worksheet.Name = "Детально";

            var col = InsertHeader("Период: " + reportParameters.НачалоПериода.ToString("dd.MM.yy") + " - " + reportParameters.КонецПериода.ToString("dd.MM.yy"), 3, 1, xls);
            col = HeaderDetailTable(xls, 1, 3);

            int row = 4;

            foreach (var data in reportData.OrderBy(t => t.КодОКПД2))
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }

                logDelegate(String.Format("Запись объекта: {0} {1} {2}", data.КодОКПД2, data.ОтечественнаяРЭП, data.НомерСзноЗпзп));

                xls[1, row].SetValue(data.НомерСзноЗпзп);
                xls[2, row].SetValue(data.КодОКПД2);
                xls[3, row].SetValue(data.НаименованиеРЭП);
                xls[4, row].NumberFormat = "@";
                xls[4, row].SetValue(data.Количество);
                xls[5, row].NumberFormat = "@";
                xls[5, row].SetValue(data.СтоимостьРуб);
                xls[6, row].SetValue(data.ОтечественнаяРЭП ? "Отечественная РЭП" : "Иностранная РЭП");
                xls[7, row].SetValue(data.НомерПлатежногоПоручения);
                xls[8, row].SetValue(data.ДатаПлатежногоПоручения);

                row++;
            }

            #endregion

            var reportDataTotal = GetTotalData(reportData);
			
            #region Вывод итоговых данных
			
            xls.AddWorksheet("Итого", true);

            col = InsertHeader("Период: " + reportParameters.НачалоПериода.ToString("dd.MM.yy") + " - " + reportParameters.КонецПериода.ToString("dd.MM.yy"), 3, 1, xls);          
            col = HeaderTotalTable(xls, 1, 3);

            row = 4;

            foreach (var data in reportDataTotal.OrderBy(t => t.КодОКПД2))
            {
                step++;
                if (step >= progressStep)
                {
                    progressBar.PerformStep();
                    step = 0;
                }

                logDelegate(String.Format("Запись объекта: {0} {1} {2}", data.КодОКПД2, data.ОтечественнаяРЭП, data.ПлатежноеПоручение));

                xls[1, row].SetValue(data.КодОКПД2);
                xls[2, row].NumberFormat = "@";
                xls[2, row].SetValue(data.Количество);
                xls[3, row].NumberFormat = "@";
                xls[3, row].SetValue(data.СтоимостьРуб);
                xls[4, row].SetValue(data.ОтечественнаяРЭП ? "Отечественная РЭП" : "Иностранная РЭП");
                xls[5, row].SetValue(data.ПлатежноеПоручение);

                row++;
            }
			
            #endregion
		}

        public static int HeaderDetailTable(Xls xls, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("Номер СЗНО/ЗПЗП", col, row, xls);
            col = InsertHeader("Код ОКПД2", col, row, xls);
            col = InsertHeader("Наименование РЭП", col, row, xls);
            col = InsertHeader("Количество, шт.", col, row, xls);
            col = InsertHeader("Стоимость, руб", col, row, xls);
            col = InsertHeader("Отечественная РЭП", col, row, xls);
            col = InsertHeader("Номер платежного поручения", col, row, xls);
            col = InsertHeader("Дата платежного поручения", col, row, xls);

            return col;
        }

        public static int HeaderTotalTable(Xls xls, int col, int row)
        {
            // формирование заголовков отчета сводной ведомости
            col = InsertHeader("Код ОКПД2", col, row, xls);
            col = InsertHeader("Количество, шт.", col, row, xls);
            col = InsertHeader("Стоимость, руб", col, row, xls);
            col = InsertHeader("Отечественная РЭП", col, row, xls);
            col = InsertHeader("Платежное поручение", col, row, xls);

            return col;
        }

        public static int InsertHeader(string header, int col, int row, Xls xls)
        {
            xls[col, row].Font.Name = "Calibri";
            xls[col, row].Font.Size = 11;
            xls[col, row].SetValue(header);
            xls[col, row].Font.Bold = true;
            xls[col, row].Interior.Color = Color.LightGray;
            xls[col, row].VerticalAlignment = XlVAlign.xlVAlignCenter;
            xls[col, row].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[col, row].Columns.AutoFit();
            //xls[col, row].WrapText = true;

            col++;

            return col;
        }
    }
}
