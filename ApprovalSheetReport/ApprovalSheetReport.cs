using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TFlex.Reporting;
using Microsoft.Office.Interop.Word;
using ReportHelpers;
using TFlex.DOCs.Model;
using TFlex.DOCs.Common;
using System.Windows.Forms;

namespace ApprovalSheetReport
{
    public class WordGenerator : IReportGenerator
    {
        public string DefaultReportFileExtension
        {
            get
            {
                return ".docx";
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
                using (new WaitCursorHelper(false))
                {
                    ApprovalSheetReport.MakeReport(context);
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

    public class ApprovalSheetReport
    {
        public void Dispose()
        {
        }

        public static void MakeReport(IReportGenerationContext context)
        {
            ApprovalSheetReport report = new ApprovalSheetReport();
            //LogDelegate m_writeToLog;

            report.Make(context);
        }

        public void Make(IReportGenerationContext context)
        {
            context.CopyTemplateFile();    // Создаем копию шаблона

            var doc = new WordDoc();

            ApprovalSheetReport report = new ApprovalSheetReport();

            TableData данныеОтчета = GetData(context);

            if (данныеОтчета == null)
            {
                return;
            }

            try
            {
                doc.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                //  xls.Open(context.ReportFilePath);

                //Формирование отчета
                MakeApprovalSheetReport(doc, данныеОтчета);

                //  xls.Save();
                doc.Application.DisplayAlerts = WdAlertLevel.wdAlertsAll;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e + "", "Exception!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                // выводим на экран                
                doc.Visible = true;
                doc.ClearApplicationLink();
                MessageBox.Show("Отчет успешно сформирован - разверните из панели внизу");
            }

        }

        private TableData GetData(IReportGenerationContext context)
        {
            var connection = ServerGateway.Connection;
            var проектыДокументовInfo = connection.ReferenceCatalog.Find(Guids.ПроектыДокументов.Id);
            var проектыДокументов = проектыДокументовInfo.CreateReference();

            var проект = проектыДокументов.Find(context.ObjectsInfo[0].ObjectId);

            // формирование данных по проекту документа
            var данныеПроекта = new TableData();
            данныеПроекта.РегистрационныйНомер = проект[Guids.ПроектыДокументов.Параметры.РегистрационныйНомер].GetString();
            данныеПроекта.ДатаРегистрации = проект[Guids.ПроектыДокументов.Параметры.ДатаРегистрации].GetDateTime();
            данныеПроекта.ВидДокумента = проект[Guids.ПроектыДокументов.Параметры.ВидДокумента].GetString();
            данныеПроекта.Закрыто = проект[Guids.ПроектыДокументов.Параметры.Закрыто].GetBoolean();
            данныеПроекта.СрокCогласования = проект[Guids.ПроектыДокументов.Параметры.СрокСогласования].GetDateTime();
            данныеПроекта.НаименованиеКраткоеСодержание = проект[Guids.ПроектыДокументов.Параметры.НаименованиеКраткоеСодержание].GetString();
            данныеПроекта.ПоследняяВерсияФайла = проект.GetObjects(Guids.ПроектыДокументов.Связи.ФайлыПроектаДокумента)
                                                       .OrderByDescending(t => t.SystemFields.CreationDate)
                                                       .First()
                                                       .ToString();            
            данныеПроекта.РешенияОСогласовании = new List<РешениеОСогласовании>();
            foreach(var решение in проект.GetObjects(Guids.ПроектыДокументов.Связи.РешенияОСогласовании))
			{
                var решениеОСогласовании = new РешениеОСогласовании();
                решениеОСогласовании.СогласующееЛицо = решение.GetObject(Guids.РешенияОСогласовании.Связи.СогласующееЛицо).ToString();
                решениеОСогласовании.ВерсияФайла = решение[Guids.РешенияОСогласовании.Параметры.ВерсияФайла].GetString();
                решениеОСогласовании.ДатаCогласования = решение[Guids.РешенияОСогласовании.Параметры.ДатаСогласования].GetDateTime();
                решениеОСогласовании.СрокCогласования = решение[Guids.РешенияОСогласовании.Параметры.СрокСогласования].GetDateTime();
                решениеОСогласовании.ТекущаяСтадия = решение.SystemFields.Stage.ToString();
                
                решениеОСогласовании.ЗамечанияПредложения = new List<ЗамечаниеПредложение>();
                foreach (var запись in решение.GetObjects(Guids.РешенияОСогласовании.СпискиОбъектов.ЗамечанияПредложения.Id))
				{
                    var замечаниеПредложение = new ЗамечаниеПредложение();
                    замечаниеПредложение.ВерсияФайла = запись[Guids.РешенияОСогласовании.СпискиОбъектов.ЗамечанияПредложения.Параметры.ВерсияФайла].GetString();
                    замечаниеПредложение.ДатаРешения = запись[Guids.РешенияОСогласовании.СпискиОбъектов.ЗамечанияПредложения.Параметры.ДатаРешения].GetDateTime();
                    замечаниеПредложение.ЗамечанияПредложения = запись[Guids.РешенияОСогласовании.СпискиОбъектов.ЗамечанияПредложения.Параметры.ЗамечанияПредложения].GetString();
                    замечаниеПредложение.СтадияРешения = запись[Guids.РешенияОСогласовании.СпискиОбъектов.ЗамечанияПредложения.Параметры.СтадияРешения].GetString();

                    решениеОСогласовании.ЗамечанияПредложения.Add(замечаниеПредложение);
                }

                данныеПроекта.РешенияОСогласовании.Add(решениеОСогласовании);
            }

            return данныеПроекта;
        } 

        private void MakeApprovalSheetReport(WordDoc doc, TableData данныеОтчета)
        {            
            doc.InsertParagraph(null, "ЛИСТ СОГЛАСОВАНИЯ", WdParagraphAlignment.wdAlignParagraphCenter, WordDoc.StyleBold);
            doc.InsertParagraph(null, String.Format("на {0}", DateTime.Now.ToString("dd MMMM yyyy")), WdParagraphAlignment.wdAlignParagraphCenter, WordDoc.StyleBold);
            doc.InsertParagraph(null, String.Format("Проект документа: {0}", данныеОтчета.РегистрационныйНомер + " от " + данныеОтчета.ДатаРегистрации.ToShortDateString()), WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);
            doc.InsertParagraph(null, String.Format("Вид документа: {0}", данныеОтчета.ВидДокумента), WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);
            doc.InsertParagraph(null, данныеОтчета.Закрыто ? "Проект закрыт" : "Проект в процессе согласования", WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);
            doc.InsertParagraph(null, String.Format("Краткое описание: {0}", данныеОтчета.НаименованиеКраткоеСодержание), WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);
            doc.InsertParagraph(null, String.Format("Последняя версия файла: {0}", данныеОтчета.ПоследняяВерсияФайла), WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);

            var table = new DocTable(4); // 4 - количество столбцов
            table.AddCells(new Action<object>[] { WordDoc.StyleBold },
            new DocTableCell2 { content = "Согласующее лицо", positionX = 1, positionY = 1 },
            new DocTableCell2 { content = "Дата согласования", positionX = 2, positionY = 1 },
            new DocTableCell2 { content = "Принятое решение", positionX = 3, positionY = 1 },
            new DocTableCell2 { content = "Версия файла проекта", positionX = 4, positionY = 1 }            
            );

            foreach (var решение in данныеОтчета.РешенияОСогласовании)
			{                
                table.AddCells(null,
                    new DocTableCell2 { content = решение.СогласующееЛицо, positionX = 1, positionY = 1 },
                    new DocTableCell2 { content = решение.ДатаCогласования.ToShortDateString() != "01.01.0001" ? решение.ДатаCогласования.ToShortDateString() : "", positionX = 2, positionY = 1 },
                    new DocTableCell2 { content = решение.ТекущаяСтадия, positionX = 3, positionY = 1 },
                    new DocTableCell2 { content = решение.ВерсияФайла, positionX = 4, positionY = 1 }                    
                    );
            }

            // формирование таблицы
            table.BuildTable(doc, WdAutoFitBehavior.wdAutoFitFixed);

            doc.InsertParagraph(null, "", WdParagraphAlignment.wdAlignParagraphLeft);

            // информация о замечаниях им предложениях по проекту

            // вставка разрыва страницы пв последней позиции документа
            // doc.Document.Range(doc.Document.Content.End, doc.Document.Content.End).InsertBreak(WdBreakType.wdPageBreak);
            doc.Document.Paragraphs[doc.Document.Paragraphs.Count].Range.InsertBreak(WdBreakType.wdPageBreak);

            var header = true;

            foreach (var решение in данныеОтчета.РешенияОСогласовании)
			{
                var таблицаЗамечаний = new DocTable(1); // 1 - количество столбцов

                if (решение.ЗамечанияПредложения.Count > 0)
				{
                    if (header)
                    {
                        doc.InsertParagraph(null, "ЗАМЕЧАНИЯ И ПРЕДЛОЖЕНИЯ,\r\n выданные при согласовании проекта документа", WdParagraphAlignment.wdAlignParagraphCenter, WordDoc.StyleBold);
                        doc.InsertParagraph(null, "", WdParagraphAlignment.wdAlignParagraphLeft);
                        header = false;
                    }

                    doc.InsertParagraph(null, решение.СогласующееЛицо, WdParagraphAlignment.wdAlignParagraphLeft, WordDoc.StyleBold);
                                        
                    foreach (var запись in решение.ЗамечанияПредложения)
                    {
                        таблицаЗамечаний.AddCells(new Action<object>[] { WordDoc.StyleBold, WordDoc.StyleBackgroundGray },
                            new DocTableCell2 { content = String.Join(", ", запись.СтадияРешения, запись.ДатаРешения.ToShortDateString(), запись.ВерсияФайла), positionX = 1, positionY = 1 });
                        таблицаЗамечаний.AddCells(null, new DocTableCell2 { content = запись.ЗамечанияПредложения, positionX = 1, positionY = 1 });
                    }
                    
                    таблицаЗамечаний.BuildTable(doc, WdAutoFitBehavior.wdAutoFitFixed);
                    
                    doc.InsertParagraph(null, "", WdParagraphAlignment.wdAlignParagraphLeft);
                }
            }
        }
    }

    public class WaitCursorHelper : IDisposable
    {
        private bool _waitCursor;

        public WaitCursorHelper(bool useWaitCursor)
        {
            _waitCursor = System.Windows.Forms.Application.UseWaitCursor;
            System.Windows.Forms.Application.UseWaitCursor = useWaitCursor;
        }

        public void Dispose()
        {
            System.Windows.Forms.Application.UseWaitCursor = _waitCursor;
        }
    }
}

