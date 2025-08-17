using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Stages;
using TFlex.Reporting;
using ReportHelpers;
using Filter = TFlex.DOCs.Model.Search.Filter;
using System.Diagnostics;
using TFlex.DOCs.Model.References.Messages;
using System.Globalization;
using static System.Windows.Forms.AxHost;
using System.Security.Cryptography;

namespace PaymentsReport
{
    public class PaymentsReport
    {
        public static void Generate(MacroProvider macro)
        {
            var posts = macro.ТекущийПользователь.РодительскиеОбъекты;
            bool rights = false;
            // Только админы и специалисты ЕРКЦ могут формировать РПП
            if (macro.ТекущийПользователь.Тип.УникальныйИдентификатор == Guids.Users.Objects.adminTypeGuid)
                rights = true;
            foreach (ReferenceObject post in posts)
            {
                if (post.SystemFields.Guid == Guids.Users.Objects.specERKCGuid)
                    rights = true;
            }

            if (!rights)
                macro.Сообщение("Внимание!", "У Вас недостаточно прав для формирования отчета");
            else
            {
                string regNumber = macro.ТекущийОбъект.Параметр[Guids.RegCardsReference.Fields.registrationNumberGuid.ToString()];
                DateTime regDate = macro.ТекущийОбъект.Параметр[Guids.RegCardsReference.Fields.registrationDateGuid.ToString()];
                DateTime startDate = macro.ТекущийОбъект.Параметр[Guids.RegCardsReference.RPP.Fields.startDateGuid.ToString()];
                DateTime finishDate = macro.ТекущийОбъект.Параметр[Guids.RegCardsReference.RPP.Fields.finishDateGuid.ToString()];

                var диалог = macro.СоздатьДиалогВвода("Реестр платежей и поступлений");
                диалог.ДобавитьСтроковое("Номер реестра платежей и поступлений"); диалог["Номер реестра платежей и поступлений"] = regNumber.Substring(4);
                диалог.ДобавитьГруппу("Выбор временных рамок отчета");
                диалог.ДобавитьДату("Начальная дата выборки"); диалог["Начальная дата выборки"] = startDate;
                диалог.ДобавитьДату("Конечная дата выборки"); диалог["Конечная дата выборки"] = finishDate;
                диалог.ДобавитьФлаг("Показывать оплаченные документы");
                диалог.ЗначениеИзменено += (имяПоля, староеЗначение, новоеЗначение) =>
                {
                    if ((имяПоля == "Начальная дата выборки" || имяПоля == "Конечная дата выборки") && диалог["Начальная дата выборки"] > диалог["Конечная дата выборки"])
                    {
                        macro.Сообщение("Ошибка!", "Начальная дата больше конечной. Отчет не будет сформирован");
                        return;
                    }
                };

                if (диалог.Показать())
                {
                    if (диалог["Начальная дата выборки"] > диалог["Конечная дата выборки"])
                    {
                        macro.Сообщение("Ошибка!", "Начальная дата больше конечной. Отчет не будет сформирован");
                        return;
                    }

                    ObjectOperations objOps = new ObjectOperations();

                    string relativePath = Guids.Files.Paths.rppFolderPath;
                    string yearFolderPath = Path.Combine(relativePath, regDate.ToString("yyyy"));
                    string monthFolderPath = Path.Combine(yearFolderPath, regDate.ToString("MM - MMMM"));
                    string correctNumber = objOps.GetCorrectFileName(regNumber);
                    string objectFolderPath = Path.Combine(monthFolderPath, correctNumber + " " + regDate.ToShortDateString());
                    string fileName = Path.Combine(monthFolderPath, correctNumber + " " + regDate.ToShortDateString() + ".xlsx");
                    // Поиск в справочнике файлов объекта с заданным названием
                    FileReference fileReference = new FileReference(ServerGateway.Connection);
                    FileReferenceObject fObject = fileReference.FindByPath(fileName);                   
                    
                    if (fObject != null)
                    {
                        macro.Сообщение("Внимание!", "Файл отчета с таким номером уже существует.\nДля нового отчета создайте новую РКК типа\n\"Реестр платежей и поступлений\"");
                        return;
                    }

                    ReportParameters reportParameters = new ReportParameters();

                    // Передача данных о начальной и конечной дате выборки СЗНО
                    reportParameters.beginDate = startDate;
                    reportParameters.endDate = finishDate;
                    var reportGenerator = new ReportGenerator();
                    reportGenerator.Make(reportParameters, correctNumber, regDate.ToShortDateString(), диалог["Показывать оплаченные документы"], macro);
                }
                else
                {
                    macro.Сообщение("Внимание!", "Создание отчета отменено пользователем");
                }
            }
        }
    }

    public class ReportGenerator
    {
        public void Make(ReportParameters reportParameters, string reportNumber, string reportDate, bool signChecked, MacroProvider macro)
        {
            var условия = new List<string>();
            условия.Add(String.Format("[Тип] = '{0}'", Guids.RegCardsReference.SZNO.typeGuid.ToString()));
            условия.Add(String.Format("[{0}] >= '{1}'", Guids.RegCardsReference.SZNO.Fields.confirmDateGuid.ToString(), reportParameters.beginDate));
            условия.Add(String.Format("[{0}] < '{1}'", Guids.RegCardsReference.SZNO.Fields.confirmDateGuid.ToString(), reportParameters.endDate.AddDays(1)));
            var stages = new string[] { Guids.RegCardsReference.Stages.stageSignedGuid.ToString(), Guids.RegCardsReference.Stages.stageInArchiveGuid.ToString() };
            условия.Add("([Стадия] входит в список '" + string.Join(",", stages) + "'");
            if (signChecked)
                условия.Add(String.Format("[{0}] = true", Guids.RegCardsReference.SZNO.Fields.paidGuid.ToString()));
            else
                условия.Add(String.Format("[{0}] = false", Guids.RegCardsReference.SZNO.Fields.paidGuid.ToString()));
            var условие = String.Join(" И ", условия);

            var regCardsSZNO = macro.НайтиОбъекты(Guids.RegCardsReference.refGuid.ToString(), условие);

            if (regCardsSZNO.Count == 0)
            {
                macro.Сообщение("Внимание!", "Нет документов в списке");
                return;
            }
            // Заполнение параметров СЗНО
            List<ReportParameters> regCards = FillParamList(regCardsSZNO, reportParameters.beginDate, reportParameters.endDate, macro);
            if (regCards.Count == 0)
            {
                macro.Сообщение("Внимание!", "Формирование отчета отменено пользователем");
                return;
            }
            // Формирование отчета
            FillReport(regCards, reportNumber, reportDate, macro);
        }
        private List<ReportParameters> FillParamList(Объекты refobjects, DateTime beginDate, DateTime endDate, MacroProvider macro)
        {
            List<ReportParameters> paramList = new List<ReportParameters>();
            int i = 0;
            int cnt = 0;

            StringBuilder sb = new StringBuilder();

            macro.ДиалогОжидания.Показать("Идет чтение данных из журнала СЗНО...", true);

            foreach (var refObject in refobjects)
            {
                ReportParameters parameters = new ReportParameters();
                i++;
                parameters.number = Convert.ToString(i);
                parameters.registrationNumber = refObject[Guids.RegCardsReference.Fields.registrationNumberGuid.ToString()] + "\n от " +
                    ((DateTime)refObject[Guids.RegCardsReference.Fields.registrationDateGuid.ToString()]).Date.ToShortDateString();
                parameters.invoiceNumber = "";
                parameters.invoiceData = "";
                parameters.contractor = refObject[Guids.RegCardsReference.SZNO.Fields.contractorGuid.ToString()] + "";
                parameters.productName = refObject[Guids.RegCardsReference.SZNO.Fields.contentGuid.ToString()] + "";
                parameters.paymentFunction = refObject[Guids.RegCardsReference.SZNO.Fields.paymentFunctionGuid.ToString()] + "";
                parameters.paymentKind = refObject[Guids.RegCardsReference.SZNO.Fields.paymentKindGuid.ToString()] + "";
                parameters.articleCode = "";
                parameters.paymentSum = 0;
                parameters.orderNumber = refObject[Guids.RegCardsReference.SZNO.Fields.orderNumberGuid.ToString()] + "";
                if (refObject.СвязанныйОбъект[Guids.RegCardsReference.SZNO.Links.DZULinks.ToString()] == null)
                {
                    parameters.contractRegNum = String.Empty;
                    parameters.contractNumber = String.Empty;
                }
                else
                {
                    parameters.contractRegNum = refObject.СвязанныйОбъект[Guids.RegCardsReference.SZNO.Links.DZULinks.ToString()]
                        [Guids.RegCardsReference.DZU.Fields.regNumberGuid.ToString()] + "";
                    parameters.contractNumber = refObject.СвязанныйОбъект[Guids.RegCardsReference.SZNO.Links.DZULinks.ToString()]
                        [Guids.RegCardsReference.DZU.Fields.numberGuid.ToString()] + "";
                }
                parameters.ptpPoint = refObject[Guids.RegCardsReference.SZNO.Fields.ptpPointGuid.ToString()] + "";
                parameters.comment = "";
                parameters.stage = (refObject[Guids.RegCardsReference.SZNO.Fields.confirmDateGuid.ToString()] == null) ? String.Empty :
                    ((DateTime)(refObject[Guids.RegCardsReference.SZNO.Fields.confirmDateGuid.ToString()])).ToShortDateString();

                // Если установлен параметр - Оплачено, то булева переменная = true, иначе false
                parameters.isStagePaid = refObject[Guids.RegCardsReference.SZNO.Fields.paidGuid.ToString()];
                parameters.draftNumber = refObject[Guids.RegCardsReference.SZNO.Fields.draftNumberGuid.ToString()] + "";
                parameters.draftDate = (refObject[Guids.RegCardsReference.SZNO.Fields.draftDateGuid.ToString()] == null) ? String.Empty :
                    ((DateTime)(refObject[Guids.RegCardsReference.SZNO.Fields.draftDateGuid.ToString()])).ToShortDateString();
                parameters.cell = String.Empty;
                parameters.beginDate = beginDate;
                parameters.endDate = endDate;

                var budgetArticlesList = refObject.СвязанныеОбъекты[Guids.RegCardsReference.SZNO.Links.budgetArticlesListLink.ToString()];
                if (budgetArticlesList.Count == 0)
                {
                    // В случае отсутствия связи со справочником "Статьи бюджета" выводим сообщение пользователю и берем сумму из поля "Всего сумма к оплате"
                    sb = sb.AppendLine(refObject[Guids.RegCardsReference.Fields.registrationNumberGuid.ToString()] + " от " +
                        ((DateTime)refObject[Guids.RegCardsReference.Fields.registrationDateGuid.ToString()]).Date.ToShortDateString() + " :");
                    parameters.paymentSum = refObject[Guids.RegCardsReference.SZNO.Fields.paymentSumGuid.ToString()] == null ? 0 :
                        Convert.ToDouble(refObject[Guids.RegCardsReference.SZNO.Fields.paymentSumGuid.ToString()].ToString());
                    paramList.Add(parameters);
                }
                else
                {
                    i--;
                    // Иначе дублируем параметры СЗНО и меняем в них поля "Код БДДС" и "Сумма к выплате" согласно строк табличной части
                    foreach (var item in budgetArticlesList)
                    {
                        i++;
                        var pars = (ReportParameters)parameters.Clone();
                        pars.number = Convert.ToString(i);
                        pars.articleCode = item[Guids.RegCardsReference.SZNO.BudgetArticlesList.budgetArticleCode.ToString()] + "";
                        pars.paymentSum = item[Guids.RegCardsReference.SZNO.BudgetArticlesList.paymentSum.ToString()] == null ? 0 :
                            Convert.ToDouble(item[Guids.RegCardsReference.SZNO.BudgetArticlesList.paymentSum.ToString()].ToString());
                        paramList.Add(pars);
                    }
                }
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.125));
                cnt++;
                double percent = refobjects.Count() == 0 ? 0 : (double)cnt / refobjects.Count() * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Читаются данные документа " + parameters.registrationNumber, percent))
                {
                    paramList.Clear();
                    sb = sb.Clear();
                }
            }

            macro.ДиалогОжидания.Скрыть();

            if (sb.ToString() != String.Empty)
            {
                sb.AppendLine("потеряна связь со справочником 'Статьи бюджета'");
                sb.AppendLine("Суммы по статьям бюджета вычисляться не будут");
                macro.Сообщение("Внимание!", sb.ToString());
            }

            return paramList;
        }

        // Создание отчета
        private void FillReport(List<ReportParameters> paramList, string reportNumber, string reportDate, MacroProvider macro)
        {
            ReferenceObject refObject = (ReferenceObject)macro.ТекущийОбъект;
            int objectID = refObject.SystemFields.Id;
            bool isCatch = false;

            var objectOps = new ObjectOperations();

            var reference = new FileReference(ServerGateway.Connection);
            var pattern = reference.FindByPath(Guids.Files.Paths.patternPath);
            // Перезапись файла шаблона в случае его отсутствия или изменения
            if (!File.Exists(pattern.LocalPath) || File.GetLastWriteTime(pattern.LocalPath) != pattern.SystemFields.EditDate)
            {
                pattern.GetHeadRevision();
            }

            string fileExtension = Path.GetExtension(pattern.LocalPath);
            string relativePath = Guids.Files.Paths.rppFolderPath;
            string newFileName = reportNumber + " " + reportDate + fileExtension;
            string appDataFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), newFileName);

            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(pattern.LocalPath);
                File.SetAttributes(pattern.LocalPath, FileAttributes.Normal);

                var excelOps = new ExcelOperations();
                // Настройки параметров страницы              
                excelOps.PageSetup(xls, XlPaperSize.xlPaperA4, XlPageOrientation.xlLandscape, 0, 0, 0, 0);
                // Создание заголовка
                excelOps.CreateHeader(xls);
                // Заполнение таблицы данными
                int sumRow = excelOps.FillTable(paramList, xls, macro);
                xls[1, 15, 17, 4 + paramList.Count].CenterText();
                xls[1, 15, 17, 4 + paramList.Count].WrapText = true;

                // Создание "шапки" документа (с подписями)
                excelOps.CreateDocTemplate(xls, paramList, sumRow, reportNumber, reportDate, macro);

                // Защита листа книги
                // string password = "erkc908";
                // xls.Worksheet.Protect(password, true, true, true, false, false, false, false, false, false, false, false, false, false, false, false);
                // Сохранение файла
                if (File.Exists(appDataFileName))
                {
                    File.Delete(appDataFileName);
                }
                xls.SaveAs(appDataFileName);
                // Сохранение файла в хранилище ОРД (справочник "Файлы" DOCs)
                FileReference fileReference = new FileReference(ServerGateway.Connection);
                // Получение папки справочника файлы по его relativePath
                FolderObject mainFolder = (FolderObject)fileReference.FindByRelativePath(relativePath);
                if (mainFolder == null)
                {
                    macro.Сообщение("Ошибка!", "Папка для хранения файла не найдена./nФайл в хранилище не добавлен");
                    isCatch = true;
                }
                else
                {
                    // Получение пути сохранения файла
                    DateTime rDate = Convert.ToDateTime(reportDate);
                    string yearFolderPath = Path.Combine(relativePath, rDate.ToString("yyyy"));
                    string monthFolderPath = Path.Combine(yearFolderPath, rDate.ToString("MM - MMMM"));
                    string objectFolderPath = Path.Combine(monthFolderPath, reportNumber + " " + rDate.ToShortDateString());
                    var yearFolder = objectOps.GetChildFolder(fileReference, mainFolder, yearFolderPath, rDate.ToString("yyyy"));
                    var monthFolder = objectOps.GetChildFolder(fileReference, yearFolder, monthFolderPath, rDate.ToString("MM - MMMM"));
                    string correctFileName = objectOps.GetCorrectFileName(reportNumber);
                    var objectFolder = objectOps.GetChildFolder(fileReference, monthFolder, objectFolderPath, reportNumber + " " + rDate.ToShortDateString());
                    // Добавление файла в папку
                    FileObject fileObject = fileReference.AddFile(appDataFileName, objectFolder);
                    Desktop.CheckIn(fileObject, "Добавление файла", false);
                    // Связывание файла с объектом РКК
                    // Получение объекта РКК
                    var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.RegCardsReference.refGuid);
                    var regCardsReference = referenceInfo.CreateReference();
                    ReferenceObject rppObject = regCardsReference.Find(objectID);

                    rppObject.BeginChanges();
                    rppObject.Links.AnyReference[Guids.RegCardsReference.Links.DocumentLinks].AddLinkedObject(fileObject);
                    rppObject.EndChanges();
                    macro.Сообщение("Успех!", "Файл отчета создан и присоединен к соответствующей РКК");
                }
            }
            catch (Exception exp)
            {
                macro.Сообщение("Исключительная ситуация!", exp.Message);
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
                    Process.Start(appDataFileName);
                }
            }
        }
    }
}