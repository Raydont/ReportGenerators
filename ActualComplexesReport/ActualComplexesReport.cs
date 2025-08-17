using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Windows.Forms;
using System.Globalization;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Macros;
using Microsoft.Office.Interop.Excel;
using TFlex.Reporting;
using TFlex.DOCs.Model.References.Files;
using ReportHelpers;
using System.Diagnostics;
using static ActualComplexesReport.Guids;

namespace ActualComplexesReport
{
    public static class ActualComplexesReportGenerator
    {
        public static void Generate(MacroProvider macro)
        {
            var customerTerms = GetCustomerList();
            var exploitationTerms = new List<string> { FormParameters.FormValues.Exploitation.True,
                FormParameters.FormValues.Exploitation.False, 
                FormParameters.FormValues.Exploitation.All };
            var structureTerms = new List<string> { FormParameters.FormValues.Structures.True,
                FormParameters.FormValues.Structures.False,
                FormParameters.FormValues.Structures.All };
            var cancelledTerms = new List<string> { FormParameters.FormValues.Cancelled.True,
                FormParameters.FormValues.Cancelled.False,
                FormParameters.FormValues.Cancelled.All };
            var actualTerms = new List<string> { FormParameters.FormValues.Actual.True,
                FormParameters.FormValues.Actual.False,
                FormParameters.FormValues.Actual.All };
            var checkedObjectsTerms = new List<string> { FormParameters.FormValues.CheckedObjects.True,
                FormParameters.FormValues.CheckedObjects.False,
                FormParameters.FormValues.CheckedObjects.All };

            Объект device = null;
            Объект order = null;
            var диалог = macro.СоздатьДиалогВвода("Параметры сводного отчета");
            диалог.ДобавитьГруппу("Отбор по аппаратуре");
            диалог.ДобавитьПодборОбъекта(FormParameters.FormNames.Device, 
                Guids.ComplexReference.refGuid.ToString(),
                Guids.ComplexReference.Fields.name.ToString());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Device, "* По умолчанию - все комплексы");
            диалог.ДобавитьСтроковое(FormParameters.FormNames.Denotation);
            диалог.ДобавитьГруппу("Отбор по заказу");
            диалог.ДобавитьПодборОбъекта(FormParameters.FormNames.Order, 
                Guids.OrderStructureReference.refGuid.ToString(), 
                Guids.OrderStructureReference.Fields.orderNumber.ToString());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Order, "* По умолчанию - все заказы");
            диалог.ДобавитьГруппу("Дополнительные параметры отбора");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.Customer, customerTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Customer, "* Отбор по типу потребителя");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.Exploitation, exploitationTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Exploitation, "* Отбор по признаку эксплуатации");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.CheckedObjects, checkedObjectsTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.CheckedObjects, "* Отбор по наличию проверяемых изделий");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.Structures, structureTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Structures, "* Отбор по наличию файлов структур");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.Actual, actualTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Actual, "* Отбор по наличию рабочих экземпляров");
            диалог.ДобавитьФлаг(FormParameters.FormNames.IsFull);
            диалог.ДобавитьКомментарий(FormParameters.FormNames.IsFull, "* С подробной информацией об аппаратуре");
            диалог.ДобавитьВыборИзСписка(FormParameters.FormNames.Cancelled, cancelledTerms.ToArray());
            диалог.ДобавитьКомментарий(FormParameters.FormNames.Cancelled, "* Отбор по признаку списания аппаратуры");
            диалог.УстановитьВидимостьЭлемента(FormParameters.FormNames.Cancelled, false);
            диалог.ДобавитьТекст("", 1);
            диалог.ЗначениеИзменено += (имяПоля, староеЗначение, новоеЗначение) =>
            {
                if (имяПоля == FormParameters.FormNames.Device)
                {
                    device = диалог[FormParameters.FormNames.Device];
                    if (device != null)
                        диалог[FormParameters.FormNames.Denotation] = device[Guids.ComplexReference.Fields.denotation.ToString()];
                    else
                        диалог[FormParameters.FormNames.Denotation] = string.Empty;
                }
                if (имяПоля == FormParameters.FormNames.Order)
                {
                    order = диалог[FormParameters.FormNames.Order];
                }
                if (имяПоля == FormParameters.FormNames.IsFull)
                {
                    if (новоеЗначение == true)
                        диалог.УстановитьВидимостьЭлемента(FormParameters.FormNames.Cancelled, true);
                    else
                        диалог.УстановитьВидимостьЭлемента(FormParameters.FormNames.Cancelled, false);
                }
            };
            // Начальная установка выбора из списка
            диалог[FormParameters.FormNames.Exploitation] = exploitationTerms[2];
            диалог[FormParameters.FormNames.Customer] = customerTerms[3];
            диалог[FormParameters.FormNames.CheckedObjects] = checkedObjectsTerms[2];
            диалог[FormParameters.FormNames.Structures] = structureTerms[2];
            диалог[FormParameters.FormNames.Actual] = actualTerms[2];
            диалог[FormParameters.FormNames.Cancelled] = cancelledTerms[2];
            диалог.Ширина = 700;
            if (диалог.Показать())
            {
                var flags =  new bool[]{ диалог[FormParameters.FormNames.IsFull] };
                var choosedParameters = new string[] { диалог[FormParameters.FormNames.Exploitation],
                    диалог[FormParameters.FormNames.Customer],
                    диалог[FormParameters.FormNames.CheckedObjects],
                    диалог[FormParameters.FormNames.Structures],
                    диалог[FormParameters.FormNames.Actual],
                    диалог[FormParameters.FormNames.Cancelled] };
                Make(device, order, choosedParameters, flags,  macro);
            }
        }

        public static void Make(Объект device, Объект order, string[] choosedParameters, bool[] flags, MacroProvider macro)
        {
            // Признак расширенного вида отчета
            bool isFull = flags[0];
            // Условие поиска
            var условия = new List<string>();
            // Обязательное условие - тип
            условия.Add(String.Format("[Тип] = '{0}'", Guids.OrderStructureReference.type.ToString()));
            // Условие по аппаратуре
            if (device != null)
                условия.Add(String.Format("[{0}]->[{1}] = '{2}' И [{0}]->[{3}] = '{4}'",
                    Guids.OrderStructureReference.Links.complexes.ToString(),
                    Guids.ComplexReference.Fields.name.ToString(),
                    device[Guids.ComplexReference.Fields.name.ToString()],
                    Guids.ComplexReference.Fields.denotation.ToString(),
                    device[Guids.ComplexReference.Fields.denotation.ToString()]));
            // Условие по номеру заказа
            if (order != null)
                условия.Add(String.Format("[{0}] = '{1}'", Guids.OrderStructureReference.Fields.orderNumber.ToString(), 
                    order[Guids.OrderStructureReference.Fields.orderNumber.ToString()]));
            // Выбираются либо заказы в эксплуатации, либо не в эксплуатации
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Exploitation] == FormParameters.FormValues.Exploitation.True)
                условия.Add(String.Format("[{0}] = true", Guids.OrderStructureReference.Fields.exploitation.ToString()));
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Exploitation] == FormParameters.FormValues.Exploitation.False)
                условия.Add(String.Format("[{0}] = false", Guids.OrderStructureReference.Fields.exploitation.ToString()));
            // Условие по типу заказчика
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Customer] != FormParameters.FormValues.Customer.All)
                условия.Add(String.Format("[{0}] = '{1}'", Guids.OrderStructureReference.Fields.customer.ToString(), choosedParameters[1]));
            // Условия по наличию списка проверяемых изделий
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.CheckedObjects] == FormParameters.FormValues.CheckedObjects.True)
                условия.Add(String.Format("[{0}] содержит какие-либо данные", Guids.OrderStructureReference.Links.checkedObjects.ToString()));
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.CheckedObjects] == FormParameters.FormValues.CheckedObjects.False)
                условия.Add(String.Format("[{0}] не содержит данных", Guids.OrderStructureReference.Links.checkedObjects.ToString()));
            // Условия по наличию структур аппаратуры
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Structures] == FormParameters.FormValues.Structures.True)
                условия.Add(String.Format("[{0}] != null", Guids.OrderStructureReference.Links.structures.ToString()));
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Structures] == FormParameters.FormValues.Structures.False)
                условия.Add(String.Format("[{0}] = null", Guids.OrderStructureReference.Links.structures.ToString()));
            // Условия по наличию рабочих экземпляров
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Actual] == FormParameters.FormValues.Actual.True)
                условия.Add(String.Format("[{0}] != null", Guids.OrderStructureReference.Links.actualComplexes.ToString()));
            if (choosedParameters[(int)FormParameters.StringTermsIndexes.Actual] == FormParameters.FormValues.Actual.False)
                условия.Add(String.Format("[{0}] = null", Guids.OrderStructureReference.Links.actualComplexes.ToString()));
            if (isFull)
            {
                // Условия по признаку списания аппаратуры
                if (choosedParameters[(int)FormParameters.StringTermsIndexes.Cancelled] == FormParameters.FormValues.Cancelled.True)
                    условия.Add(String.Format("[{0}]->[{1}] = true", Guids.OrderStructureReference.Links.actualComplexes.ToString(),
                        Guids.ActualComplexReference.Fields.cancelled.ToString()));
                if (choosedParameters[(int)FormParameters.StringTermsIndexes.Cancelled] == FormParameters.FormValues.Cancelled.False)
                    условия.Add(String.Format("[{0}]->[{1}] = false", Guids.OrderStructureReference.Links.actualComplexes.ToString(),
                        Guids.ActualComplexReference.Fields.cancelled.ToString()));
            }
            var условие = String.Join(" И ", условия);
            var orders = macro.НайтиОбъекты(Guids.OrderStructureReference.refGuid.ToString(), условие);
            var reportParameters = FillReportParameters(orders, isFull, macro);
            // Формирование отчета
            if (reportParameters.Count > 0)
                FillReport(reportParameters, isFull, macro);
            else
                macro.Сообщение("Внимание!", "По заданному Вами условию\nне найдено ни одного объекта");
        }
        private static List<string> GetCustomerList()
        {
            // Справочник "Структуры заказов"
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.OrderStructureReference.refGuid);
            var reference = referenceInfo.CreateReference();
            return reference.Objects.Select(t => t[Guids.OrderStructureReference.Fields.customer].Value.ToString()).
                Distinct().Append(FormParameters.FormValues.Customer.All).ToList();
        }
        private static List<ComplexParameters> FillReportParameters(Объекты orders, bool isFull, MacroProvider macro)
        {
            List<ComplexParameters> reportParameters = new List<ComplexParameters>();
            macro.ДиалогОжидания.Показать("Идет чтение данных из справочной системы...", true);
            int cnt = 0;
            foreach (var order in orders)
            {
                // Параметры по связи со справочником "Аппаратура"
                var name = order[String.Format("[{0}]->[{1}]", Guids.OrderStructureReference.Links.complexes.ToString(),
                    Guids.ComplexReference.Fields.name.ToString())];
                var denotation = order[String.Format("[{0}]->[{1}]", Guids.OrderStructureReference.Links.complexes.ToString(),
                    Guids.ComplexReference.Fields.denotation.ToString())];
                var letter = order[String.Format("[{0}]->[{1}]", Guids.OrderStructureReference.Links.complexes.ToString(),
                    Guids.ComplexReference.Fields.letter.ToString())];
                // Параметры справочника "Структура заказов"
                var customer = order[Guids.OrderStructureReference.Fields.customer.ToString()];
                var orderNumber = order[Guids.OrderStructureReference.Fields.orderNumber.ToString()];
                var contractNumber = order[Guids.OrderStructureReference.Fields.contractNumber.ToString()];
                var contractor = order[Guids.OrderStructureReference.Fields.contractor.ToString()];
                // Получение списка проверяемых объектов
                var checkedObjects = order.СвязанныеОбъекты[Guids.OrderStructureReference.Links.checkedObjects.ToString()];
                // Список аппаратур в эксплуатации
                var complexes = order.СвязанныеОбъекты[Guids.OrderStructureReference.Links.actualComplexes.ToString()];
                if (isFull && complexes.Count > 0)
                {
                    // Если вид полный, и в справочнике есть аппаратуры по этому заказу
                    foreach (var complex in complexes)
                    {
                        var parameters = new ComplexParameters();
                        parameters.complexName = name;
                        parameters.denotation = denotation;
                        parameters.letter = letter;
                        parameters.customer = customer;
                        parameters.orderNumber = orderNumber;
                        parameters.contractNumber = contractNumber;
                        parameters.contractor = contractor;
                        foreach (var item in checkedObjects)
                        {
                            parameters.checkedObjects.Add(item[Guids.OrderStructureReference.CheckedObjects.name.ToString()]);
                        }
                        parameters.factoryNumber = complex[Guids.ActualComplexReference.Fields.number.ToString()];
                        parameters.dislocationPlace = complex[Guids.ActualComplexReference.Fields.dislocationPlace.ToString()];
                        parameters.deliveryDate = complex[Guids.ActualComplexReference.Fields.deliveryDate.ToString()];
                        parameters.deploymentDate = complex[Guids.ActualComplexReference.Fields.deploymentDate.ToString()];
                        parameters.warranty = complex[Guids.ActualComplexReference.Fields.warranty.ToString()];
                        parameters.resourceExtending = complex[Guids.ActualComplexReference.Fields.resourceExtending.ToString()];
                        parameters.comments = complex[Guids.ActualComplexReference.Fields.comments.ToString()];
                        reportParameters.Add(parameters);
                    }
                }
                else
                {
                    var parameters = new ComplexParameters();
                    parameters.complexName = name;
                    parameters.denotation = denotation;
                    parameters.letter = letter;
                    parameters.customer = customer;
                    parameters.orderNumber = orderNumber;
                    parameters.contractNumber = contractNumber;
                    parameters.contractor = contractor;
                    foreach (var item in checkedObjects)
                    {
                        parameters.checkedObjects.Add(item[Guids.OrderStructureReference.CheckedObjects.name.ToString()]);
                    }
                    reportParameters.Add(parameters);
                }

                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.15));
                cnt++;
                double percent = orders.Count == 0 ? 0 : (double)cnt / orders.Count * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Считываются данные заказа №" + orderNumber, percent))
                    reportParameters.Clear();
            }
            reportParameters.OrderBy(t => t.orderNumber).ThenBy(t => t.denotation).ThenBy(t => t.factoryNumber).ToList();
            macro.ДиалогОжидания.Скрыть();
            return reportParameters;
        }

        public static void FillReport(List<ComplexParameters> reportParameters, bool isFull, MacroProvider macro)
        {
            bool isCatch = false;
            // Создание копии шаблона
            var reference = new FileReference(ServerGateway.Connection);
            var pattern = reference.FindByPath(@"Служебные файлы\Шаблоны отчётов\Отчеты Excel\Excel.xlsx");
            string appDataFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "actualComplexReport.xlsx");
            // Перезапись файла шаблона в случае его отсутствия или изменения
            if (!File.Exists(pattern.LocalPath) || File.GetLastWriteTime(pattern.LocalPath) != pattern.SystemFields.EditDate)
            {
                pattern.GetHeadRevision();
            }
            // Открытие нового документа
            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(pattern.LocalPath);
                File.SetAttributes(pattern.LocalPath, FileAttributes.Normal);
                // Создание заголовка
                ExcelOperations.CreateHeader(isFull, xls);
                // Заполнение таблицы данными
                int row = ExcelOperations.FillTable(reportParameters, isFull, xls, macro);
                if (row > 0)
                {
                    // Создание границ
                    ExcelOperations.CreateBorders(row, isFull, xls);
                    xls.AutoWidth();
                    // Сохранение документа
                    // Сохранение файла
                    if (File.Exists(appDataFileName))
                        File.Delete(appDataFileName);
                    xls.SaveAs(appDataFileName);
                }
                else isCatch = true;
            }
            catch (Exception e)
            {
                macro.Сообщение("Исключительная ситуация!", e.Message);
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
