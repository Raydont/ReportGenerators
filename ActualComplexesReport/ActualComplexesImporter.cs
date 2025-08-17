using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.References;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using ReportHelpers;

namespace ActualComplexesReport
{
    public static class ActualComplexesImporter
    {
        // Основной метод
        public static void Import(MacroProvider macro)
        {
            string выбранныйФайл = string.Empty;
            var диалог = macro.СоздатьДиалогВвода("Загрузка данных в систему \"Аппаратура и заказы\"");
            var имяФайла = выбранныйФайл;
            диалог.ДобавитьСтроковое("Файл списка заказов", имяФайла);
            диалог.ДобавитьКнопку("...", new Action<string>((имя) =>
            {
                var диалогВыбораФайла = macro.СоздатьДиалогВыбораФайла();
                диалогВыбораФайла.МножественныйВыбор = false;
                диалогВыбораФайла.Фильтр = "Excel (*.xls, *.xlsx)|*.xls;*.xlsx";
                if (диалогВыбораФайла.Показать())
                {
                    имяФайла = Path.ChangeExtension(диалогВыбораФайла.ИмяФайла, Path.GetExtension(диалогВыбораФайла.ИмяФайла).ToLower());
                    выбранныйФайл = имяФайла;
                    диалог["Файл списка заказов"] = имяФайла;
                }
            }), 30);
            диалог.УстановитьРазмер(500, 0);

            if (диалог.Показать())
            {
                var complexes = ReadData(имяФайла, macro);
                if (complexes != null && complexes.Count > 0 && CheckData(complexes, macro))
                {
                    FillReferenceData(complexes, macro);
                    macro.Сообщение("Успех!", "Импорт данных завершен");
                }
                else
                {
                    macro.Сообщение("Внимание!", "Импорт данных не завершен");
                }
            }
        }
        // Чтение данных из файла Эксель
        private static List<ComplexParameters> ReadData(string fileName, MacroProvider macro)
        {
            Xls xls = new Xls();
            xls.Application.DisplayAlerts = false;
            xls.Open(fileName);
            xls.SelectWorksheet(1);

            Range startDO;
            Range finishDO;
            try
            {
                startDO = xls["СтартАЗ"];
                finishDO = xls["ФинишАЗ"];
            }
            catch
            {
                xls.Quit(false);
                macro.Сообщение("Ошибка!", "Формат файла не соответствует требованиям");
                return null;
            }
            int col = 1;
            int row = 2;
            // коллекция - <целое значение из колонки, номер колонки>
            var valueColumns = new Dictionary<int, int>();
            while (true)
            {
                // Читаем строку с нумерацией колонок
                var currentCell = xls[col, row];
                var cellData = currentCell.Text.ToString();
                if (string.IsNullOrEmpty(cellData) && col > 10)
                    break;
                int number;
                if (int.TryParse(cellData, out number))
                    valueColumns[number] = col;
                col++;
            }
            row = 3;
            col = valueColumns[1];
            var complexes = new List<ComplexParameters>();
            ComplexParameters lastComplex = null;
            while (true)
            {
                // Читаем первую ячейку каждой строки данных (начиная с третьей) таблицы до тех пор, пока не дойдем до ячейки finishDO
                var currentCell = xls[col, row];
                if (currentCell.Row == finishDO.Row && currentCell.Column == finishDO.Column)
                    break;
                var cellData = currentCell.Text.ToString();
                if (!string.IsNullOrEmpty(cellData))
                {
                    int rowNumber;
                    if (int.TryParse(cellData.Trim(), out rowNumber))
                    {
                        // Если значение преобразуется в целое, то создается новый комплекс
                        lastComplex = new ComplexParameters();
                        lastComplex.row = row;
                        lastComplex.countRows = 1;
                        complexes.Add(lastComplex);
                    }
                }
                else
                {
                    // Если же значение не заполнено, но при этом создан комплекс, то количество строк, посвященное этому комплексу, увеличивается
                    if (lastComplex != null)
                        lastComplex.countRows++;
                }
                row++;
            }
            macro.ДиалогОжидания.Показать("Идет чтение данных из файла...", true);
            int cnt = 0;
            // Заполнение данных комплекса
            foreach (var complexParams in complexes)
            {
                ReadComplexParams(xls, complexParams, valueColumns);
                cnt++;
                double percent = complexes.Count() == 0 ? 0 : (double)cnt / complexes.Count() * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Считываются данные заказа №" + complexParams.orderNumber, percent))
                    complexes.Clear();
            }
            macro.ДиалогОжидания.Скрыть();
            xls.Quit(false);
            return complexes;
        }
        private static void ReadComplexParams(Xls xls, ComplexParameters complexParams, Dictionary<int, int> valueColumns)
        {
            complexParams.complexName = ExcelOperations.ReadMultilineString(xls, FileString.NameIndex, valueColumns, complexParams);
            complexParams.denotation = ExcelOperations.ReadMultilineString(xls, FileString.DenotationIndex, valueColumns, complexParams);

            string denotationPart = "";
            for (int i = 0; i < complexParams.denotation.Length; i++)
            {
                if (char.IsLetter(complexParams.denotation[i]))
                    denotationPart += complexParams.denotation[i];
                else
                    break;
            }
            if (complexParams.denotation.IndexOf(denotationPart + ".") < 0)
                complexParams.denotation = complexParams.denotation.Replace(denotationPart, denotationPart + ".");
            complexParams.denotation = complexParams.denotation.Replace(" ", "");
            complexParams.orderNumber = ExcelOperations.ReadMultilineString(xls, FileString.OrderNumberIndex, valueColumns, complexParams);
            complexParams.letter = ExcelOperations.ReadMultilineString(xls, FileString.LetterIndex, valueColumns, complexParams);
            complexParams.factoryNumber = ExcelOperations.ReadMultilineString(xls, FileString.FactoryNumberIndex, valueColumns, complexParams);
            complexParams.customer = ExcelOperations.ReadMultilineString(xls, FileString.CustomerIndex, valueColumns, complexParams);
            complexParams.deliveryDate = ExcelOperations.ReadMultilineString(xls, FileString.DeliveryDateIndex, valueColumns, complexParams);
            complexParams.deploymentDate = ExcelOperations.ReadMultilineString(xls, FileString.DeploymentDateIndex, valueColumns, complexParams);
            complexParams.contractNumber = ExcelOperations.ReadMultilineString(xls, FileString.ContractNumberIndex, valueColumns, complexParams);
            complexParams.contractor = ExcelOperations.ReadMultilineString(xls, FileString.ContractorIndex, valueColumns, complexParams);
            complexParams.warranty = ExcelOperations.ReadMultilineString(xls, FileString.WarrantyIndex, valueColumns, complexParams);
            complexParams.comments = ExcelOperations.ReadMultilineString(xls, FileString.CommentsIndex, valueColumns, complexParams, true);
            complexParams.resourceExtending = ExcelOperations.ReadMultilineString(xls, FileString.ResourceExtendingIndex, valueColumns, complexParams);
            complexParams.dislocationPlace = ExcelOperations.ReadMultilineString(xls, FileString.DislocationPlaceIndex, valueColumns, complexParams);
            // списки
            complexParams.checkedObjects = ExcelOperations.ReadList(xls, FileString.CheckedObjectsIndex, valueColumns, complexParams);
            complexParams.contractorContacts = ExcelOperations.ReadList(xls, FileString.ContractorContactIndex, valueColumns, complexParams);
            complexParams.bulletins = ExcelOperations.ReadList(xls, FileString.BulletinsIndex, valueColumns, complexParams);
            complexParams.technicalActs = ExcelOperations.ReadList(xls, FileString.TechnicalActsIndex, valueColumns, complexParams);
        }
        // Проверка данных
        private static bool CheckData(List<ComplexParameters> complexes, MacroProvider macro)
        {
            macro.ДиалогОжидания.Показать("Идет проверка считанных данных...", true);
            int cnt = 0;
            foreach (var complexParams in complexes)
            {
                if (string.IsNullOrEmpty(complexParams.complexName))
                {
                    macro.Сообщение("Ошибка!", "Строка " + complexParams.row + ": не задан обязательный для заполнения параметр - Наименование аппаратуры");
                    macro.ДиалогОжидания.Скрыть();
                    return false;
                }
                if (string.IsNullOrEmpty(complexParams.denotation))
                {
                    macro.Сообщение("Ошибка!", "Строка " + complexParams.row + ": не задан обязательный для заполнения параметр - Обозначение аппаратуры");
                    macro.ДиалогОжидания.Скрыть();
                    return false;
                }
                if (string.IsNullOrEmpty(complexParams.orderNumber))
                {
                    macro.Сообщение("Ошибка!", "Строка " + complexParams.row + ": не задан обязательный для заполнения параметр - Номер заказа");
                    macro.ДиалогОжидания.Скрыть();
                    return false;
                }
                if (string.IsNullOrEmpty(complexParams.customer))
                {
                    macro.Сообщение("Ошибка!", "Строка " + complexParams.row + ": не задан обязательный для заполнения параметр - Потребитель");
                    macro.ДиалогОжидания.Скрыть();
                    return false;
                }
                if (string.IsNullOrEmpty(complexParams.factoryNumber))
                {
                    macro.Сообщение("Ошибка!", "Строка " + complexParams.row + ": не задан обязательный для заполнения параметр - Заводской номер аппаратуры");
                    macro.ДиалогОжидания.Скрыть();
                    return false;
                }
                cnt++;
                double percent = complexes.Count() == 0 ? 0 : (double)cnt / complexes.Count() * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Идет проверка данных заказа №" + complexParams.orderNumber, percent))
                    return false;
            }
            macro.ДиалогОжидания.Скрыть();
            return true;
        }
        // Заполнение данных в справочниках
        private static void FillReferenceData(List<ComplexParameters> complexes, MacroProvider macro)
        {
            int cnt = 0;
            macro.ДиалогОжидания.Показать("Идет запись данных в справочную систему...", true);
            foreach (var complexParams in complexes)
            {
                string условие = string.Format("([{0}] = '{1}' И [{2}] = '{3}')", Guids.ComplexReference.Fields.name.ToString(),
                    complexParams.complexName, 
                    Guids.ComplexReference.Fields.denotation.ToString(),
                    complexParams.denotation);
                var complex = macro.НайтиОбъект(Guids.ComplexReference.refGuid.ToString(), условие);
                // Поиск комплекса в справочнике : взятие на редактирование - если есть, создание - если нет
                if (complex != null)
                    complex.Изменить();
                else
                {
                    complex = macro.СоздатьОбъект(Guids.ComplexReference.refGuid.ToString(), Guids.ComplexReference.type.ToString());
                    complex[Guids.ComplexReference.Fields.name.ToString()] = complexParams.complexName;
                    complex[Guids.ComplexReference.Fields.denotation.ToString()] = complexParams.denotation;
                }
                complex[Guids.ComplexReference.Fields.letter.ToString()] = complexParams.letter;
                complex.Сохранить();
                Desktop.CheckIn((DesktopObject)complex, "Импорт комплекса", false);
                // Поиск заказа в справочнике : взятие на редактирование - если есть, создание - если нет
                условие = string.Format("([{0}] = '{1}' И [{2}]->[ID] = '{3}')", Guids.OrderStructureReference.Fields.orderNumber.ToString(),
                    complexParams.orderNumber, 
                    Guids.OrderStructureReference.Links.complexes.ToString(),
                    ((ReferenceObject)complex).SystemFields.Id);
                var order = macro.НайтиОбъект(Guids.OrderStructureReference.refGuid.ToString(), условие);
                if (order != null)
                    order.Изменить();
                else
                {
                    order = macro.СоздатьОбъект(Guids.OrderStructureReference.refGuid.ToString(), Guids.OrderStructureReference.type.ToString());
                    order[Guids.OrderStructureReference.Fields.orderNumber.ToString()] = complexParams.orderNumber;
                }
                order[Guids.OrderStructureReference.Fields.contractNumber.ToString()] = complexParams.contractNumber;
                order[Guids.OrderStructureReference.Fields.contractor.ToString()] = complexParams.contractor;
                order[Guids.OrderStructureReference.Fields.customer.ToString()] = complexParams.customer;
                order[Guids.OrderStructureReference.Fields.exploitation.ToString()] = true;
                order.Подключить(Guids.OrderStructureReference.Links.complexes.ToString(), complex);
                ((ReferenceObject)order).ClearObjectList(Guids.OrderStructureReference.Links.checkedObjects);
                foreach (var listItemText in complexParams.checkedObjects)
                {
                    Объект listItem = order.СоздатьОбъектСписка(Guids.OrderStructureReference.Links.checkedObjects.ToString(),
                        Guids.OrderStructureReference.CheckedObjects.type.ToString());
                    listItem[Guids.OrderStructureReference.CheckedObjects.name.ToString()] = listItemText;
                    listItem.Сохранить();
                }
                order.Сохранить();
                Desktop.CheckIn((DesktopObject)order, "Импорт заказа", false);
                // Поиск аппаратур с заводскими номерами в справочнике : взятие на редактирование - если есть, создание - если  нет
                условие = string.Format("([{0}] = '{1}' И [{2}]->[ID] = '{3}')", Guids.ActualComplexReference.Fields.number,
                    complexParams.factoryNumber,
                    Guids.ActualComplexReference.Links.order.ToString(),
                    ((ReferenceObject)order).SystemFields.Id);
                var actualComplex = macro.НайтиОбъект(Guids.ActualComplexReference.refGuid.ToString(), условие);
                if (actualComplex != null)
                    actualComplex.Изменить();
                else
                {
                    actualComplex = macro.СоздатьОбъект(Guids.ActualComplexReference.refGuid.ToString(), Guids.ActualComplexReference.type.ToString());
                    actualComplex[Guids.ActualComplexReference.Fields.number.ToString()] = complexParams.factoryNumber;
                    actualComplex.Подключить(Guids.ActualComplexReference.Links.order.ToString(), order);
                }
                actualComplex[Guids.ActualComplexReference.Fields.dislocationPlace.ToString()] = complexParams.dislocationPlace;
                actualComplex[Guids.ActualComplexReference.Fields.deliveryDate.ToString()] = complexParams.deliveryDate;
                actualComplex[Guids.ActualComplexReference.Fields.deploymentDate.ToString()] = complexParams.deploymentDate;
                actualComplex[Guids.ActualComplexReference.Fields.comments.ToString()] = complexParams.comments;
                actualComplex[Guids.ActualComplexReference.Fields.warranty.ToString()] = complexParams.warranty;
                actualComplex[Guids.ActualComplexReference.Fields.resourceExtending.ToString()] = complexParams.resourceExtending;
                ((ReferenceObject)actualComplex).ClearObjectList(Guids.ActualComplexReference.Links.bulletins);
                foreach (var listItemText in complexParams.bulletins)
                {
                    Объект listItem = actualComplex.СоздатьОбъектСписка(Guids.ActualComplexReference.Links.bulletins.ToString(),
                        Guids.ActualComplexReference.Bulletins.type.ToString());
                    listItem[Guids.ActualComplexReference.Bulletins.name.ToString()] = listItemText;
                    listItem.Сохранить();
                }
                 ((ReferenceObject)actualComplex).ClearObjectList(Guids.ActualComplexReference.Links.contractorContacts);
                foreach (var listItemText in complexParams.contractorContacts)
                {
                    Объект listItem = actualComplex.СоздатьОбъектСписка(Guids.ActualComplexReference.Links.contractorContacts.ToString(),
                        Guids.ActualComplexReference.ContractorContacts.type.ToString());
                    listItem[Guids.ActualComplexReference.ContractorContacts.name.ToString()] = listItemText;
                    listItem.Сохранить();
                }
                 ((ReferenceObject)actualComplex).ClearObjectList(Guids.ActualComplexReference.Links.technicalActs);
                foreach (var listItemText in complexParams.technicalActs)
                {
                    Объект listItem = actualComplex.СоздатьОбъектСписка(Guids.ActualComplexReference.Links.technicalActs.ToString(),
                        Guids.ActualComplexReference.TechnicalActs.type.ToString());
                    listItem[Guids.ActualComplexReference.TechnicalActs.name.ToString()] = listItemText;
                    listItem.Сохранить();
                }
                actualComplex.Сохранить();
                Desktop.CheckIn((DesktopObject)actualComplex, "Импорт аппаратуры", false);
                cnt++;
                double percent = complexes.Count() == 0 ? 0 : (double)cnt / complexes.Count() * 100;
                if (!macro.ДиалогОжидания.СледующийШаг("Заполняются данные заказа №" + complexParams.orderNumber, percent))
                    break;
            }
            macro.ДиалогОжидания.Скрыть();
        }
    }
}
