using System;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Search;
using System.Collections.Generic;
using TFlex.DOCs.Model.Structure;
using Filter = TFlex.DOCs.Model.Search.Filter;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using TFlex.Reporting;
using static TFlex.DOCs.Model.Licensing.ModuleLicense;
using System.Linq;

namespace ReportIncomingDoc
{
    public class GenerateExcel
    {
        public List<ReferenceObject> GetObject(FormSelection formSelection)
        {
            //Дата начала периода
            DateTime start = new DateTime(formSelection.dtpStart.Value.Year, formSelection.dtpStart.Value.Month, formSelection.dtpStart.Value.Day, 15, 00, 00);
            //Дата конца периода
            DateTime end = new DateTime(formSelection.dtpEnd.Value.Year, formSelection.dtpEnd.Value.Month, formSelection.dtpEnd.Value.Day, 15, 00, 00);
            //Создаем ссылку на справочник
            ReferenceInfo info = ServerGateway.Connection.ReferenceCatalog.Find(Guids.СправочникРКК);
            Reference reference = info.CreateReference();
            //Находим тип «Входящие документы»
            ClassObject classObject = info.Classes.Find(Guids.ВходящиеДокументы);
            //Находим группу параметров ВД
            ParameterGroup parameterGroup = classObject.ParameterGroups.Find(Guids.ГруппаПараметровВД);
            //Создаем фильтр
            Filter filter = new Filter(info);
            //Добавляем условие поиска – «Тип = Входящие документы»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
            //Добавляем условие поиска – «Дата документа >= даты начала периода»
            filter.Terms.AddTerm(parameterGroup[Guids.ДатаПостановкиНаКонтроль], ComparisonOperator.GreaterThanOrEqual, start);
            //Добавляем условие поиска – «Дата документа <= даты конца периода»
            filter.Terms.AddTerm(parameterGroup[Guids.ДатаПостановкиНаКонтроль], ComparisonOperator.LessThanOrEqual, end);
            //Добавляем условие поиска – «Дата документа <= даты конца периода»
            filter.Terms.AddTerm(parameterGroup[Guids.ВхНаКонтроле], ComparisonOperator.Equal, true);
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр
            //Фильтруем полученные объекты по статусу Вх. на контроле = true
            List<ReferenceObject> refObjects = reference.Find(filter);
            return refObjects;
        }

        //Заполняем excel данными объекта
        public int CreateReportFromExcel(Worksheet xlSht, int row, ReferenceObject refObject, int index)
        {
            var name = (refObject[Guids.Резолюция].Value.ToString() != string.Empty ? (refObject[Guids.Резолюция].Value.ToString() + "\r\n") : string.Empty)
                        + refObject[Guids.РегистрационнныйНомер].Value.ToString()
                        + "\r\nот " + Convert.ToDateTime(refObject[Guids.ДатаОт].Value).ToString("d")
                        + "\r\n" + refObject[Guids.Контрагент].Value.ToString()
                        + "\r\n" + refObject[Guids.КраткоеСодержаниеВД].Value.ToString();

            xlSht.Range["A" + row, Type.Missing].Value2 = index;
            xlSht.Range["B" + row.ToString(), Type.Missing].Value2 = name;
            xlSht.Range["C" + row.ToString(), Type.Missing].Value2 = refObject[Guids.Исполнитель].Value.ToString();
            xlSht.Range["D" + row.ToString(), Type.Missing].Value2 = Convert.ToDateTime(refObject[Guids.Срок].Value).ToString("d");
            xlSht.Range["E" + row.ToString(), Type.Missing].Value2 = GetShortNameByFullName(refObject[Guids.Контролер].Value.ToString());
            xlSht.Columns.AutoFit(); //Автосайз колонок в зависимости от содержимого
            row++;
            return row;
        }

        public string GetShortNameByFullName(string fullName)
        {
            //Создаем ссылку на справочник
            ReferenceInfo info = ServerGateway.Connection.ReferenceCatalog.Find(Guids.СправочникГруппыИПользователи);
            Reference reference = info.CreateReference();
            //Находим тип «Сотрудник»
            ClassObject classObject = info.Classes.Find(Guids.ТипСотрудник);
            //Создаем фильтр
            Filter filter = new Filter(info);
            //Добавляем условие поиска – «Тип = Сотрудник»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, classObject);
            //Добавляем условие поиска – «Наименование Сотрудника = Полное имя контролера ВД»
            filter.Terms.AddTerm(Guids.ПолеНаименование.ToString(), ComparisonOperator.Equal, fullName);
            //Возвращаем короткое имя сотрудника, в качестве условия поиска – сформированный фильтр
            return reference.Find(filter).FirstOrDefault()[Guids.ПолеКороткоеИмя].ToString();
        }

        public void ExportToExcel(FormSelection formSelection, FormIndication formIndication, IReportGenerationContext context)
        {
            formIndication.WriteToLog("Загрузка шаблона отчета");
            context.CopyTemplateFile();
            //Статус возникновения исключительной ситуации в процессе формирования отчета
            bool isCatch = false;
            //Путь сохранения отчета
            var filePath = string.Empty;
            formIndication.WriteToLog("Загрузка объектов справочника");
            //Получаем список объектов для формирования отчета
            var refObjects = GetObject(formSelection);
            formIndication.WriteToLog(String.Format("Загружено {0} объектов", refObjects.Count));
            //Запускаем Excel
            formIndication.WriteToLog("Запуск Excel");
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWB; //рабочая книга откуда будем копировать лист
            Worksheet xlSht; //лист Excel, откуда будем копировать данные
            //Открываем в шаблон отчета
            xlWB = xlApp.Workbooks.Open(context.TemplateFilePath); //название файла Excel откуда будем копировать лист
            //Открываем рабочую страницу
            xlSht = xlWB.Worksheets[1]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            //Редактируем заголовок документа
            xlSht.Range["B6", Type.Missing].Value2 = "Докладная записка   " + DateTime.Today.ToString("d") + "   №  ДЗ 906/";
            xlSht.Range["B6", Type.Missing].Font.Bold = true;
            formIndication.WriteToLog("Заполнение отчета данными");
            try
            {
                var row = 9;
                var index = 1;
                foreach (ReferenceObject obj in refObjects)
                {
                    formIndication.WriteToLog("Добавлен Входящий документ - " + obj[Guids.РегистрационнныйНомер].ToString());
                    //Заполняем excel данными объекта
                    row = CreateReportFromExcel(xlSht, row, obj, index);
                    index++;
                }
                formIndication.WriteToLog("Сохранение шаболна - " + "Отчет по входящим на контроле за " + DateTime.Now.ToString("d") + ".xlsx");
                //Путь до сформированного отчета
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + Path.DirectorySeparatorChar + "Отчет по входящим на контроле за " + DateTime.Now.ToString("d") + ".xlsx";
                //Проверяем существование файла с таким наименованием
                if (File.Exists(filePath))
                {
                    File.Delete(filePath); //Если файл уже существует, то удаляем
                }
                xlWB.SaveAs(filePath); //Сохраняем отчет в папку Документы пользователя
            }
            catch (Exception ex)
            {
                MessageBox.Show("Системная информация:\r\n\r\n" + ex.ToString(), "Ошибка при создании отчета!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                isCatch = true;
            }
            finally
            {
                // При исключительной ситуации закрываем xls без сохранения, если все в порядке - выводим на экран
                if (isCatch)
                {
                    xlWB.Close(false); //true - сохранить изменения, false - не сохранять изменения в файле
                    xlApp.Quit(); //закрываем Excel
                }
                else
                {
                    xlWB.Close(true); //true - сохранить изменения, false - не сохранять изменения в файле
                    xlApp.Quit(); //закрываем Excel
                    formIndication.WriteToLog("Открытие шаблона в Excel");
                    Process.Start(filePath);
                }
                GC.Collect();
            }
        }
    }
}