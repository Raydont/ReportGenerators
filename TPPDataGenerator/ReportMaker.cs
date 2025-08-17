using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using TFlex.DOCs.Model.Desktop;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model;
using ReportHelpers;
using System.Globalization;
using TFlex.DOCs.Model.Stages;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.References.Nomenclature;

namespace TPPDataGenerator
{
    public class ReportMaker
    {
        private readonly MacroProvider macro;
        public ReportMaker(MacroProvider macro)
        {
            this.macro = macro;
        }
        // Для нового КД, либо ИИ - установка флагов изменения КД и выгрузки в 1С для структуры заказа (ЭСЗ)
        public void УстановитьПризнакиЭСЗ(List<int> объектыЭСЗ_ID, bool кдИзменена)
        {
            // Только для ЭСЗ, у которых ID > 0
            if (объектыЭСЗ_ID.Where(t => t > 0).Count() == 0)
                return;
            var объектыЭСЗ = ПолучитьСписокЭСЗ(объектыЭСЗ_ID.Where(t => t > 0).ToList());
            var стадияДляРедактирования = Stage.Find(ServerGateway.Connection, Guids.Стадии.ЭСЗ.установкаТПАктуальны);
            foreach (var объектЭСЗ in объектыЭСЗ)
            {
                var текущаяСтадия = объектЭСЗ.SystemFields.Stage.Stage;
                стадияДляРедактирования.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
                объектЭСЗ.BeginChanges();
                if (кдИзменена)
                    объектЭСЗ[Guids.СтруктурыЗаказов.Поля.кдИзменена].Value = true;
                объектЭСЗ[Guids.СтруктурыЗаказов.Поля.выгрузитьв1С].Value = true;
                объектЭСЗ.EndChanges();
                // Запись в лог
                var message = "ТП утвержден";
                if (кдИзменена)
                    message = "КД изменена";
                СоздатьЗаписьвЛогВыгрузки(объектЭСЗ, message);
                // Возврат текущей стадии
                текущаяСтадия.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
            }
        }
        private List<ReferenceObject> ПолучитьСписокЭСЗ(List<int> объектыЭСЗ_ID)
        {
            return macro.НайтиОбъекты(Guids.СтруктурыЗаказов.id.ToString(), "[ID] входит в список '" + string.Join(",", объектыЭСЗ_ID) + "'")
            .Select(t => (ReferenceObject)t).ToList();
        }
        private void СоздатьЗаписьвЛогВыгрузки(ReferenceObject объектЭСЗ, string message)
        {
            var объектЭСЗ_Объект = Объект.CreateInstance(объектЭСЗ, macro.Context);
            объектЭСЗ_Объект.Изменить();
            var логИзменений = объектЭСЗ_Объект.СоздатьОбъектСписка(Guids.ЛогВыгрузки.id.ToString(), Guids.ЛогВыгрузки.Типы.лог.ToString());
            логИзменений[Guids.ЛогВыгрузки.Поля.инициатор.ToString()] = macro.ТекущийПользователь[Guids.ГруппыПользователи.Поля.короткоеИмя.ToString()].ToString()
                        + " (" + message + ")";
            логИзменений[Guids.ЛогВыгрузки.Поля.флагВыгрузитьв1С.ToString()] = true;
            логИзменений.Сохранить();
            объектЭСЗ_Объект.Сохранить();
        }
        public void СоздатьЗаписьВЛогЗапуска(string обозначение, bool этоИзвещение)
        {
            var логЗапуска = macro.СоздатьОбъект(Guids.ЛогЗапусков.id.ToString(), Guids.ЛогЗапусков.Типы.лог.ToString());
            логЗапуска[Guids.ЛогЗапусков.Поля.обозначение.ToString()] = обозначение;
            логЗапуска[Guids.ЛогЗапусков.Поля.этоИзвещение.ToString()] = этоИзвещение;
            логЗапуска.Сохранить();
        }
        // Для нового КД, либо ИИ - установка флагов неактуальности ТП для объектов номенклатуры
        public void УстановитьФлагиНеактуальностиТП(List<ReferenceObject> объектыНоменклатуры)
        {
            foreach (var объектНоменклатуры in объектыНоменклатуры)
            {
                ИзменитьФлагНеактуальностиТПдляЭСИ(объектНоменклатуры, true);
            }
        }
        // Для проверяемых ЭСЗ - установка признаков неактуальности входящих ЭСИ и самого ЭСЗ
        public void УстановкаПризнаковНеактуальностиЭСИиЭСЗ(List<СтрокаОтчета> строкиОтчета)
        {
            var объектыЭСЗ_ID = строкиОтчета.Select(t => t.объектЭСЗ_id).Distinct().ToList();
            foreach (var объектЭСЗ_ID in объектыЭСЗ_ID)
            {
                var объектыЭСИ_ID = строкиОтчета.Where(t => t.объектЭСЗ_id == объектЭСЗ_ID)
                    .Select(t => t.объектЭСИ_id).Distinct().ToList();
                if (объектыЭСИ_ID.Count > 0)
                {
                    try
                    {
                        // Все объекты ЭСИ с неактуальными ТП, входящие в текущий ЭСЗ со снятым флагом "ТП отсутствует или неактуален"
                        var объектыЭСИ = macro.НайтиОбъекты(Guids.Номенклатура.id.ToString(), "[ID] входит в список '" + string.Join(",", объектыЭСИ_ID) + "'")
                            .Select(t => (ReferenceObject)t).Where(t => t[Guids.Номенклатура.Поля.тпНеактуален].GetBoolean() == false).ToList();
                        объектыЭСИ.ForEach(t => ИзменитьФлагНеактуальностиТПдляЭСИ(t, true));
                    }
                    catch
                    { }
                    var объектЭСЗ = (ReferenceObject)macro.НайтиОбъект(Guids.СтруктурыЗаказов.id.ToString(), "[ID] = " + объектЭСЗ_ID);
                    ИзменитьФлагАктуальностиТПдляЭСЗ(объектЭСЗ, false);
                }
            }
        }
        // Для проверяемых ЭСЗ - снятие признаков неактуальности
        public void СнятьПризнакиНеактуальностиЭСЗ(List<int> исходныеЭСЗ_ID, List<СтрокаОтчета> строкиОтчета)
        {
            var неактуальныеЭСЗ_ID = строкиОтчета.Select(t => t.объектЭСЗ_id).Distinct().ToList();
            var актуальныеЭСЗ_ID = исходныеЭСЗ_ID.Where(t => !неактуальныеЭСЗ_ID.Contains(t)).ToList();
            if (актуальныеЭСЗ_ID.Count > 0)
            {
                foreach (var объектЭСЗ_ID in актуальныеЭСЗ_ID)
                {
                    var объектЭСЗ = (ReferenceObject)macro.НайтиОбъект(Guids.СтруктурыЗаказов.id.ToString(), "[ID] = " + объектЭСЗ_ID);
                    ИзменитьФлагАктуальностиТПдляЭСЗ(объектЭСЗ, true);
                }
            }
        }        
        // Для проверяемого ЭСЗ - изменение признака актуальности
        private void ИзменитьФлагАктуальностиТПдляЭСЗ(ReferenceObject объектЭСЗ, bool value)
        {
            /*
             * Если состояние флага отличается от устанавливаемого, то для объекта ЭСЗ
             * устанавливается специализированная стадия, позволяющая любому пользователю редактировать этот объект,
             * затем объект ЭСЗ берется на редактирование, состояние флага изменяется и возвращается предыдущая стадия
             */
            if (объектЭСЗ[Guids.СтруктурыЗаказов.Поля.тпАктуальны].GetBoolean() != value)
            {
                var текущаяСтадия = объектЭСЗ.SystemFields.Stage.Stage;
                var стадияДляРедактирования = Stage.Find(ServerGateway.Connection, Guids.Стадии.ЭСЗ.установкаТПАктуальны);
                стадияДляРедактирования.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
                объектЭСЗ.BeginChanges();
                объектЭСЗ[Guids.СтруктурыЗаказов.Поля.тпАктуальны].Value = value;
                // Если флаг "ТП Актуальны" устанавливается в true, то флаг "Выгрузить в 1С" тоже устанавливается в true
                if (value == true)
                {
                    объектЭСЗ[Guids.СтруктурыЗаказов.Поля.выгрузитьв1С].Value = value;
                    // Запись в лог
                    СоздатьЗаписьвЛогВыгрузки(объектЭСЗ, "ЭСЗ актуализирована");
                }
                объектЭСЗ.EndChanges();
                // Возврат текущей стадии
                текущаяСтадия.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
            }
        }
        private void ИзменитьФлагНеактуальностиТПдляЭСИ(ReferenceObject объектЭСИ, bool value)
        {
            /*
             * Если объект ЭСИ не был взят на редактирование, и состояние флага отличается от устанавливаемого, то для объекта ЭСИ
             * устанавливается специализированная стадия, позволяющая любому пользователю редактировать этот объект,
             * затем объект ЭСИ берется на редактирование, состояние флага изменяется и возвращается предыдущая стадия
             */
            if (!объектЭСИ.IsCheckedOut && объектЭСИ[Guids.Номенклатура.Поля.тпНеактуален].GetBoolean() != value)
            {
                var текущаяСтадияЭСИ = объектЭСИ.SystemFields.Stage.Stage;
                var стадияТПАктуальны = Stage.Find(ServerGateway.Connection, Guids.Стадии.КД.установкаТПАктуальны);
                стадияТПАктуальны.AutomaticChange(new List<ReferenceObject> { объектЭСИ });
                Desktop.CheckOut(объектЭСИ, false);
                объектЭСИ.BeginChanges();
                объектЭСИ[Guids.Номенклатура.Поля.тпНеактуален].Value = value;
                объектЭСИ.EndChanges();
                Desktop.CheckIn(объектЭСИ, "ТП неактуален (" + macro.ТекущийПользователь + ", " + DateTime.Now.ToShortDateString() + ")", false);
                текущаяСтадияЭСИ.AutomaticChange(new List<ReferenceObject> { объектЭСИ });
            }
        }
        // Для проверяемых ЭСЗ - снятие флага "КД изменена"
        public void СнятьФлагиКдИзмененаЭСЗ(List<int> объектыЭСЗ_ID)
        {
            foreach (var объектЭСЗ_ID in объектыЭСЗ_ID)
            {
                var объектЭСЗ = (ReferenceObject)macro.НайтиОбъект(Guids.СтруктурыЗаказов.id.ToString(), "[ID] = " + объектЭСЗ_ID);
                СнятьФлагКдИзмененадляЭСЗ(объектЭСЗ);
            }
        }
        private void СнятьФлагКдИзмененадляЭСЗ(ReferenceObject объектЭСЗ)
        {
            /*
             * Если состояние флага отличается от устанавливаемого, то для объекта ЭСЗ
             * устанавливается специализированная стадия, позволяющая любому пользователю редактировать этот объект,
             * затем объект ЭСЗ берется на редактирование, состояние флага изменяется и возвращается предыдущая стадия
             */
            if (объектЭСЗ[Guids.СтруктурыЗаказов.Поля.кдИзменена].GetBoolean() == true)
            {
                var текущаяСтадия = объектЭСЗ.SystemFields.Stage.Stage;
                var стадияДляРедактирования = Stage.Find(ServerGateway.Connection, Guids.Стадии.ЭСЗ.установкаТПАктуальны);
                стадияДляРедактирования.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
                объектЭСЗ.BeginChanges();
                объектЭСЗ[Guids.СтруктурыЗаказов.Поля.кдИзменена].Value = false;
                объектЭСЗ.EndChanges();
                // Возврат текущей стадии
                текущаяСтадия.AutomaticChange(new List<ReferenceObject> { объектЭСЗ });
            }    
        }
        public void УстановитьФлагиУчтен(Объекты списокОбъектов)
        {
            foreach (var объектСписка in списокОбъектов)
            {
                объектСписка.Изменить();
                объектСписка[Guids.ЭСИсУтвержденнымиТП.Поля.учтен.ToString()] = true;
                объектСписка.Сохранить();
            }
        }
        public void СформироватьОтчет(List<СтрокаОтчета> строкиОтчета, Guid справочник, bool направитьПоПочте)
        {
            bool isCatch = false;
            var справочникФайлы = new FileReference(ServerGateway.Connection);
            var шаблон = справочникФайлы.FindByPath(Settings.путьКШаблону);
            var стадияТПУтвержден = Stage.Find(ServerGateway.Connection, Guids.Стадии.ТП.тпУтвержден);
            var стадияТПОткрыт = Stage.Find(ServerGateway.Connection, Guids.Стадии.ТП.тпОткрыт);

            // Перезапись файла шаблона в случае его отсутствия или изменения
            if (!File.Exists(шаблон.LocalPath) || File.GetLastWriteTime(шаблон.LocalPath) != шаблон.SystemFields.EditDate)
            {
                шаблон.GetHeadRevision();
            }
            string расширениеФайла = Path.GetExtension(шаблон.LocalPath);
            string имяФайлаОтчета = "Мониторинг потребности разработки (актуализации) ТП от " + DateTime.Now.ToShortDateString() + расширениеФайла;
            string путьДоФайлаОтчета = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), имяФайлаОтчета);
            // Заголовок отчета
            var заголовок = ПолучитьЗаголовокОтчетаИлиСообщения(false, справочник);
            Xls xls = new Xls();
            try
            {
                xls.Application.DisplayAlerts = false;
                xls.Open(шаблон.LocalPath);
                File.SetAttributes(шаблон.LocalPath, FileAttributes.Normal);

                // Первый лист - объекты без ТП (или с открытыми ТП)
                var строкиОтчетаОткрыто = строкиОтчета.Where(t => t.объектЭСИ_стадияТП == стадияТПОткрыт.Name || t.объектЭСИ_стадияТП == "ТП отсутствует").ToList();
                ЗаполнитьЛистОтчета(1, "Объекты без ТП", заголовок, строкиОтчетаОткрыто, xls);
                // Второй лист - объекты с ТП в работе (идут по маршруту)
                var строкиОтчетаВРаботе = строкиОтчета.Where(t => !new List<string> { стадияТПОткрыт.Name, стадияТПУтвержден.Name, "ТП отсутствует" }
                .Contains(t.объектЭСИ_стадияТП)).ToList();
                ЗаполнитьЛистОтчета(2, "Объекты с ТП в работе", заголовок, строкиОтчетаВРаботе, xls);
                // Третий лист - объекты с утвержденными ТП (попавшие из ИИ)
                var строкиОтчетаУтверждено = строкиОтчета.Where(t => t.объектЭСИ_стадияТП == стадияТПУтвержден.Name).ToList();
                ЗаполнитьЛистОтчета(3, "Объекты с утвержденными ТП", заголовок, строкиОтчетаУтверждено, xls);

                xls.SelectWorksheet(1);
                // Сохранение файла
                if (File.Exists(путьДоФайлаОтчета))
                    File.Delete(путьДоФайлаОтчета);
                xls.SaveAs(путьДоФайлаОтчета);
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
                    if (направитьПоПочте)
                        ОтправкаСообщения(путьДоФайлаОтчета, справочник);
                    Process.Start(путьДоФайлаОтчета);
                }
            }
        }
        private void ЗаполнитьЛистОтчета(int номерЛиста, string имяЛиста, string заголовок, List<СтрокаОтчета> строкиОтчета, Xls xls)
        {
            xls.SelectWorksheet(номерЛиста);
            ВставитьЗаголовок(заголовок, 1, 1, "Arial", 12, true, xls);
            xls[1, 1, Settings.столбцыКоличество, 1].Merge();
            xls[1, 1].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            xls[1, 1].WrapText = true;
            var строка = СформироватьШапкуТаблицы(3, xls);
            ЗаполнитьТаблицуОтчета(строкиОтчета, строка, xls);
            xls.Worksheet.Name = имяЛиста;
        }
        private int СформироватьШапкуТаблицы(int строка, Xls xls)
        {
            int столбец = 1;
            xls[столбец, строка].Cells.ColumnWidth = 30;
            столбец = ВставитьЗаголовок("Обозначение объекта", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 50;
            столбец = ВставитьЗаголовок("Наименование объекта / Номер заказа", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 30;
            столбец = ВставитьЗаголовок("Тип объекта", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 35;
            столбец = ВставитьЗаголовок("Стадия ТП", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 30;
            столбец = ВставитьЗаголовок("Извещение", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 20;
            столбец = ВставитьЗаголовок("Наличие КД", строка, столбец, "Calibri", 12, false, xls);
            xls[столбец, строка].Cells.ColumnWidth = 35;
            столбец = ВставитьЗаголовок("Примечание ТП", строка, столбец, "Calibri", 12, false, xls);
            // Выделяем границы заголовка
            xls[1, строка, столбец - 1, 1].SetBorders(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlBordersIndex.xlEdgeTop,
                                                                                                          XlBordersIndex.xlEdgeBottom,
                                                                                                          XlBordersIndex.xlEdgeLeft,
                                                                                                          XlBordersIndex.xlEdgeRight,
                                                                                                          XlBordersIndex.xlInsideVertical);
            строка++;
            return строка;
        }
        private string[,] ПолучитьМассивДляВставкиВExcel(List<СтрокаОтчета> строкиОтчета, int строка, Dictionary<int, int> строкиГруппировки)
        {
            var количествоСтрок = строкиОтчета.Select(t => t.объектЭСЗ_id).Distinct().Count() + строкиОтчета.Count;
            var количествоСтолбцов = Settings.столбцыКоличество;
            var массивДанныхОтчета = new string[количествоСтрок, количествоСтолбцов];
            var объектыЭСЗ_ID = строкиОтчета.Select(t => t.объектЭСЗ_id).Distinct().ToList();
            int номерСтроки = 0;
            foreach (var объектЭСЗ_ID in объектыЭСЗ_ID)
            {
                var строкиОтчетаЭСЗ = строкиОтчета.Where(t => t.объектЭСЗ_id == объектЭСЗ_ID).OrderBy(t => t.объектЭСИ_тип).ThenBy(t => t.объектЭСИ_стадияТП)
                    .ThenBy(t => t.объектЭСИ_обозначение).ToList();
                массивДанныхОтчета[номерСтроки, 0] = строкиОтчетаЭСЗ.First().объектЭСЗ_обозначение;
                массивДанныхОтчета[номерСтроки, 1] = строкиОтчетаЭСЗ.First().объектЭСЗ_номерЗаказа;
                массивДанныхОтчета[номерСтроки, 2] = string.Empty;
                массивДанныхОтчета[номерСтроки, 3] = string.Empty;
                массивДанныхОтчета[номерСтроки, 4] = string.Empty;
                массивДанныхОтчета[номерСтроки, 5] = string.Empty;
                массивДанныхОтчета[номерСтроки, 6] = string.Empty;
                номерСтроки++;
                var строкаНачалоГруппы = номерСтроки + строка;
                foreach (var строкаОтчета in строкиОтчетаЭСЗ)
                {
                    массивДанныхОтчета[номерСтроки, 0] = строкаОтчета.объектЭСИ_обозначение;
                    массивДанныхОтчета[номерСтроки, 1] = строкаОтчета.объектЭСИ_наименование;
                    массивДанныхОтчета[номерСтроки, 2] = строкаОтчета.объектЭСИ_тип;
                    массивДанныхОтчета[номерСтроки, 3] = строкаОтчета.объектЭСИ_стадияТП;
                    массивДанныхОтчета[номерСтроки, 4] = строкаОтчета.объектЭСИ_извещение;
                    массивДанныхОтчета[номерСтроки, 5] = строкаОтчета.объектЭСИ_естьКД;
                    массивДанныхОтчета[номерСтроки, 6] = строкаОтчета.объектЭСИ_примечаниеТП;
                    номерСтроки++;
                }
                var строкаКонецГруппы = номерСтроки + строка - 1;
                строкиГруппировки.Add(строкаНачалоГруппы, строкаКонецГруппы);
            }
            return массивДанныхОтчета;
        }
        private void ЗаполнитьТаблицуОтчета(List<СтрокаОтчета> строкиОтчета, int строка, Xls xls)
        {
            var номерСтроки = строка;
            var строкиГруппировки = new Dictionary<int, int>();
            var массивДанныхОтчета = ПолучитьМассивДляВставкиВExcel(строкиОтчета, строка, строкиГруппировки);

            xls[1, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].Value = массивДанныхОтчета;
            xls[1, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            xls[1, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].VerticalAlignment = XlVAlign.xlVAlignTop;
            xls[1, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].BorderTable(XlLineStyle.xlContinuous, XlBorderWeight.xlThin);
            xls[2, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].WrapText = true;
            xls[6, номерСтроки, Settings.столбцыКоличество, массивДанныхОтчета.Length / Settings.столбцыКоличество].WrapText = true;

            foreach (var строкаГруппировки in строкиГруппировки)
            {
                // Группировка
                xls[1, строкаГруппировки.Key, Settings.столбцыКоличество, строкаГруппировки.Value - строкаГруппировки.Key + 1].Rows.Group();
                xls[1, строкаГруппировки.Key, Settings.столбцыКоличество, строкаГруппировки.Value - строкаГруппировки.Key + 1].EntireRow.Hidden = true;
                // Сверху-вниз
                xls.Worksheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
                // Жирный шрифт в заголовке группировки
                xls[1, строкаГруппировки.Key - 1, 2, 1].Font.Bold = true;
            }
        }
        // Запись в ячейку элемента заголовка таблицы
        private int ВставитьЗаголовок(string заголовок, int строка, int столбец, dynamic шрифт, dynamic размерШрифта, bool подчеркивание, Xls xls)
        {
            xls[столбец, строка].Font.Name = шрифт;
            xls[столбец, строка].Font.Size = размерШрифта;
            xls[столбец, строка].Font.Bold = true;
            xls[столбец, строка].Font.Underline = подчеркивание;
            xls[столбец, строка].SetValue(заголовок);
            столбец++;
            return столбец;
        }
        // Отправка сообщения начальнику отдела 60
        private void ОтправкаСообщения(string путьДоФайлаОтчета, Guid справочник)
        {
            /*var отдел911 = macro.НайтиОбъект(Guids.ГруппыПользователи.id.ToString(), "[Guid]", Guids.ГруппыПользователи.Объекты.отдел911);
            var начальник911 = отдел911.ДочерниеОбъекты.Where(t => t["[Guid]"].ToString() == Guids.ГруппыПользователи.Объекты.начальник911.ToString())
                .FirstOrDefault().ДочерниеОбъекты.FirstOrDefault();
            var отдел60 = macro.НайтиОбъект(Guids.ГруппыПользователи.id.ToString(), "[Guid]", Guids.ГруппыПользователи.Объекты.отдел60);
            var начальник60 = отдел60.ДочерниеОбъекты.Where(t => t["[Guid]"].ToString() == Guids.ГруппыПользователи.Объекты.начальник60.ToString())
                .FirstOrDefault().ДочерниеОбъекты.FirstOrDefault();*/
            var бюро603 = macro.НайтиОбъект(Guids.ГруппыПользователи.id.ToString(), "[Guid]", Guids.ГруппыПользователи.Объекты.бюро603);
            var начальник603 = бюро603.ДочерниеОбъекты.Where(t => t["[Guid]"].ToString() == Guids.ГруппыПользователи.Объекты.начальник603.ToString())
                .FirstOrDefault().ДочерниеОбъекты.FirstOrDefault();

            // Заголовок отчета
            var заголовок = ПолучитьЗаголовокОтчетаИлиСообщения(true, справочник);

            var message = new MailMessage(macro.Context.Connection.Mail.DOCsAccount)
            {
                Subject = заголовок,
                Body = "Отчет содержит список объектов ЭСИ, входящих в состав выбранного объекта, у которых ТП отсутствует либо неактуален с группировкой по ЭСЗ, в которые эти объекты входят"
            };
            //var получатель = (User)macro.ТекущийПользователь;
            //var получатель = (User)начальник911;
            //var получатель = (User)начальник60;
            var получатель = (User)начальник603;
            message.To.AddUser(получатель);
            message.Attachments.Add(new FileAttachment(путьДоФайлаОтчета));
            message.Send();
        }
        private string ПолучитьЗаголовокОтчетаИлиСообщения(bool этоПочта, Guid справочник)
        {
            var названиеОтчета = "Потребность в разработке (актуализации) ТП от " + DateTime.Now;
            var обозначениеКорневогоОбъекта = string.Empty;
            if (справочник == Guids.Номенклатура.id)
                обозначениеКорневогоОбъекта = " для объекта " + macro.ТекущийОбъект[Guids.Номенклатура.Поля.обозначение.ToString()].ToString() + " (новая КД)";
            if (справочник == Guids.Извещения.id)
                обозначениеКорневогоОбъекта = " для объекта " + macro.ТекущийОбъект[Guids.Извещения.Поля.обозначение.ToString()].ToString() + " (извещение об изменении)";
            if (справочник == Guids.СтруктурыЗаказов.id)
                обозначениеКорневогоОбъекта = " для объектов ЭСЗ (структура заказа)";
            var заголовок = этоПочта ? "Отчет \"" + названиеОтчета + обозначениеКорневогоОбъекта + "\"" : названиеОтчета + обозначениеКорневогоОбъекта;
            return заголовок;
        }
    }
    public class СтрокаОтчета
    {
        public int объектЭСЗ_id = 0;
        public string объектЭСЗ_номерЗаказа;
        public string объектЭСЗ_обозначение;
        public int объектЭСИ_id = 0;
        public string объектЭСИ_обозначение;
        public string объектЭСИ_наименование;
        public string объектЭСИ_тип;
        public string объектЭСИ_стадияТП;
        public string объектЭСИ_извещение;
        public string объектЭСИ_примечаниеТП;
        public string объектЭСИ_естьКД;
    }
}