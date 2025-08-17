using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Macros.ObjectModel;
using System.Globalization;

namespace DocumentsListReport
{
    public class DocumentsRegistryReport
    {
        public static void Generate(MacroProvider macro)
        {
            var rOps = new ReportOperations(macro);
            var currentUser = macro.ТекущийПользователь;
            var currentObject = macro.ТекущийОбъект;
            // Поиск подразделения, в которое входит текущий юзер
            List<Объект> userDepartments = new List<Объект>();
            rOps.FillDepartments(currentUser, userDepartments);
            // Получение всех подразделений АО "РКБ Глобус"
            var globus = macro.НайтиОбъект(Guids.UserReference.refGuid.ToString(), String.Format("[Guid] = {0}", Guids.UserReference.Departments.Objects.globus));
            // Все входящие в АО "РКБ Глобус", где тип объектов наследуется от "Производственное подразделение", и не являющиеся руководством предприятия
            var departments = globus.ДочерниеОбъекты.Where(t => t.Тип.ПорожденОт(Guids.UserReference.Types.department) &&
            ((ReferenceObject)t).SystemFields.Guid != Guids.UserReference.Departments.Objects.chiefs).OrderBy(t => t.ToString()).ToList();
            // Получение всех подразделений и ролей текущего пользователя
            var userParents = currentUser.РодительскиеОбъекты;

            var dialog = macro.СоздатьДиалогВвода("Параметры отчета");
            dialog.ДобавитьГруппу("Временные рамки отчета");
            dialog.ДобавитьДату("Начальная дата");
            dialog.ДобавитьДату("Конечная дата");
            dialog["Начальная дата"] = Convert.ToDateTime("01." + DateTime.Today.Month + "." + DateTime.Today.Year);
            dialog["Конечная дата"] = DateTime.Today;
            dialog.ЗначениеИзменено += (имяПоля, староеЗначение, новоеЗначение) =>
            {
                if ((имяПоля == "Начальная дата" || имяПоля == "Конечная дата") && dialog["Начальная дата"] > dialog["Конечная дата"])
                {
                    macro.Сообщение("Ошибка!", "Начальная дата больше конечной\nВведите корректные даты");
                    dialog["Начальная дата"] = dialog["Конечная дата"];
                }
            };
            // Для приказов, исходящих и распоряжений
            if (currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.InputDoc.type ||
                currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.OrderDoc.type ||
                currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.DecreeDoc.type)
            {
                bool rights = false;
                // Если пользователь - админ, то у него есть право на отчет
                if (currentUser.Тип.УникальныйИдентификатор == Guids.UserReference.Types.admin)
                    rights = true;
                // Если пользователь из канцелярии, то у него есть право на отчет
                foreach (ReferenceObject item in userParents)
                {
                    if (item.SystemFields.Guid == Guids.UserReference.Departments.Objects.chancellory)
                        rights = true;
                }
                if (!rights)
                {
                    macro.Сообщение("Внимание!", "У Вас недостаточно прав для формирования отчета");
                    return;
                }

                if (dialog.Показать())
                {
                    Объекты objects = rOps.GetObjectList(currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"]);
                    if (objects.Count == 0)
                    {
                        macro.Сообщение("Внимание!", "По Вашему запросу документы не найдены");
                        return;
                    }
                    rOps.MakeReport(objects, currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"]);
                }
                else
                    macro.Сообщение("Внимание!", "Отменено пользователем");
            }
            // Для ДЗУ
            if (currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.ContractDoc.type)
            {
                if (dialog.Показать())
                {
                    Объекты objects = rOps.GetObjectList(currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"]);
                    if (objects.Count == 0)
                    {
                        macro.Сообщение("Внимание!", "По Вашему запросу документы не найдены");
                        return;
                    }
                    rOps.MakeReport(objects, currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"]);
                }
                else
                    macro.Сообщение("Внимание!", "Отменено пользователем");
            }
            // Для СЗ
            if (currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.ServiceDoc.type)
            {
                dialog.Ширина = 500;
                dialog.ДобавитьГруппу("Отбор по типу документа");
                dialog.ДобавитьФлаг("Входящие");
                dialog.ДобавитьФлаг("Исходящие");
                dialog.ДобавитьФлаг("Все");
                dialog.ДобавитьГруппу("Отбор по подразделению");
                dialog.ДобавитьВыборИзСписка("Подразделение", rOps.GetDepartmentList(currentUser, userParents, userDepartments, departments));
                dialog.ДобавитьГруппу("Отбор по стадии прохождения документа");
                dialog.ДобавитьФлаг("Закрыто");
                dialog.ДобавитьКомментарий("Закрыто", "* Документ принят / получен");
                dialog.ДобавитьФлаг("Согласуется");
                dialog.ДобавитьФлаг("Доработка");
                dialog.ДобавитьФлаг("Отклонен");
                dialog.ДобавитьФлаг("Все стадии");
                dialog.ДобавитьКомментарий("Все стадии", "* Все документы");
                dialog.ДобавитьФлаг("Сброс отбора");
                dialog.ДобавитьКомментарий("Сброс отбора", "* Отменить выделение всех стадий");

                dialog["Исходящие"] = true;
                dialog["Закрыто"] = true;
                dialog["Подразделение"] = userDepartments.FirstOrDefault();
                dialog["Отбор по подразделению"] = "Подразделение-отправитель";

                dialog.ЗначениеИзменено += (имяПоля, староеЗначение, новоеЗначение) =>
                {
                    // Тип документа
                    if (имяПоля == "Входящие" && новоеЗначение == true)
                    {
                        dialog["Исходящие"] = false;
                        dialog["Все"] = false;
                        dialog["Отбор по подразделению"] = "Подразделение-получатель";
                    }
                    if (имяПоля == "Исходящие" && новоеЗначение == true)
                    {
                        dialog["Входящие"] = false;
                        dialog["Все"] = false;
                        dialog["Отбор по подразделению"] = "Подразделение-отправитель";
                    }
                    if (имяПоля == "Все" && новоеЗначение == true)
                    {
                        dialog["Исходящие"] = false;
                        dialog["Входящие"] = false;
                        dialog["Отбор по подразделению"] = "Документы подразделения";
                    }
                    // Стадия прохождения
                    if ((имяПоля == "Закрыто" || имяПоля == "Согласуется" || имяПоля == "Доработка" || имяПоля == "Отклонен") && новоеЗначение == false)
                        dialog["Все стадии"] = false;
                    if (имяПоля == "Все стадии" && новоеЗначение == true)
                    {
                        dialog["Закрыто"] = true;
                        dialog["Доработка"] = true;
                        dialog["Отклонен"] = true;
                        dialog["Согласуется"] = true;
                        dialog["Сброс отбора"] = false;
                    }
                    if (имяПоля == "Сброс отбора" && новоеЗначение == true)
                    {
                        dialog["Закрыто"] = false;
                        dialog["Доработка"] = false;
                        dialog["Отклонен"] = false;
                        dialog["Согласуется"] = false;
                        dialog["Все стадии"] = false;
                    }
                };

                if (dialog.Показать())
                {
                    Dictionary<int, bool> stageFlags = new Dictionary<int, bool>();
                    stageFlags.Add(0, dialog["Закрыто"]);
                    stageFlags.Add(1, dialog["Согласуется"]);
                    stageFlags.Add(2, dialog["Доработка"]);
                    stageFlags.Add(3, dialog["Отклонен"]);
                    Dictionary<int, bool> typeFlags = new Dictionary<int, bool>();
                    typeFlags.Add(0, dialog["Входящие"]);
                    typeFlags.Add(1, dialog["Исходящие"]);
                    typeFlags.Add(2, dialog["Все"]);
                    Объекты objects = rOps.GetServiceDocList(dialog["Начальная дата"], dialog["Конечная дата"], dialog["Подразделение"],
                        typeFlags, stageFlags);
                    if (objects.Count == 0)
                    {
                        macro.Сообщение("Внимание!", "По Вашему запросу документы не найдены");
                        return;
                    }
                    rOps.MakeReport(objects, currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"], dialog["Подразделение"]);
                }
                else
                    macro.Сообщение("Внимание!", "Отменено пользователем");
            }
            // Для ДЗ
            if (currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.ReportDoc.type)
            {
                dialog.Ширина = 500;
                dialog.ДобавитьГруппу("Отбор некорректных документов");
                dialog.ДобавитьФлаг("Некорректные");
                dialog.ДобавитьКомментарий("Некорректные", "* Без присоединенных файлов");
                dialog.ДобавитьГруппу("Отбор по подразделению");
                dialog.ДобавитьВыборИзСписка("Подразделение", rOps.GetDepartmentList(currentUser, userParents, userDepartments, departments));
                dialog.ДобавитьГруппу("Отбор по стадии прохождения документа");
                dialog.ДобавитьФлаг("Закрыто");
                dialog.ДобавитьКомментарий("Закрыто", "* Документ принят / получен");
                dialog.ДобавитьФлаг("Согласуется");
                dialog.ДобавитьФлаг("Доработка");
                dialog.ДобавитьФлаг("Все стадии");
                dialog.ДобавитьКомментарий("Все стадии", "* Все документы");
                dialog.ДобавитьФлаг("Сброс отбора");
                dialog.ДобавитьКомментарий("Сброс отбора", "* Отменить выделение всех стадий");

                dialog["Закрыто"] = true;
                dialog["Подразделение"] = userDepartments.FirstOrDefault();
                dialog["Отбор по подразделению"] = "Подразделение-отправитель";

                dialog.ЗначениеИзменено += (имяПоля, староеЗначение, новоеЗначение) =>
                {
                    // Стадия прохождения
                    if ((имяПоля == "Закрыто" || имяПоля == "Согласуется" || имяПоля == "Доработка") && новоеЗначение == false)
                        dialog["Все стадии"] = false;
                    if (имяПоля == "Все стадии" && новоеЗначение == true)
                    {
                        dialog["Закрыто"] = true;
                        dialog["Доработка"] = true;
                        dialog["Согласуется"] = true;
                        dialog["Сброс отбора"] = false;
                    }
                    if (имяПоля == "Сброс отбора" && новоеЗначение == true)
                    {
                        dialog["Закрыто"] = false;
                        dialog["Доработка"] = false;
                        dialog["Согласуется"] = false;
                        dialog["Все стадии"] = false;
                    }
                    if (имяПоля == "Некорректные")
                    {
                        if (новоеЗначение == true)
                        {
                            dialog.УстановитьВидимостьЭлемента("Отбор по подразделению", false);
                            dialog.УстановитьВидимостьЭлемента("Подразделение", false);
                            dialog.УстановитьВидимостьЭлемента("Отбор по стадии прохождения документа", false);
                            dialog.УстановитьВидимостьЭлемента("Закрыто", false);
                            dialog.УстановитьВидимостьЭлемента("Согласуется", false);
                            dialog.УстановитьВидимостьЭлемента("Доработка", false);
                            dialog.УстановитьВидимостьЭлемента("Все стадии", false);
                            dialog.УстановитьВидимостьЭлемента("Сброс отбора", false);
                        }
                        else
                        {
                            dialog.УстановитьВидимостьЭлемента("Отбор по подразделению", true);
                            dialog.УстановитьВидимостьЭлемента("Подразделение", true);
                            dialog.УстановитьВидимостьЭлемента("Отбор по стадии прохождения документа", true);
                            dialog.УстановитьВидимостьЭлемента("Закрыто", true);
                            dialog.УстановитьВидимостьЭлемента("Согласуется", true);
                            dialog.УстановитьВидимостьЭлемента("Доработка", true);
                            dialog.УстановитьВидимостьЭлемента("Все стадии", true);
                            dialog.УстановитьВидимостьЭлемента("Сброс отбора", true);
                        }
                    }
                };

                if (dialog.Показать())
                {
                    Dictionary<int, bool> stageFlags = new Dictionary<int, bool>();
                    stageFlags.Add(0, dialog["Закрыто"]);
                    stageFlags.Add(1, dialog["Согласуется"]);
                    stageFlags.Add(2, dialog["Доработка"]);
                    Объекты objects = rOps.GetReportDocList(dialog["Начальная дата"], dialog["Конечная дата"], dialog["Подразделение"],
                        stageFlags, dialog["Некорректные"]);
                    if (objects.Count == 0)
                    {
                        macro.Сообщение("Внимание!", "По Вашему запросу документы не найдены");
                        return;
                    }
                    rOps.MakeReport(objects, currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"],
                        dialog["Подразделение"], dialog["Некорректные"]);
                }
                else
                    macro.Сообщение("Внимание!", "Отменено пользователем");
            }
            // Для СЗНО
            if (currentObject.Тип.УникальныйИдентификатор == Guids.RegCardsReference.PaymentServiceDoc.type)
            {
                dialog.Ширина = 500;
                dialog.ДобавитьГруппу("Отбор по подразделению");
                dialog.ДобавитьВыборИзСписка("Подразделение", rOps.GetDepartmentList(currentUser, userParents, userDepartments, departments));
                dialog.ДобавитьГруппу("Отбор по признаку оплачено / не оплачено");
                dialog.ДобавитьФлаг("Оплачено");

                dialog["Подразделение"] = userDepartments.FirstOrDefault();
                dialog["Оплачено"] = false;

                if (dialog.Показать())
                {
                    Объекты objects = rOps.GetPaymentServiceDocList(dialog["Начальная дата"], dialog["Конечная дата"], dialog["Подразделение"], dialog["Оплачено"]);
                    rOps.MakeReport(objects, currentObject.Тип.УникальныйИдентификатор, dialog["Начальная дата"], dialog["Конечная дата"],
                        dialog["Подразделение"], dialog["Оплачено"]);
                }
                else
                    macro.Сообщение("Внимание!", "Отменено пользователем");
            }
        }
    }
}
