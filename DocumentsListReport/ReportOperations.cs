using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Files;
using TFlex.DOCs.Model.References.Macros;
using TFlex.DOCs.Model.Stages;
using ReportHelpers;
using System.Diagnostics;

namespace DocumentsListReport
{
    public class ReportOperations
    {
        MacroProvider macro;
        public ReportOperations(MacroProvider macro)
        {
            this.macro = macro;
        }
        public void FillDepartments(Объект user, List<Объект> departments)
        {
            if (user.Тип.ПорожденОт(Guids.UserReference.Types.department) && user.РодительскийОбъект != null &&
                ((ReferenceObject)user.РодительскийОбъект).SystemFields.Guid == Guids.UserReference.Departments.Objects.globus)
            {
                departments.Add(user);
                return;
            }

            foreach (var parent in user.РодительскиеОбъекты)
            {
                FillDepartments(parent, departments);
            }
        }
        // Получение списка подразделений для формы отчета
        public Объект[] GetDepartmentList(Объект user, Объекты userParents, List<Объект> userDepartments, List<Объект> allDepartments)
        {
            bool isLimited = true;
            List<Объект> departmentList = new List<Объект>();
            // Если роли или подразделения сотрудника - это админы, руководство или канцелярия, то ограничений нет. Иначе - только свое подразделение
            if (user.Тип.УникальныйИдентификатор == Guids.UserReference.Types.admin)
                isLimited = false;
            foreach (ReferenceObject item in userParents)
            {
                if (item.SystemFields.Guid == Guids.UserReference.Departments.Objects.chiefs ||
                    item.SystemFields.Guid == Guids.UserReference.Departments.Objects.chancellory)
                    isLimited = false;
            }
            if (isLimited)
                departmentList.Add(userDepartments.FirstOrDefault());
            else
                departmentList.AddRange(allDepartments);
            return departmentList.ToArray();
        }
        // Отбор объектов для исходящих, приказов, распоряжений, ДЗУ
        public Объекты GetObjectList(Guid typeGuid, DateTime startData, DateTime finishData)
        {
            // Условие поиска
            var условия = new List<string>();
            // Документ заданного типа
            условия.Add(String.Format("[Тип] = '{0}'", typeGuid.ToString()));
            // Дата между startData и finishData
            условия.Add(String.Format("[{0}] >= '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), startData));
            условия.Add(String.Format("[{0}] < '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), finishData.AddDays(1)));
            // Для ДЗУ - это рамочный договор
            if (typeGuid == Guids.RegCardsReference.ContractDoc.type)
                условия.Add(String.Format("[{0}] = {1}", Guids.RegCardsReference.ContractDoc.Fields.frameContract.ToString(), true));
            var условие = String.Join(" И ", условия);
            return macro.НайтиОбъекты(String.Format("{0}", Guids.RegCardsReference.refGuid.ToString()), условие);
        }
        // Отбор объектов для СЗ
        public Объекты GetServiceDocList(DateTime startData, DateTime finishData, Объект department, Dictionary<int, bool> typeFlags, Dictionary<int, bool> stageFlags)
        {
            // Условие поиска
            var условия = new List<string>();
            // Документ заданного типа
            условия.Add(String.Format("[Тип] = '{0}'", Guids.RegCardsReference.ServiceDoc.type.ToString()));
            // Дата между startData и finishData
            условия.Add(String.Format("[{0}] >= '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), startData));
            условия.Add(String.Format("[{0}] < '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), finishData.AddDays(1)));
            // Условие со стадией
            List<string> stages = new List<string>();
            foreach (var item in stageFlags)
            {
                Stage stage = null;
                if (item.Key == 0 && item.Value == true) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.closeStage);
                if (item.Key == 1 && item.Value == true) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.processStage);
                if (item.Key == 2 && item.Value == true) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.returnStage);
                if (item.Key == 3 && item.Value == true) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.cancelStage);
                if (stage != null) stages.Add(stage.Guid.ToString());
            }
            условия.Add("[Стадия] входит в список '" + string.Join(",", stages.ToArray()) + "'");
            // Условие по типу документа (входящие, исходящие)
            if (typeFlags[0] == true)
            {
                // Входящие
                условия.Add(String.Format("[{0}]->[Guid] = '{1}'", Guids.RegCardsReference.Links.receiversLink.ToString(), 
                    ((ReferenceObject)department).SystemFields.Guid));
            }
            if (typeFlags[1] == true)
            {
                // Исходящие
                условия.Add(String.Format("[{0}] = '{1}'", Guids.RegCardsReference.Fields.departmentNumber.ToString(),
                    department[Guids.UserReference.Departments.Fields.code.ToString()]));
            }
            if (typeFlags[2] == true)
            {
                условия.Add(String.Format("(([{0}] = '{1}' ИЛИ [{2}]->[Guid] = '{3}'))", Guids.RegCardsReference.Fields.departmentNumber.ToString(),
                    department[Guids.UserReference.Departments.Fields.code.ToString()],
                    Guids.RegCardsReference.Links.receiversLink.ToString(),
                    ((ReferenceObject)department).SystemFields.Guid));
            }
            var условие = String.Join(" И ", условия);
            return macro.НайтиОбъекты(String.Format("{0}", Guids.RegCardsReference.refGuid.ToString()), условие);
        }
        // Отбор объектов для ДЗ
        public Объекты GetReportDocList(DateTime startData, DateTime finishData, Объект department, Dictionary<int, bool> stageFlags, bool incorrect)
        {
            // Условие поиска
            var условия = new List<string>();
            // Документ заданного типа
            условия.Add(String.Format("[Тип] = '{0}'", Guids.RegCardsReference.ReportDoc.type.ToString()));
            // Дата между startData и finishData
            условия.Add(String.Format("[{0}] >= '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), startData));
            условия.Add(String.Format("[{0}] < '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), finishData.AddDays(1)));
            // Условие со стадией (если корректные документы - то набранные стадии, если некорректные - то все кроме Открыто)
            List<string> stages = new List<string>();
            foreach (var item in stageFlags)
            {
                Stage stage = null;
                if (item.Key == 0 && (item.Value == true || incorrect)) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.closeStage);
                if (item.Key == 1 && (item.Value == true || incorrect)) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.processStage);
                if (item.Key == 2 && (item.Value == true || incorrect)) stage = Stage.Find(ServerGateway.Connection, Guids.DocumentStages.returnStage);
                if (stage != null) stages.Add(stage.Guid.ToString());
            }
            условия.Add("[Стадия] входит в список '" + string.Join(",", stages.ToArray()) + "'");
            // Условие по подразделению (только для корректных документов), для некорректных - нет присоединенных файлов
            if (!incorrect)
            {
                условия.Add(String.Format("[{0}] = '{1}'", Guids.RegCardsReference.Fields.departmentNumber.ToString(),
                    department[Guids.UserReference.Departments.Fields.code.ToString()]));
            }
            else
            {
                условия.Add(String.Format("[{0}] не содержит данных", Guids.RegCardsReference.Links.docLink.ToString()));
            }
            var условие = String.Join(" И ", условия);
            return macro.НайтиОбъекты(String.Format("{0}", Guids.RegCardsReference.refGuid.ToString()), условие);
        }
        // Отбор объектов для СЗНО
        public Объекты GetPaymentServiceDocList(DateTime startData, DateTime finishData, Объект department, bool isPaid)
        {
            // Условие поиска
            var условия = new List<string>();
            // Документ заданного типа
            условия.Add(String.Format("[Тип] = '{0}'", Guids.RegCardsReference.PaymentServiceDoc.type.ToString()));
            // Дата между startData и finishData
            условия.Add(String.Format("[{0}] >= '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), startData));
            условия.Add(String.Format("[{0}] < '{1}'", Guids.RegCardsReference.Fields.regDate.ToString(), finishData.AddDays(1)));
            // Условие по подразделению
            условия.Add(String.Format("[{0}] = '{1}'", Guids.RegCardsReference.Fields.departmentNumber.ToString(),
                department[Guids.UserReference.Departments.Fields.code.ToString()]));
            // Если Оплачено - то все, у кого установлен соответствующий флажок, если нет, то все выбранных стадий, у кого он не установлен
            if (isPaid)
            {
                условия.Add(String.Format("[{0}] = true", Guids.RegCardsReference.PaymentServiceDoc.Fields.paid));
            }
            else
            {
                // Список стадий
                List<string> stages = new List<string>
                {
                    Guids.DocumentStages.signedStage.ToString(),
                    Guids.DocumentStages.inArchiveStage.ToString(),
                    Guids.DocumentStages.processStage.ToString(),
                    Guids.DocumentStages.returnStage.ToString(),
                    Guids.DocumentStages.closeStage.ToString()
                };
                условия.Add(String.Format("[{0}] = false", Guids.RegCardsReference.PaymentServiceDoc.Fields.paid));
                условия.Add("[Стадия] входит в список '" + string.Join(",", stages.ToArray()) + "'");
            }

            var условие = String.Join(" И ", условия);
            return macro.НайтиОбъекты(String.Format("{0}", Guids.RegCardsReference.refGuid.ToString()), условие);
        }
        // Создать отчет для исходящих, приказов, распоряжений, ДЗУ
        public void MakeReport(Объекты objects, Guid typeGuid, DateTime startDate, DateTime finishDate, Объект department = null, bool incorrect = false)
        {
            bool isCatch = false;
            // Создание копии шаблона
            var reference = new FileReference(ServerGateway.Connection);
            var pattern = reference.FindByPath(@"Служебные файлы\Шаблоны отчётов\Отчеты Excel\Excel.xlsx");
            string appDataFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Реестр документов предприятия.xlsx");
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
                var rHeaders = new ReportHeaders();
                var rTables = new ReportTables(macro);
                int row = 0;
                // Исходящие
                if (typeGuid == Guids.RegCardsReference.InputDoc.type)
                {
                    row = rHeaders.InputDocHeader(xls, startDate, finishDate);
                    rTables.InputDocTable(xls, row, objects);
                }
                // Приказы
                if (typeGuid == Guids.RegCardsReference.OrderDoc.type)
                {
                    row = rHeaders.OrderDocHeader(xls, startDate, finishDate);
                    rTables.OrderDocTable(xls, row, objects);
                }
                // Распоряжения
                if (typeGuid == Guids.RegCardsReference.DecreeDoc.type)
                {
                    row = rHeaders.DecreeDocHeader(xls, startDate, finishDate);
                    rTables.DecreeDocTable(xls, row, objects);
                }
                // ДЗУ
                if (typeGuid == Guids.RegCardsReference.ContractDoc.type)
                {
                    row = rHeaders.ContractDocHeader(xls, startDate, finishDate);
                    rTables.ContractDocTable(xls, row, objects);
                }
                // СЗ
                if (typeGuid == Guids.RegCardsReference.ServiceDoc.type)
                {
                    row = rHeaders.ServiceDocHeader(xls, startDate, finishDate, department);
                    rTables.ServiceDocTable(xls, row, objects);
                }
                // ДЗ
                if (typeGuid == Guids.RegCardsReference.ReportDoc.type)
                {
                    // incorrect - в данном случае - нет присоединенных файлов
                    row = rHeaders.ReportDocHeader(xls, startDate, finishDate, department, incorrect);
                    rTables.ReportDocTable(xls, row, objects);
                }
                // СЗНО
                if (typeGuid == Guids.RegCardsReference.PaymentServiceDoc.type)
                {
                    // incorrect - в данном случае - оплачено / неоплачено
                    row = rHeaders.PaymentServiceDocHeader(xls, startDate, finishDate, department, incorrect);
                    rTables.PaymentServiceDocTable(xls, row, objects, incorrect);
                }
                // Сохранение файла
                if (File.Exists(appDataFileName))
                    File.Delete(appDataFileName);
                xls.SaveAs(appDataFileName);
            }
            catch (Exception ex)
            {
                macro.Сообщение("Исключительная ситуация!", ex.Message);
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
