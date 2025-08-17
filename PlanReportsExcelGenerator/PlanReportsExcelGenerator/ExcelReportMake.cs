using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;

using TFlex.Reporting;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using ReportHelpers;
using PlanReport.Types;

namespace Globus.PlanReportsExcel
{
    public class ReportMaker : IDisposable
    {
        public void Dispose()
        {
        }
        // Формирование отчета
        public void Make(IReportGenerationContext context)
        {
            bool isShowReport = true;
            Xls xls = new Xls();
            try
            {
                // Создание копии шаблона
                context.CopyTemplateFile();

                // Справочник "Планы и работы"
                var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(PlanReferenceGuids.guid);
                var planReference = referenceInfo.CreateReference();
                // Получение ID объекта, на который получаем отчет
                int planID = Initialize(context);
                if (planID == -1) return;
                // План (график), полученный по ID
                ReferenceObject planObject = planReference.Find(planID);

                if ((planObject.Class.Guid != PlanReferenceGuids.Plan.SimplePlan) && (planObject.Class.Guid != PlanReferenceGuids.Plan.DocDevelopPlan) &&
                    (planObject.Class.Guid != PlanReferenceGuids.Plan.PeriodicTestPlan) && (planObject.Class.Guid != PlanReferenceGuids.Plan.PreliminaryTestPlan) &&
                    (planObject.Class.Guid != PlanReferenceGuids.Plan.ProductionPlan))
                {
                    MessageBox.Show("Для данного типа объекта эта форма отчета не поддерживается", "ОШИБКА!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    isShowReport = false;
                    return;
                }

                MessageBox.Show(planObject.Class.Guid + "");

                // Открытие документа
                xls.Application.DisplayAlerts = false;
                xls.Open(context.ReportFilePath, false);

                ProgressForm p_form;
                int iterations;
                int step;

                // Получение количества колонок в таблице утверждающих и согласующих подписей
                int columnsCount = Convert.ToInt32(planObject[PlanReferenceGuids.Plan.ColInSignBlockTable].Value);

                if (columnsCount == 0)
                {
                    MessageBox.Show("Необходимо указать количество колонок в таблице подписей, отличное от 0");
                    isShowReport = false;
                    return;
                }

                // Получение списка работ выбранного объекта плана (графика)
                List<ReferenceObject> works = WorksReload(planReference, referenceInfo, planObject.Children.RecursiveLoad());
                // Список подписей из объекта плана (графика)
                List<ReferenceObject> signList = planObject.GetObjects(PlanReferenceGuids.Plan.SignBlocks.guid);
                // Определение типа отчета по классу объекта плана(графика)
                WorkTypes workType = getWorkType(planObject.Class.Guid);

                // Экземпляр отчета
                PlanReportExcel engine = new PlanReportExcel();

                p_form = new ProgressForm();
                int maximum = p_form.getProgressBarMax();
                iterations = signList.Count + works.Count * 2 + 1;
                step = (int)(maximum / iterations);
                p_form.progressFormCaption("Идет сбор информации...");
                p_form.Visible = true;

                // Формирование структуры блоков утверждающих и согласующих подписей
                List<SignBlock> signBlockList = fillSignBlockList(signList, p_form, step);
                // Получение списка нижнего блока подписей
                List<SignBlock> footerSignList = new List<SignBlock>();
                footerSignList = fillSignFooterList(works, p_form, step);

                // Получение заголовка плана (графика)
                WorkHeader workheader = fillWorkheader(planObject, p_form, step);

                // Формирование отчета
                switch (workType)
                {
                    case WorkTypes.SimplePlan:
                        // Формирование списка параметров отчета
                        List<WorkParameters> workList = fillWorkList(works, p_form, step);
                        p_form.Close();
                        iterations = signBlockList.Count + workList.Count + footerSignList.Count + 2;
                        step = (int)(maximum / iterations);
                        // Передача данных в диалоговую форму (формирование таблицы Excel идет в методе формы)
                        ReportParametersForm rpForm = new ReportParametersForm();
                        rpForm.engine = engine;
                        rpForm.workType = workType;
                        rpForm.signBlockList = signBlockList;
                        rpForm.footerSignList = footerSignList;
                        rpForm.workHeader = workheader;
                        rpForm.workList = workList;
                        rpForm.xls = xls;
                        rpForm.p_form = new ProgressForm();
                        rpForm.step = step;
                        rpForm.columnsCount = columnsCount;
                        rpForm.ShowDialog();
                        break;
                    case WorkTypes.DocDevelopPlan:
                        // Формирование списка параметров отчета
                        List<WorkDocDevelopParameters> workDocDevelopList = fillDocDevelopList(works, p_form, step);
                        p_form.Close();
                        iterations = signBlockList.Count + workDocDevelopList.Count + footerSignList.Count + 2;
                        step = (int)(maximum / iterations);
                        p_form = new ProgressForm();
                        p_form.progressFormCaption("Идет формирование отчета...");
                        p_form.Visible = true;
                        // Формирование таблицы Excel
                        engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workheader, WorkTypes.DocDevelopPlan, workDocDevelopList, xls, p_form, step);
                        p_form.Close();
                        break;
                    case WorkTypes.PeriodicTestPlan:
                        // Формирование списка параметров отчета
                        List<WorkPeriodicTestParameters> periodicList = fillPeriodicTestList(works, p_form, step);
                        p_form.Close();
                        iterations = signBlockList.Count + periodicList.Count + footerSignList.Count + 2;
                        step = (int)(maximum / iterations);
                        p_form = new ProgressForm();
                        p_form.progressFormCaption("Идет формирование отчета...");
                        p_form.Visible = true;
                        // Формирование таблицы Excel
                        engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workheader, WorkTypes.PeriodicTestPlan, periodicList, xls, p_form, step);
                        p_form.Close();
                        break;
                    case WorkTypes.PreliminaryTestPlan:
                        // Формирование списка параметров отчета
                        List<WorkPreliminaryTestParameters> preliminaryList = fillPreliminaryTestList(works, p_form, step);
                        p_form.Close();
                        iterations = signBlockList.Count + preliminaryList.Count + footerSignList.Count + 2;
                        step = (int)(maximum / iterations);
                        p_form = new ProgressForm();
                        p_form.progressFormCaption("Идет формирование отчета...");
                        p_form.Visible = true;
                        // Формирование таблицы Excel
                        engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workheader, WorkTypes.PreliminaryTestPlan, preliminaryList, xls, p_form, step);
                        p_form.Close();
                        break;
                    case WorkTypes.ProductionPlan:
                        // Формирование списка параметров отчета
                        List<WorkProductionParameters> productionList = fillProductionList(works, p_form, step);
                        p_form.Close();
                        iterations = signBlockList.Count + productionList.Count + footerSignList.Count + 2;
                        step = (int)(maximum / iterations);
                        p_form = new ProgressForm();
                        p_form.progressFormCaption("Идет формирование отчета...");
                        p_form.Visible = true;
                        // Формирование таблицы Excel
                        engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workheader, WorkTypes.ProductionPlan, productionList, xls, p_form, step);
                        p_form.Close();
                        break;
                }

                // Сохранение документа
                xls.Save();
            }
            finally
            {
                // Закрытие документа, вывод файла Excel
                xls.Quit(true);
                if (isShowReport)
                    Process.Start(context.ReportFilePath);
            }
        }
        static int Initialize(IReportGenerationContext context)
        {
            // --------------------------------------------------------------------------
            // Инициализация
            // --------------------------------------------------------------------------

            int documentID;

            // Получаем ID выделенного в интерфейсе T-FLEX DOCs объекта
            if (context.ObjectsInfo.Count == 1) documentID = context.ObjectsInfo[0].ObjectId;
            else
            {
                MessageBox.Show("Выделите объект для формирования отчета");
                return -1;
            }

            return documentID;
        }

        // Определение типа отчета по классу выбранного объекта
        private WorkTypes getWorkType(Guid classGuid)
        {
            if (classGuid == PlanReferenceGuids.Plan.SimplePlan)
                return WorkTypes.SimplePlan;
            if (classGuid == PlanReferenceGuids.Plan.DocDevelopPlan)
                return WorkTypes.DocDevelopPlan;
            if (classGuid == PlanReferenceGuids.Plan.PeriodicTestPlan)
                return WorkTypes.PeriodicTestPlan;
            if (classGuid == PlanReferenceGuids.Plan.PreliminaryTestPlan)
                return WorkTypes.PreliminaryTestPlan;
            if (classGuid == PlanReferenceGuids.Plan.ProductionPlan)
                return WorkTypes.ProductionPlan;
            throw new Exception("Неподдерживаемый тип работы");
        }
        // Заполнение списка подписей верхней части документа (в зависимости от позиции блока и количества колонок)
        private List<SignBlock> fillSignBlockList(List<ReferenceObject> signBlockList, ProgressForm p_form, int step)
        {
            List<SignBlock> list = new List<SignBlock>();

            foreach (var signBlock in signBlockList)
            {
                var block = new SignBlock();
                // Тип подписи (Согласующая, Утверждающая)
                int signType = Convert.ToInt32(signBlock[PlanReferenceGuids.Plan.SignBlocks.SignBlockType].Value);
                switch (signType)
                {
                    case 0: block.signType = SignType.Agreed; break;
                    case 1: block.signType = SignType.Confirmed; break;
                }
                // Должность подписывающего
                int departmentType = Convert.ToInt32(signBlock[PlanReferenceGuids.Plan.SignBlocks.SignBlockPost].Value);
                switch (departmentType)
                {
                    case 0: block.departmentType = DepartmentType.GeneralDirector; break;
                    case 1: block.departmentType = DepartmentType.ChiefDesigner; break;
                    case 2: block.departmentType = DepartmentType.ChiefEngineer; break;
                    case 3: block.departmentType = DepartmentType.DeputyGeneralDirector; break;
                    case 4: block.departmentType = DepartmentType.VPMO; break;
                }
                // ФИО подписывающего
                block.fio = signBlock[PlanReferenceGuids.Plan.SignBlocks.FIO].Value.ToString();
                // Позиция блока в таблице
                block.position = Convert.ToInt32(signBlock[PlanReferenceGuids.Plan.SignBlocks.SignBlockPos]);
                list.Add(block);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Заполнение списка подписей нижней части документа
        private List<SignBlock> fillSignFooterList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<SignBlock> list = new List<SignBlock>();

            List<ReferenceObject> depExecutors = getDepartments(works, p_form, step);

            // Для каждого подразделения-исполнителя определяем тип, код подразделения и его начальника
            foreach (var depExecutor in depExecutors)
            {
                SignBlock block = new SignBlock();
                
                // Код подразделения
                if (depExecutor[GroupsAndUsersReferenceGuids.Code].Value != null)
                    block.departmentNumber = depExecutor[GroupsAndUsersReferenceGuids.Code].Value.ToString();
                else
                    depExecutor.ToString();
               
                // Тип подразделения
                switch (depExecutor.Class.Id)
                {
                    case 9:
                         // Цех
                         block.departmentType = DepartmentType.Shop;
                         break;
                    case 12:
                         // Бюро
                         block.departmentType = DepartmentType.Bureau;
                         break;
                    case 19:
                         // Отдел
                         block.departmentType = DepartmentType.Depart;
                         break;
                    case 23:
                         // Подразделение
                         block.departmentType = DepartmentType.Division;
                         break;
                }
                
                // Имя начальника подразделения
                block.fio = GetBossName(depExecutor);
                
                list.Add(block);
            }

            return list;
        }
        // Заполнение параметров заголовка плана (графика)
        private WorkHeader fillWorkheader(ReferenceObject planObject, ProgressForm p_form, int step)
        {
            WorkHeader header = new WorkHeader();

            header.GraphName = planObject.ToString();

            p_form.progressBarParam(step);

            return header;
        }
        // Заполнение параметров отчета "Простой план"
        private List<WorkParameters> fillWorkList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<WorkParameters> list = new List<WorkParameters>();

            works = works.OrderBy(t => t[PlanReferenceGuids.Work.WorkCode].GetString()).ToList();

            foreach (var item in works)
            {
                WorkParameters work = new WorkParameters();

                if (item.Class.IsInherit(PlanReferenceGuids.Header.ParentGuid))
                {
                    work.IsHeader = true;
                    // Название работы
                    work.WorkSummary = item[PlanReferenceGuids.Header.WorkName].Value.ToString();
                }
                else
                {
                    work.IsHeader = false;
                    // Пункт плана                    
                    work.PlanItem = item[PlanReferenceGuids.Work.WorkNumber].Value.ToString();
                    // Название работы
                    work.WorkSummary = item[PlanReferenceGuids.Work.WorkName].Value.ToString();
                    // Срок исполнения
                    if ((DateTime)item[PlanReferenceGuids.Work.Deadline] != null)
                        work.Deadline = ((DateTime)item[PlanReferenceGuids.Work.Deadline].Value).ToShortDateString();
                    // Примечание
                    work.Comment = item[PlanReferenceGuids.Work.Comment].Value.ToString();
                    #region Заполнение списка исполнителей работы
                    // Для каждой работы считать исполнение заданий
                    List<ReferenceObject> workExecutions = item.GetObjects(PlanReferenceGuids.Work.WorkExecutionLink.guid);
                    List<ReferenceObject> depExecutors = new List<ReferenceObject>();
                    // Заполнение листа исполнителей
                    foreach (var execution in workExecutions)
                    {
                        ReferenceObject depExecutor = execution.GetObject(PlanReferenceGuids.Work.WorkExecutionLink.ExecutorDepartmentLink);
                        depExecutors.Add(depExecutor);
                    }
                    // Удаление повторяющихся записей
                    depExecutors = depExecutors.Distinct().ToList();
                    // Заполнение строки исполнителей
                    string Executors = String.Empty;
                    foreach (var exec in depExecutors)
                    {
                        Executors = Executors + exec + "\n";
                    }
                    #endregion
                    // Исполнители
                    work.Executor = Executors;
                }

                list.Add(work);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Заполнение параметров отчета "График разработки КД"
        private List<WorkDocDevelopParameters> fillDocDevelopList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<WorkDocDevelopParameters> list = new List<WorkDocDevelopParameters>();

            works = works.OrderBy(t => t[PlanReferenceGuids.Work.WorkCode].GetString()).ToList();

            foreach (var item in works)
            {
                WorkDocDevelopParameters work = new WorkDocDevelopParameters();

                if (item.Class.Guid == PlanReferenceGuids.Header.DocDevelopPlanHeader)
                {
                    work.IsHeader = true;
                    // Наименование устройства / документа
                    work.DeviceName = item[PlanReferenceGuids.Header.WorkName].Value.ToString();
                }
                else
                {
                    work.IsHeader = false;
                    // Пункт структуры
                    work.StructItem = item[PlanReferenceGuids.Work.WorkNumber].Value.ToString();
                    // Наименование устройства / документа
                    work.DeviceName = item[PlanReferenceGuids.Work.WorkName].Value.ToString();
                    // Получение списка исполнителей
                    List<ReferenceObject> executors = item.GetObjects(PlanReferenceGuids.Work.ExecutorsLink);
                    string executorsList = String.Empty;
                    foreach (var executor in executors)
                    {
                        executorsList = executorsList + executor[GroupsAndUsersReferenceGuids.ShortName].Value.ToString() + "\n";
                    }
                    // Отдел - разработчик
                    work.Developer = executorsList;
                    // Примечание
                    work.Comments = item[PlanReferenceGuids.Work.Comment].Value.ToString();
                    // Получение всех подработ
                    List<ReferenceObject> subWorks = item.GetObjects(PlanReferenceGuids.Work.SubWorksLink);
                    foreach (var subWork in subWorks)
                    {
                        int subWorkType = subWork[PlanReferenceGuids.Work.SubWork.SubWorkType].GetInt32();
                        #region Заполнение параметров подработ
                        switch (subWorkType)
                        {
                            case 1:
                                // Утвержденное ТЗ выдано
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.TechTaskConfirmDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 2:
                                // ТЗ в смежные подразделения
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.TechTaskStatementDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 3:
                                // Э3 разработано, КД выдано
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.StructDevelopDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 4:
                                // СЗ об уточнении структуры
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.ServiceDocTo906Date = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 5:
                                // ИД на ПП выданы
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.InitialDataDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 6:
                                // КД в отд.20 выдана
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.EngineeringDocDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 7:
                                // Комплект КД в БТД сдан
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.EngineeringDocComplectDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 8:
                                // Текстовая КД разработана
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.TextEngineeringDocDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                        }
                        #endregion
                    }
                }

                list.Add(work);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Заполнение параметров отчета "График проведения периодических испытаний"
        private List<WorkPeriodicTestParameters> fillPeriodicTestList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<WorkPeriodicTestParameters> list = new List<WorkPeriodicTestParameters>();

            works = works.OrderBy(t => t[PlanReferenceGuids.Work.WorkCode].GetString()).ToList();

            foreach (var item in works)
            {
                WorkPeriodicTestParameters work = new WorkPeriodicTestParameters();

                if (item.Class.Guid == PlanReferenceGuids.Header.PeriodicTestPlanHeader)
                {
                    work.IsHeader = true;
                    // Наименование устройства
                    work.DeviceName = item[PlanReferenceGuids.Header.WorkName].Value.ToString();
                }
                else
                {
                    work.IsHeader = false;
                    // Пункт плана
                    work.StructItem = item[PlanReferenceGuids.Work.WorkNumber].Value.ToString();
                    // Наименование устройства
                    work.DeviceName = item[PlanReferenceGuids.Work.WorkName].Value.ToString();
                    // Акт ПИ
                    work.ActPT = item[PlanReferenceGuids.Work.PeriodicTestWork.ActPT].Value.ToString();
                    // Дата окончания действия акта ПИ
                    if (item[PlanReferenceGuids.Work.PeriodicTestWork.ActFinishDate].Value != null)
                        work.ActPTFinishDate = ((DateTime)item[PlanReferenceGuids.Work.PeriodicTestWork.ActFinishDate].Value).ToShortDateString();
                    // График изготовления
                    work.ProductionGraph = item[PlanReferenceGuids.Work.PeriodicTestWork.ProductionGraph].Value.ToString();
                    // Получение всех подработ
                    List<ReferenceObject> subWorks = item.GetObjects(PlanReferenceGuids.Work.SubWorksLink);
                    foreach (var subWork in subWorks)
                    {
                        int subWorkType = subWork[PlanReferenceGuids.Work.SubWork.SubWorkType].GetInt32();
                        #region Заполнение параметров подработ
                        switch (subWorkType)
                        {
                            case 1:
                                // Срок изготовления
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.ProductionDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 2:
                                // Срок проведения ПИ
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.PTDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                        }
                        #endregion
                    }
                }

                list.Add(work);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Заполнение параметров отчета "График проведения предварительных испытаний"
        private List<WorkPreliminaryTestParameters> fillPreliminaryTestList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<WorkPreliminaryTestParameters> list = new List<WorkPreliminaryTestParameters>();

            works = works.OrderBy(t => t[PlanReferenceGuids.Work.WorkCode].GetString()).ToList();

            foreach (var item in works)
            {
                WorkPreliminaryTestParameters work = new WorkPreliminaryTestParameters();

                if (item.Class.Guid == PlanReferenceGuids.Header.PreliminaryTestPlanHeader)
                {
                    work.IsHeader = true;
                    // Краткое содержание пункта программы
                    work.PlanItemName = item[PlanReferenceGuids.Header.WorkName].Value.ToString();
                }
                else
                {
                    work.IsHeader = false;
                    // Пункт программы
                    work.PlanItem = item[PlanReferenceGuids.Work.WorkNumber].Value.ToString();
                    // Краткое содержание пункта программы
                    work.PlanItemName = item[PlanReferenceGuids.Work.WorkName].Value.ToString();
                    // Получение списка исполнителей
                    List<ReferenceObject> executors = item.GetObjects(PlanReferenceGuids.Work.ExecutorsLink);
                    string executorsList = String.Empty;
                    foreach (var executor in executors)
                    {
                        executorsList = executorsList + executor + "\n";
                    }
                    // Исполнитель
                    work.Executor = executorsList;
                    // Примечание
                    work.Comments = item[PlanReferenceGuids.Work.Comment].Value.ToString();
                    // Получение всех подработ
                    List<ReferenceObject> subWorks = item.GetObjects(PlanReferenceGuids.Work.SubWorksLink);
                    foreach (var subWork in subWorks)
                    {
                        int subWorkType = subWork[PlanReferenceGuids.Work.SubWork.SubWorkType].GetInt32();
                        #region Заполнение параметров подработ
                        switch (subWorkType)
                        {
                            case 1:
                                // Методика разработана
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.MethodologyDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 2:
                                // Испытания проведены
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.TestDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                            case 3:
                                // Протокол согласован
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.ProtocolConfirmDate = ((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value).ToShortDateString();
                                break;
                        }
                        #endregion
                    }
                }

                list.Add(work);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Заполнение параметров отчета "График изготовления"
        private List<WorkProductionParameters> fillProductionList(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<WorkProductionParameters> list = new List<WorkProductionParameters>();

            works = works.OrderBy(t => t.SystemFields.Id).ToList();

            foreach (var item in works)
            {
                WorkProductionParameters work = new WorkProductionParameters();
                if (item.Class.Guid == PlanReferenceGuids.Header.ProductionPlanHeader)
                {
                    work.IsHeader = true;
                    // Наименование и шифр
                    work.Name = item[PlanReferenceGuids.Header.WorkName].Value.ToString();
                }
                else
                {
                    work.IsHeader = false;
                    // Пункт программы
                    work.PlanItem = item[PlanReferenceGuids.Work.WorkNumber].Value.ToString();
                    // Наименование и шифр
                    work.Name = item[PlanReferenceGuids.Work.WorkName].Value.ToString();
                    // Децимальный номер
                    string[] denotation = item[PlanReferenceGuids.Work.ProductionWork.DecimalNum].GetString().Split('.');
                    string denotationString = String.Empty;
                    for (int i = 0; i < denotation.Count(); i++)
                        denotationString = denotationString + denotation[i] + ".\n";
                    work.Denotation = denotationString.Remove(denotationString.Count()-2,1);
                    // Количество
                    if (item[PlanReferenceGuids.Work.ProductionWork.Amount].Value.ToString() != "0")
                        work.Amount = item[PlanReferenceGuids.Work.ProductionWork.Amount].Value.ToString();
                    else
                        work.Amount = String.Empty;
                    // Получение разработчика устройства (документа)
                    ReferenceObject executor = item.GetObject(PlanReferenceGuids.Work.ProductionWork.Executor);
                    work.Executor = executor[GroupsAndUsersReferenceGuids.ShortName].Value.ToString();
                    // Номер
                    work.DocNumber = item[PlanReferenceGuids.Work.ProductionWork.DocNumber].Value.ToString();
                    // Примечание
                    work.Comment = item[PlanReferenceGuids.Work.Comment].Value.ToString();
                    // Получение всех подработ
                    List<ReferenceObject> subWorks = item.GetObjects(PlanReferenceGuids.Work.SubWorksLink);
                    foreach (var subWork in subWorks)
                    {
                        int subWorkType = subWork[PlanReferenceGuids.Work.SubWork.SubWorkType].GetInt32();
                        #region Заполнение параметров подработ 
                        switch (subWorkType)
                        {
                            case 1:
                                // Дата
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.StartDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 2:
                                // Н.з. в отд.60
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Bureau219Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 3:
                                // Технологич.обр-ка н.-з.
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Dep60Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 4:
                                // Нормирование
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Bureau904Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 5:
                                // Материалы
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.MaterialsDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 6:
                                // КРЭ
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.OthersDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 7:
                                // 800
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Sector800Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 8:
                                // 850
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Sector850Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 9:
                                // 750
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.Sector750Date = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 10:
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                {
                                    // Настройка / Проверка
                                    work.SettingDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                    string settingExecutors = String.Empty;
                                    foreach (var setExec in subWork.GetObjects(PlanReferenceGuids.Work.SubWork.ExecutorsLink))
                                    {
                                        settingExecutors = settingExecutors + setExec[GroupsAndUsersReferenceGuids.ShortName].Value.ToString() + "\n";
                                    }
                                    // Исполнитель настройки
                                    work.SettingExecutor = settingExecutors;
                                }
                                break;
                            case 11:
                                // Дефектировка
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                    work.DefectiveDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, false);
                                break;
                            case 12:
                                if (subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value != null)
                                {
                                    // Сдача
                                    work.DeliveryDate = getShortDateFormat((DateTime)subWork[PlanReferenceGuids.Work.SubWork.Deadline].Value, true);
                                    string deliveryExecutors = String.Empty;
                                    foreach (var setExec in subWork.GetObjects(PlanReferenceGuids.Work.SubWork.ExecutorsLink))
                                    {
                                        deliveryExecutors = deliveryExecutors + setExec[GroupsAndUsersReferenceGuids.ShortName].Value.ToString() + "\n";
                                    }
                                    // Исполнитель сдачи
                                    work.DeliveryExecutor = deliveryExecutors;
                                }
                                break;
                        }
                        #endregion
                    }
                }

                list.Add(work);
                p_form.progressBarParam(step);
            }

            return list;
        }
        // Приведение даты к укороченной форме
        private string getShortDateFormat(DateTime dTime, bool withYear)
        {
            string shortDateString;
            string day;
            string month;
            string year;

            if (withYear)
            {
                day = dTime.Day.ToString();
                if (day.Length == 1)
                    day = "0" + day;
                month = dTime.Month.ToString();
                if (month.Length == 1)
                    month = "0" + month;
                year = dTime.Year.ToString().Replace("20", String.Empty);
                shortDateString = day + "." + month + "." + year;
            }
            else
            {
                day = dTime.Day.ToString();
                if (day.Length == 1)
                    day = "0" + day;
                month = dTime.Month.ToString();
                if (month.Length == 1)
                    month = "0" + month;
                shortDateString = day + "." + month;
            }

            return shortDateString;
        }
        // Получение списка подразделений-исполнителей
        private List<ReferenceObject> getDepartments(List<ReferenceObject> works, ProgressForm p_form, int step)
        {
            List<ReferenceObject> list = new List<ReferenceObject>();
            List<int> workIDs = new List<int>();
            List<ReferenceObject> executors = new List<ReferenceObject>();

            // Составляем список ID заданий по связи с работами
            List<int> tasksIds = new List<int>();
            foreach (var work in works)
            {
                tasksIds.AddRange(work.GetObjects(PlanReferenceGuids.Work.WorkExecutionLink.guid).Select(t => t.SystemFields.Id));
                p_form.progressBarParam(step);
            }

            // Составляем список заданий из справочника с фильтром по списку ID
            var workExecutionReferenceInfo = ReferenceCatalog.FindReference(TaskReferenceGuids.guid);
            var workExecutionReference = workExecutionReferenceInfo.CreateReference();
            workExecutionReference.LoadSettings.AddRelation(TaskReferenceGuids.ExecutorDepartmentLink.guid).AddAllParameters();
            var filter = new Filter(workExecutionReferenceInfo);
            filter.Terms.AddTerm("ID", ComparisonOperator.IsOneOf, tasksIds);
            var tasks = workExecutionReference.Find(filter);
            // Получаем список исполнителей по связи
            executors.AddRange(tasks.Select(t => t.GetObject(TaskReferenceGuids.ExecutorDepartmentLink.guid)).Distinct().ToList());

            foreach (var executor in executors)
            {
                // Если подразделение-исполнитель входит в ОАО "РКБ Глобус", то записываем его в выходной список, иначе записываем его родителя
                if (executor.Parent.SystemFields.Id == 87)
                    list.Add(executor);
                else
                    list.Add(executor.Parent);
            }

            list = list.Distinct().ToList();

            return list;
        }
        // Нахождение начальника подразделения
        private string GetBossName(ReferenceObject depExecutor)
        {
            string bossFIO = String.Empty;
            // Имя начальника подразделения
            foreach (var child in depExecutor.Children)
            {
                if (child[GroupsAndUsersReferenceGuids.Name].GetString() == "Начальник")
                {
                    ReferenceObject boss = child.Children.FirstOrDefault();
                    if (boss[GroupsAndUsersReferenceGuids.UserShortName].Value != null)
                        bossFIO = boss[GroupsAndUsersReferenceGuids.UserShortName].Value.ToString();
                    else
                        bossFIO = boss.ToString();
                }
            }
            return bossFIO;
        }
        // Перезагрузка списка работ
        private List<ReferenceObject> WorksReload(Reference workReference, ReferenceInfo workReferenceInfo, List<ReferenceObject> works)
        {
            List<ReferenceObject> list = new List<ReferenceObject>();
            List<int> workIDs = new List<int>();

            // Записываем ID работ в список
            foreach (var work in works)
            {
                if (work.Class.IsInherit(PlanReferenceGuids.Work.guid) || work.Class.IsInherit(PlanReferenceGuids.Header.ParentGuid))
                    workIDs.Add(work.SystemFields.Id);
            }

            // Еще раз считываем работы со всеми параметрами из справочника с фильтром по списку ID
            workReference.LoadSettings.AddRelation(PlanReferenceGuids.Work.WorkExecutionLink.guid).AddAllParameters();
            workReference.LoadSettings.AddRelation(PlanReferenceGuids.Work.SubWorksLink).AddAllParameters();
            workReference.LoadSettings.AddAllParameters();
            var filter = new Filter(workReferenceInfo);
            filter.Terms.AddTerm("ID", ComparisonOperator.IsOneOf, workIDs);
            list = workReference.Find(filter);

            return list;
        }
    }
}
