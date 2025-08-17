using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanReport.Types
{
    public enum WorkTypes
    {
        SimplePlan,
        PreliminaryTestPlan,
        DocDevelopPlan,
        PeriodicTestPlan,
        ProductionPlan
    }
    // Заголовок плана (графика)
    public class WorkHeader
    {
        string graphName;

        // Номер графика
        public string GraphName
        {
            get { return graphName; }
            set { graphName = value; }
        }
    }
    // Тип Простая работа
    public class WorkParameters
    {
        bool isHeader;
        string workName;
        string workSummary;
        string planItem;
        string executor;
        string deadline;
        string comment;

        // Является ли работа заголовком
        public bool IsHeader
        {
            get { return isHeader; }
            set { isHeader = value; }
        }

        // Название работы (план, мероприятия) - заголовок
        public string WorkName
        {
            get { return workName; }
            set { workName = value; }
        }

        // Краткое содержание
        public string WorkSummary
        {
            get { return workSummary; }
            set { workSummary = value; }
        }

        // Номер пункта плана (мероприятий)
        public string PlanItem
        {
            get { return planItem; }
            set { planItem = value; }
        }

        // Исполнитель (для пункта плана)
        public string Executor
        {
            get { return executor; }
            set { executor = value; }
        }

        // Срок выполнения (для пункта плана)
        public string Deadline
        {
            get { return deadline; }
            set { deadline = value; }
        }

        // Примечание
        public string Comment
        {
            get { return comment; }
            set { comment = value; }
        }
    }
    // Тип График разработки КД
    public class WorkDocDevelopParameters
    {
        bool isHeader;
        string structItem;
        string deviceName;
        string developer;
        string techTaskConfirmDate;
        string techTaskStatementDate;
        string structDevelopDate;
        string serviceDocTo906Date;
        string initialDataDate;
        string engineeringDocDate;
        string engineeringDocComplectDate;
        string textEngineeringDocDate;
        string comments;

        // Является ли работа заголовком
        public bool IsHeader
        {
            get { return isHeader; }
            set { isHeader = value; }
        }
        // Наименование устройства / документа
        public string DeviceName
        {
            get { return deviceName; }
            set { deviceName = value; }
        }
        // Номер пункта структуры
        public string StructItem
        {
            get { return structItem; }
            set { structItem = value; }
        }
        // Отдел-разработчик
        public string Developer
        {
            get { return developer; }
            set { developer = value; }
        }
        // Утвержденное ТЗ выдано
        public string TechTaskConfirmDate
        {
            get { return techTaskConfirmDate; }
            set { techTaskConfirmDate = value; }
        }
        // ТЗ в смежные подразделения
        public string TechTaskStatementDate
        {
            get { return techTaskStatementDate; }
            set { techTaskStatementDate = value; }
        }
        // Э3 разработано, КД выдано
        public string StructDevelopDate
        {
            get { return structDevelopDate; }
            set { structDevelopDate = value; }
        }
        // СЗ в отдел 906 выдано
        public string ServiceDocTo906Date
        {
            get { return serviceDocTo906Date; }
            set { serviceDocTo906Date = value; }
        }
        // ИД на ПП выданы
        public string InitialDataDate
        {
            get { return initialDataDate; }
            set { initialDataDate = value; }
        }
        // КД в отд.20 передана
        public string EngineeringDocDate
        {
            get { return engineeringDocDate; }
            set { engineeringDocDate = value; }
        }
        // Комплект КД в БТД сдан
        public string EngineeringDocComplectDate
        {
            get { return engineeringDocComplectDate; }
            set { engineeringDocComplectDate = value; }
        }
        // Текстовая КД разработана
        public string TextEngineeringDocDate
        {
            get { return textEngineeringDocDate; }
            set { textEngineeringDocDate = value; }
        }
        // Примечания
        public string Comments
        {
            get { return comments; }
            set { comments = value; }
        }
    }
    // Тип График периодических испытаний
    public class WorkPeriodicTestParameters
    {
        bool isHeader;
        string structItem;
        string deviceName;
        string actPT;
        string actPTFinishDate;
        string productionGraph;
        string productionDate;
        string pTDate;

        // Является ли работа заголовком
        public bool IsHeader
        {
            get { return isHeader; }
            set { isHeader = value; }
        }
        // Наименование устройства и его шифр
        public string DeviceName
        {
            get { return deviceName; }
            set { deviceName = value; }
        }
        // Номер пункта структуры
        public string StructItem
        {
            get { return structItem; }
            set { structItem = value; }
        }
        // Акт ПИ
        public string ActPT
        {
            get { return actPT; }
            set { actPT = value; }
        }
        // Дата окончания действия акта ПИ
        public string ActPTFinishDate
        {
            get { return actPTFinishDate; }
            set { actPTFinishDate = value; }
        }
        // График изготовления
        public string ProductionGraph
        {
            get { return productionGraph; }
            set { productionGraph = value; }
        }
        // Срок изготовления
        public string ProductionDate
        {
            get { return productionDate; }
            set { productionDate = value; }
        }
        // Сроки проведения ПИ
        public string PTDate
        {
            get { return pTDate; }
            set { pTDate = value; }
        }
    }
    // Тип График предварительных испытаний
    public class WorkPreliminaryTestParameters
    {
        bool isHeader;
        string planItem;
        string planItemName;
        string executor;
        string methodologyDate;
        string testDate;
        string protocolConfirmDate;
        string comments;

        // Является ли работа заголовком
        public bool IsHeader
        {
            get { return isHeader; }
            set { isHeader = value; }
        }
        // Краткое наименование пункта программы
        public string PlanItemName
        {
            get { return planItemName; }
            set { planItemName = value; }
        }
        // Номер пункта структуры
        public string PlanItem
        {
            get { return planItem; }
            set { planItem = value; }
        }
        // Исполнитель
        public string Executor
        {
            get { return executor; }
            set { executor = value; }
        }
        // Методика разработана
        public string MethodologyDate
        {
            get { return methodologyDate; }
            set { methodologyDate = value; }
        }
        // Испытания проведены
        public string TestDate
        {
            get { return testDate; }
            set { testDate = value; }
        }
        // Протокол согласован и утвержден
        public string ProtocolConfirmDate
        {
            get { return protocolConfirmDate; }
            set { protocolConfirmDate = value; }
        }
        // Примечание
        public string Comments
        {
            get { return comments; }
            set { comments = value; }
        }
    }
    // Тип График изготовления
    public class WorkProductionParameters
    {
        bool isHeader;
        string planItem;
        string name;
        string denotation;
        string amount;
        string executor;
        string startDate;
        string docNumber;
        string bureau219Date;
        string dep60Date;
        string bureau904Date;
        string materialsDate;
        string othersDate;
        string sector800Date;
        string sector850Date;
        string sector750Date;
        string settingDate;
        string settingExecutor;
        string defectiveDate;
        string deliveryDate;
        string deliveryExecutor;
        string comment;

        // Является ли работа заголовком
        public bool IsHeader
        {
            get { return isHeader; }
            set { isHeader = value; }
        }
        // Номер пункта плана
        public string PlanItem
        {
            get { return planItem; }
            set { planItem = value; }
        }
        // Наименование устройства и его шифр
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        // Децимальный номер
        public string Denotation
        {
            get { return denotation; }
            set { denotation = value; }
        }
        // Количество
        public string Amount
        {
            get { return amount; }
            set { amount = value; }
        }
        // Разработчик
        public string Executor
        {
            get { return executor; }
            set { executor = value; }
        }
        // Дата
        public string StartDate
        {
            get { return startDate; }
            set { startDate = value; }
        }
        // Номер документа
        public string DocNumber
        {
            get { return docNumber; }
            set { docNumber = value; }
        }
        // Н.-з. и к.д. в отд.60
        public string Bureau219Date
        {
            get { return bureau219Date; }
            set { bureau219Date = value; }
        }
        // Технологическая обр-ка н.з.
        public string Dep60Date
        {
            get { return dep60Date; }
            set { dep60Date = value; }
        }
        // Нормирование
        public string Bureau904Date
        {
            get { return bureau904Date; }
            set { bureau904Date = value; }
        }
        // Материалы
        public string MaterialsDate
        {
            get { return materialsDate; }
            set { materialsDate = value; }
        }
        // КРЭ
        public string OthersDate
        {
            get { return othersDate; }
            set { othersDate = value; }
        }
        // Участок 800
        public string Sector800Date
        {
            get { return sector800Date; }
            set { sector800Date = value; }
        }
        // Участок 850
        public string Sector850Date
        {
            get { return sector850Date; }
            set { sector850Date = value; }
        }
        // Участок 750
        public string Sector750Date
        {
            get { return sector750Date; }
            set { sector750Date = value; }
        }
        // Настройка / проверка
        public string SettingDate
        {
            get { return settingDate; }
            set { settingDate = value; }
        }
        // Исполнитель настройки
        public string SettingExecutor
        {
            get { return settingExecutor; }
            set { settingExecutor = value; }
        }
        // Дефектировка
        public string DefectiveDate
        {
            get { return defectiveDate; }
            set { defectiveDate = value; }
        }
        // Сдача
        public string DeliveryDate
        {
            get { return deliveryDate; }
            set { deliveryDate = value; }
        }
        // Исполнитель сдачи
        public string DeliveryExecutor
        {
            get { return deliveryExecutor; }
            set { deliveryExecutor = value; }
        }
        // Примечание
        public string Comment
        {
            get { return comment; }
            set { comment = value; }
        }
    }

}
