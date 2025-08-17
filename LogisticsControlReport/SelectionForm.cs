using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace LogisticsControlReport
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public ReportParameters reportParameters = new ReportParameters();
        public SelectionForm()
        {
            InitializeComponent();
        }
        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private List<ReferenceObject> employee914 = new List<ReferenceObject>();

        private void buttonMakeReport_Click_1(object sender, EventArgs e)
        {
            //Выюор периода исполнения
            reportParameters.AllTime = radioButtonAllTime.Checked ? true : false;
            reportParameters.SelectedPeriod = radioButtonSelectedPeriod.Checked ? true : false;
            reportParameters.BeginPeriod = dateTimePickerBegin.Value.Date;
            reportParameters.EndPeriod = dateTimePickerEnd.Value.Date;

            //Выюор периода исполнения по незачетам
            reportParameters.SelectedPeriodFails = radioButtonSelectedPeriodsFails.Checked ? true : false;
            reportParameters.BeginPeriodFails = dateTimePickerBeginFails.Value.Date;
            reportParameters.EndPeriodFails = dateTimePickerEndFails.Value.Date;

            //Выбор исполнения
            reportParameters.AllExecuter = radioButtonAllMaker.Checked ? true : false;
            reportParameters.SelectedExecuter = radioButtonSelectedMaker.Checked ? true : false;
            if (reportParameters.SelectedExecuter)
                reportParameters.Executer = employee914.Where(t => t[Guids.UsersParameters.ShortName].Value.ToString() == comboBoxMaker.Text).FirstOrDefault();

            //Выбор даты исполнения
            reportParameters.DateExecute = comboBoxDateExecute.Text;
          
            DialogResult = DialogResult.OK;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            dateTimePickerEnd.Value = DateTime.Now.Date;
            dateTimePickerBegin.Value = DateTime.Now.AddDays(-30).Date;
            dateTimePickerEndFails.Value = DateTime.Now.Date;
            dateTimePickerBeginFails.Value = DateTime.Now.AddDays(-30).Date;
            comboBoxDateExecute.Text = "все";

            //Получаем справочник Справочник ОМТС
            var referenceOMTSInfo = ReferenceCatalog.FindReference(Guids.References.OMTS);
            var referenceOMTS = referenceOMTSInfo.CreateReference();

            referenceOMTS.Objects.Load();
            var makers = referenceOMTS.Objects.Select(t => t[Guids.OMTSParameters.Executer].Value.ToString()).Distinct();

            //Получаем справочник Группы и пользователи
            var referenceUsersInfo = ReferenceCatalog.FindReference(Guids.References.GroupsAndUsers);
            var referenceUsers = referenceUsersInfo.CreateReference();

            //Набираем объекты для фильтра по сотрудникам отдела 914
            List<ReferenceObject> dep914 = new List<ReferenceObject>();
            // Отдел 914
            dep914.Add(referenceUsers.Find(Guids.UserObject.Department914));
            // Начальник отдела 914
            dep914.Add(referenceUsers.Find(Guids.UserObject.Chief914));
            // Склад 1
            dep914.Add(referenceUsers.Find(Guids.UserObject.Store1));
            // Склад 4
            dep914.Add(referenceUsers.Find(Guids.UserObject.Store4));

            //Формируем фильтр по сотрудникам отдел 914
            var filter = new Filter(referenceUsersInfo);
            // Родительский объект входим в список
            filter.Terms.AddTerm("[Родительский объект]", ComparisonOperator.IsOneOf, dep914);
            // Тип- сотрудник 
            filter.Terms.AddTerm(referenceUsers.ParameterGroup[SystemParameterType.Class], ComparisonOperator.Equal, Guids.UserType.Employee);

            //Сформированный фильтр
            employee914 = referenceUsers.Find(filter).Where(t => makers.Contains(t.ToString())).ToList();
            //Заполняем comboBox исполнители
            foreach (var employee in employee914.Select(t => t[Guids.UsersParameters.ShortName].Value.ToString()).OrderBy(t => t))
            {
                comboBoxMaker.Items.Add(employee);
            }

            comboBoxMaker.Text = employee914.First()[Guids.UsersParameters.ShortName].Value.ToString();
        }

        private void radioButtonSelectedMaker_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSelectedMaker.Checked)
            {
                comboBoxMaker.Enabled = true;
            }
            else
            {
                comboBoxMaker.Enabled = false;
            }
        }

        private void radioButtonAllMaker_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonAllMaker.Checked)
            {
                comboBoxMaker.Enabled = false;
            }
            else
            {
                comboBoxMaker.Enabled = true;
            }
        }

        private void radioButtonSelectedPeriod_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSelectedPeriod.Checked)
            {
                dateTimePickerBegin.Enabled = true;
                dateTimePickerEnd.Enabled = true;
                dateTimePickerBeginFails.Enabled = false;
                dateTimePickerEndFails.Enabled = false;
            }
            else
            {
                dateTimePickerBegin.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBeginFails.Enabled = false;
                dateTimePickerEndFails.Enabled = false;
            }
        }

        private void radioButtonAllTime_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonAllTime.Checked)
            {
                dateTimePickerBegin.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBeginFails.Enabled = false;
                dateTimePickerEndFails.Enabled = false;
            }
            else
            {
                dateTimePickerBegin.Enabled = true;
                dateTimePickerEnd.Enabled = true;
                dateTimePickerBeginFails.Enabled = true;
                dateTimePickerEndFails.Enabled = true;
            }
        }

        private void radioButtonSelectedPeriodsFails_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSelectedPeriodsFails.Checked)
            {
                dateTimePickerBeginFails.Enabled = true;
                dateTimePickerEndFails.Enabled = true;
                dateTimePickerBegin.Enabled = false;
                dateTimePickerEnd.Enabled = false;
            }
            else
            {
                dateTimePickerBeginFails.Enabled = false;
                dateTimePickerEndFails.Enabled = false;
                dateTimePickerBegin.Enabled = false;
                dateTimePickerEnd.Enabled = false;
            }
        }
    }

    public class ReportParameters
    {
        // Выбор периода исполнения
        public bool AllTime;
        public bool SelectedPeriod;
        public DateTime BeginPeriod;
        public DateTime EndPeriod;

        // Выбор периода по незачетам
        public bool SelectedPeriodFails;
        public DateTime BeginPeriodFails;
        public DateTime EndPeriodFails;

        //Выбор исполнителя
        public bool AllExecuter;
        public bool SelectedExecuter;
        public ReferenceObject Executer;

        //Выбор даты исполнения
        public string DateExecute;
    }
}
