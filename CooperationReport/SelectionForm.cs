using System;
using System.Windows.Forms;
using TFlex.Reporting;



namespace CooperationReport
{
    public partial class SelectionForm : Form
    {


        private IReportGenerationContext _context;
        public ReportParameters reportParameters;

        public SelectionForm()
        {
            InitializeComponent();
        }

        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            comboBoxMonth.SelectedIndex = DateTime.Now.Month-1;
            comboBoxYear.Text = DateTime.Now.Year.ToString();
            for(int i = 0; i <= 10; i++)
            {
                comboBoxYear.Items.Add((DateTime.Now.Year - i).ToString());
            }
        }

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }


        private bool MakeReportParameters()
        {
            reportParameters = new ReportParameters();
         
            reportParameters.MonthReport = radioButtonMonth.Checked;
            reportParameters.DateReport = radioButtonDate.Checked;
            reportParameters.PeriodReport = radioButtonForPeriod.Checked;

            if (reportParameters.PeriodReport)
            {
                reportParameters.BeginDate = dateTimePickerBegin.Value;
                reportParameters.EndDate = dateTimePickerEnd.Value;
            }

            if (reportParameters.DateReport)
            {
                reportParameters.Date = dateTimePicker.Value;
            }
            if (reportParameters.MonthReport)
            {
                reportParameters.Month = comboBoxMonth.SelectedIndex+1;
                reportParameters.Year = Convert.ToInt32(comboBoxYear.Text.Trim());
                reportParameters.MonthString = comboBoxMonth.Text;
            }

            reportParameters.ExpenseData = (radButExpense.Checked) ? true : false;
            reportParameters.SumOnDevices = (radButDevice.Checked) ? true : false;

            return true;
        }

        private void radioButtonDate_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker.Visible = true;
            comboBoxMonth.Visible = false;
            comboBoxYear.Visible = false;

            if (!radioButtonForPeriod.Checked)
            {
                dateTimePickerBegin.Visible = false;
                dateTimePickerEnd.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
            }
        }

        private void radioButtonMonth_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker.Visible = false;
            comboBoxMonth.Visible = true;
            comboBoxYear.Visible = true;

            if (!radioButtonForPeriod.Checked)
            {
                dateTimePickerBegin.Visible = false;
                dateTimePickerEnd.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void radButArrival_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radButContract_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radButDevice_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForPeriod.Checked)
            {
                dateTimePickerBegin.Visible = true;
                dateTimePickerEnd.Visible = true;
                label1.Visible = true;
                label2.Visible = true;

                

                dateTimePicker.Visible = false;
                comboBoxMonth.Visible = false;
                comboBoxYear.Visible = false;
            }
            else
            {
                dateTimePickerBegin.Visible = false;
                dateTimePickerEnd.Visible = false;
                label1.Visible = false;
                label2.Visible = false;
            }
        }
    }
}
