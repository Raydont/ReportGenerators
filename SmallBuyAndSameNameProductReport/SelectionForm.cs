using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

namespace SmallBuyAndSameNameProductReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            reportTypeComboBox.SelectedIndex = 0;
        }

        public void Init(IReportGenerationContext context)
        {
            GetReportParameters getReportParameters = new GetReportParameters();
            AttributeReport attributeReport = new AttributeReport();

            /*attributeReport.isSmallBuyReport = false;
            if (reportTypeComboBox.SelectedIndex == 0)
                attributeReport.isSmallBuyReport = true;*/

            attributeReport.startDate = startPeriodDateTimePicker.Value;
            attributeReport.endDate = endPeriodDateTimePicker.Value;

            getReportParameters.FillRKK(context, attributeReport);
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
