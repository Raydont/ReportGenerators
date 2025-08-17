using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

namespace RunningContractsReport
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

        private void checkBoxActualContract_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxActualContract.CheckState == CheckState.Unchecked)
                groupBox2.Enabled = true;
            else
                groupBox2.Enabled = false;

        }

        private void buttonCansel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            reportParameters = new ReportParameters();

            reportParameters.ActualContract = (checkBoxActualContract.CheckState == CheckState.Checked) ? true : false;
            reportParameters.DOD = (radButDOD.Checked) ? true : false;
            reportParameters.BeginDate = Convert.ToDateTime(dateTimePickerBegin.Value);
            reportParameters.EndDate = Convert.ToDateTime(dateTimePickerEnd.Value);

            DialogResult = DialogResult.OK;
        }

        private void radButDZU_CheckedChanged(object sender, EventArgs e)
        {
            if (radButDZU.Checked)
            {
                checkBoxActualContract.Enabled = false;
                groupBox2.Enabled = true;
                checkBoxActualContract.CheckState = CheckState.Unchecked;
            }
        }

        private void radButDOD_CheckedChanged(object sender, EventArgs e)
        {
            if (radButDOD.Checked)
            {
                checkBoxActualContract.Enabled = true;
                groupBox2.Enabled = false;
                checkBoxActualContract.CheckState = CheckState.Checked;
            }
        }



    }

}
