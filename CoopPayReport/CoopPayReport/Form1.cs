using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CoopPayReport
{
    public partial class mainPeriodForm : Form
    {
        public DateTime startPeriod;
        public DateTime endPeriod;

        public mainPeriodForm()
        {
            InitializeComponent();

            dtpYear.CustomFormat = "yyyy";
        }

        private void reportButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;

            if (calendarCheckBox.Checked)
            {
                startPeriod = dtpStartPeriod.Value;
                endPeriod = dtpEndPeriod.Value;
            }
            else
            {
                startPeriod = new DateTime(dtpYear.Value.Year, 1, 1);
                endPeriod = new DateTime(dtpYear.Value.Year, 12, 31);
            }
        }

        private void calendarCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (calendarCheckBox.Checked)
            {
                yearGroupBox.Enabled = false;
                calendarGroupBox.Enabled = true;
            }
            else
            {
                yearGroupBox.Enabled = true;
                calendarGroupBox.Enabled = false;
            }
        }
    }
}
