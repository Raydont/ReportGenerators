using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using PlanReport.Types;
using ReportHelpers;

namespace Globus.PlanReportsExcel
{
    public partial class ReportParametersForm : Form
    {
        public PlanReportExcel engine;
        public List<SignBlock> signBlockList;
        public List<SignBlock> footerSignList;
        public WorkHeader workHeader;
        public WorkTypes workType;
        public List<WorkParameters> workList;
        public Xls xls;
        public ProgressForm p_form;
        public int step;
        public int columnsCount;

        bool isTable;
        bool isBook;
        bool newPage;

        public ReportParametersForm()
        {
            InitializeComponent();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            if (isTable)
            {
                p_form.progressFormCaption("Идет формирование отчета...");
                p_form.Visible = true;
                engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workHeader, workType, workList, isBook, newPage, xls, p_form, step);
                p_form.Close();
            }
            else
            {
                p_form.progressFormCaption("Идет формирование отчета...");
                p_form.Visible = true;
                engine.CreatePlanReport(signBlockList, columnsCount, footerSignList, workHeader, workType, workList, isBook, xls, p_form, step);
                p_form.Close();
            }

            this.Close();
        }

        private void tableFormRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (tableFormRadioButton.Checked)
                isTable = true;
            else
                isTable = false;

            if (!isTable)
            {
                newPageGroupBox.Enabled = false;
                firstPageRadioButton.Checked = true;
            }
            else
            {
                newPageGroupBox.Enabled = true;
            }
        }

        private void albumRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (albumRadioButton.Checked)
                isBook = false;
            else
                isBook = true;
        }

        private void firstPageRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (firstPageRadioButton.Checked)
                newPage = false;
            else
                newPage = true;
        }

        private void ReportParametersForm_Load(object sender, EventArgs e)
        {
            if (tableFormRadioButton.Checked)
                isTable = true;
            else
                isTable = false;

            if (albumRadioButton.Checked)
                isBook = false;
            else
                isBook = true;

            if (firstPageRadioButton.Checked)
                newPage = false;
            else
                newPage = true;
        }
    }
}
