using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

using TFlex.Reporting;
using TFlex.DOCs.Model.References;
using ReportHelpers;

namespace WorkingClothesReport
{
    public partial class ReportForm : Form
    {
        public List<ReferenceObject> departments = new List<ReferenceObject>();

        private WorkingClothesReport report;

        public ReportForm()
        {
            InitializeComponent();
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            overallRadioButton.Checked = true;
            halfYearComboBox.SelectedIndex = 0;

            int year = DateTime.Now.Year;
            for (int i = 0; i < 7; i++)
                yearComboBox.Items.Add(year + 1 - i);

            yearComboBox.SelectedIndex = 1;
        }

        private void overallRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (overallRadioButton.Checked)
            {
                departComboBox.Items.Clear();
                departComboBox.Items.Add(String.Empty);
                departComboBox.Items.Clear();
            }
        }

        private void departRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (departRadioButton.Checked)
            {
                departComboBox.Text = String.Empty;
                foreach (var dep in departments)
                {
                    departComboBox.Items.Add(dep);
                }
                departComboBox.SelectedIndex = 0;
                departComboBox.Enabled = true;
            }
        }

        // Точка входа (создание отчета)
        public void MakeReport(IReportGenerationContext context)
        {
            this.ShowDialog();

            report = new WorkingClothesReport();

            // Формирование отчета
            Xls xls = new Xls();
            bool needMake = false;

            context.CopyTemplateFile();    // Создаем копию шаблона      

            this.Close();

            if (DialogResult == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(context.ReportFilePath);
                    // Создание отчета (Сводная ведомость или Заявка подразделения)
                    if (overallRadioButton.Checked)
                        needMake = report.Make(true, String.Empty, halfYearComboBox.Text, yearComboBox.Text, xls);
                    else
                        needMake = report.Make(false, departComboBox.Text, halfYearComboBox.Text, yearComboBox.Text, xls);
                }
                finally
                {
                    xls.Quit(true);
                    if (needMake)
                        Process.Start(context.ReportFilePath);
                }
            }
        }

        private void createButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
