using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InventoryBook
{
    public partial class ReportForm : Form
    {
        public ReportParameters reportParameters = new ReportParameters();

        private bool iNumber;
        private DateTime startDate;
        private DateTime endDate;
        private int startNumber;
        private int endNumber;

        public ReportForm()
        {
            InitializeComponent();
        }

        private void ReportButton_Click(object sender, EventArgs e)
        {
            // Класс, форматирующий текст
            TextFormatter tForm = new TextFormatter(new System.Drawing.Font("T-FLEX Type T", 2.5f, GraphicsUnit.Millimeter));

            if (iNumber)
            {
                startNumber = Convert.ToInt32(startNTextBox.Text);
                endNumber = Convert.ToInt32(endNTextBox.Text);
                if (startNumber > endNumber)
                {
                    MessageBox.Show("Неверно задан диапазон инвентарных номеров", "Ошибка!");
                    this.DialogResult = DialogResult.Cancel;
                }
                else
                    this.DialogResult = DialogResult.OK;
            }
            else
            {
                startDate = new DateTime(startDateTimePicker.Value.Year, startDateTimePicker.Value.Month, startDateTimePicker.Value.Day, 0, 0, 0);
                endDate = new DateTime(endDateTimePicker.Value.Year, endDateTimePicker.Value.Month, endDateTimePicker.Value.Day, 23, 59, 59);
                if (startDate > endDate)
                {
                    MessageBox.Show("Неверно задан период отчета", "Ошибка!");
                    this.DialogResult = DialogResult.Cancel;
                }
                else
                    this.DialogResult = DialogResult.OK;
            }

            reportParameters.iNumber = iNumber;
            reportParameters.startDate = startDate;
            reportParameters.endDate = endDate;
            reportParameters.startNumber = startNumber;
            reportParameters.endNumber = endNumber;
            reportParameters.tForm = tForm;
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            startDateTimePicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1, 0, 0, 0);
            endDateTimePicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 23, 59, 59);
            startNTextBox.Text = "479100";
            endNTextBox.Text = "479500";
            iNumber = false;
            groupBox2.Enabled = false;
        }

        private void iNumberCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (iNumberCheckBox.Checked)
            {
                groupBox1.Enabled = false;
                groupBox2.Enabled = true;
                iNumber = true;
            }
            else
            {
                groupBox1.Enabled = true;
                groupBox2.Enabled = false;
                iNumber = false;
            }
        }

        private void startNTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyDigits(e);
        }

        private void endNTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            OnlyDigits(e);
        }

        private void OnlyDigits(KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!Char.IsDigit(ch) && ch != 8)
            {
                e.Handled = true;
            }
        }
    }
}
