using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocumentStagesReport
{
    public partial class ReportForm : Form
    {
        public DateTime beginDate;
        public DateTime endDate;

        public ReportForm()
        {
            InitializeComponent();
        }

        private void generateButton_Click(object sender, EventArgs e)
        {
            // Проверка, что дата начала должна быть меньше даты окончания
            beginDate = dateTimePickerBegin.Value;
            endDate = dateTimePickerEnd.Value;
            if (beginDate > endDate)
            {
                MessageBox.Show("Неверно задан период отчета");
            }
            else
            {
                DialogResult = DialogResult.OK;
            }
        }
    }
}
