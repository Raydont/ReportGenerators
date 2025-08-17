using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportIncomingDoc
{
    public partial class FormSelection : Form
    {
        public FormSelection()
        {
            InitializeComponent();
            dtpStart.MaxDate = DateTime.Today.AddDays(-1);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (dtpStart.Value > dtpEnd.Value)
            {
                MessageBox.Show("Дата конца периода меньше даты начала!", "Введены некорректные данные.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                DialogResult = DialogResult.OK;
            }
        }
    }
}