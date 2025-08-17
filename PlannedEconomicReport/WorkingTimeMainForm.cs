using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;

namespace PlannedEconomicReport
{
    public partial class WorkingTimeMainForm : Form
    {
        public WorkingTimeMainForm()
        {
            InitializeComponent();

            var cadresReference = ReferenceCatalog.FindReference(Guids.Cadres.RefereceGuid);
            var cadresReferenceInfo = cadresReference.CreateReference();

            foreach (var department in cadresReferenceInfo.Objects)
            {
                departmentComboBox.Items.Add(department);
            }

            departmentComboBox.SelectedIndex = 0;

            date.Value = DateTime.Now;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
