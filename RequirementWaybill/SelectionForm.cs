using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.Reporting;

namespace RequirementWaybill
{
    public partial class SelectionForm : Form
    {
        public ReportParameters reportParameters;
        public SelectionForm()
        {
            InitializeComponent();
            TopMost = true;
        }

        internal IReportGenerationContext reportContext;

        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            reportParameters = new ReportParameters();
            reportParameters.ItemsCount = Convert.ToInt32(NUDItemsCount.Value);
            reportParameters.OtherItems = otherItemsRadioButton.Checked;
            reportParameters.StandardItems = standatItemsRadioButton.Checked;
            reportParameters.SeparateAssemblies = separateAssemblies.Checked;

            reportParameters.ActionType = actionTypeTextBox.Text;
            reportParameters.Receiver = receiverTextBox.Text;
            reportParameters.Schet = schetTextBox.Text;
            reportParameters.Department = departmentTextBox.Text;

            DialogResult = DialogResult.OK;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        internal void Init(NomenclatureObject rootObject)
        {
            var shifr = rootObject[new Guid("e49f42c1-3e91-42b7-8764-28bde4dca7b6")].GetString();
            if (!string.IsNullOrWhiteSpace(shifr))
            {
                actionTypeTextBox.Text = shifr.Trim();
            }
            else
            {
                actionTypeTextBox.Text = rootObject[new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")].GetString();
            }
        }
    }
}
