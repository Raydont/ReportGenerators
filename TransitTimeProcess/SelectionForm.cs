using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace TransitTimeProcess
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public SelectionForm()
        {
            InitializeComponent();
        }
        public ReportParameters reportParameters = new ReportParameters();

        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            var referenceInfo = ReferenceCatalog.FindReference(_context.ObjectsInfo[0].ReferenceId);
            var reference = referenceInfo.CreateReference();

            var filter = new TFlex.DOCs.Model.Search.Filter(referenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, _context.ObjectsInfo[0].ObjectId);

            ReferenceObject selectedObj = reference.Find(filter).FirstOrDefault();
            reportParameters.RefObject = selectedObj;

            radioButtonForSelectedObject.Text = "Для " + selectedObj;
            dateTimePickerBegin.Value = DateTime.Now.AddMonths(-1).Date;
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }

        private bool MakeReportParameters()
        {
            reportParameters.SelectedObject = radioButtonForSelectedObject.Checked ? true : false;
            reportParameters.ObjectsByPeriods = radioButtonForList.Checked ? true : false;
            reportParameters.BeginPeriod = dateTimePickerBegin.Value.Date;
            reportParameters.EndPeriod = dateTimePickerEnd.Value.Date;

            return true;
        }

       

        private void radioButtonForListPKI_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePickerBegin.Enabled = true;
            dateTimePickerEnd.Enabled = true;
        }

        private void radioButtonForSelectedObject_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePickerBegin.Enabled = false;
            dateTimePickerEnd.Enabled = false;
        }

     
    }
}
