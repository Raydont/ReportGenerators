using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

namespace EntranceReport
{
    public partial class AddFilterForm : Form
    {
        public AddFilterForm()
        {
            InitializeComponent();
        }

        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();

            while (SelectionFormClose == false)
            {
                this.UpdateUI();
            }

            var reportParameters = new ReportParameters();

            reportParameters.AddLikeName.AddRange(nameList);
            reportParameters.AddLikeDenotation.AddRange(denotationList);

            reportParameters.Documentation = (chkDocumentation.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Complex = (chkComplex.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Detail = (chkDetail.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Assembly = (chkAssembly.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Standart = (chkStandart.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Other = (chkOther.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Materials = (chkMaterial.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Complement = (chkComplement.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComplexProgram = (chkComplexProgram.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ComponentProgram = (chkComponentProgram.CheckState == CheckState.Checked) ? true : false;

            reportParameters.Positions = (chkPosition.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Zone = (chkZone.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Format = (chkFormat.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Denotation = (chkDenotation.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Name = (chkName.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Count = (chkCount.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Remarks = (chkRemarks.CheckState == CheckState.Checked) ? true : false;
            reportParameters.UnitMeasure = (chkUnitMeasure.CheckState == CheckState.Checked) ? true : false;
            reportParameters.CountOnKnot = (chkCountOnKnot.CheckState == CheckState.Checked) ? true : false;
            reportParameters.CountOnProduct = (chkCountOnProduct.CheckState == CheckState.Checked) ? true : false;

            reportParameters.EntranceName = (checkBoxName.CheckState == CheckState.Checked) ? true : false;
            reportParameters.EntranceDenotation = (checkBoxDenotation.CheckState == CheckState.Checked) ? true : false;



            // ТОЧКА ВХОДА МАКРОСА
            if (this._haltForm == false)
            {
                this.Close();
                var report = new Report();
                report.Make(context, reportParameters);
            }
            else
            {
                this.Close();
            }
        }

        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }

        private void addFilterButton_Click(object sender, EventArgs e)
        {
            var itmx = new ListViewItem[1];

            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked && AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add(textBoxAddLikeName.Text);
                itmx[0].SubItems.Add(textBoxAddLikeDenotation.Text);
                nameList.Add(textBoxAddLikeName.Text.Trim());
                denotationList.Add(textBoxAddLikeDenotation.Text.Trim());
                filterListView.Items.AddRange(itmx);

                return;
            }


            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add(textBoxAddLikeName.Text);
                itmx[0].SubItems.Add("");

                denotationList.Add(null);
                nameList.Add(textBoxAddLikeName.Text.Trim());

                filterListView.Items.AddRange(itmx);
                return;
            }

            if (AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                itmx[0] = new ListViewItem((filterListView.Items.Count + 1).ToString());

                itmx[0].SubItems.Add("");
                itmx[0].SubItems.Add(textBoxAddLikeDenotation.Text);

                nameList.Add(null);
                denotationList.Add(textBoxAddLikeDenotation.Text.Trim());

                filterListView.Items.AddRange(itmx);
                return;
            }
        }

        private List<string> nameList = new List<string>();
        private List<string> denotationList = new List<string>();




        private bool _selectionFormClose = false;
        private bool _haltForm = false;

        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            if ((chkDocumentation.CheckState == CheckState.Unchecked) && (chkComplex.CheckState == CheckState.Unchecked) && (chkDetail.CheckState == CheckState.Unchecked) &&
               (chkAssembly.CheckState == CheckState.Unchecked) && (chkStandart.CheckState == CheckState.Unchecked) && (chkOther.CheckState == CheckState.Unchecked) &&
               (chkMaterial.CheckState == CheckState.Unchecked) && (chkComplement.CheckState == CheckState.Unchecked) && (chkComplexProgram.CheckState == CheckState.Unchecked) &&
               (chkComponentProgram.CheckState == CheckState.Unchecked))
            {
                MessageBox.Show("Вы не выбрали ни одного типа объекта для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }

            if ((chkName.CheckState == CheckState.Unchecked) && (chkDenotation.CheckState == CheckState.Unchecked) && (chkFormat.CheckState == CheckState.Unchecked) &&
               (chkZone.CheckState == CheckState.Unchecked) && (chkPosition.CheckState == CheckState.Unchecked) && (chkCount.CheckState == CheckState.Unchecked) &&
               (chkRemarks.CheckState == CheckState.Unchecked) && (chkCountOnKnot.CheckState == CheckState.Unchecked) && (chkCountOnProduct.CheckState == CheckState.Unchecked) &&
               (chkUnitMeasure.CheckState == CheckState.Unchecked) && (checkBoxDenotation.CheckState == CheckState.Unchecked) && (checkBoxName.CheckState == CheckState.Unchecked))
            {
                MessageBox.Show("Вы не выбрали ни одного поля для отчета", "Внимание!!!", MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                return;
            }
            _selectionFormClose = true;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this._haltForm = true;
            this._selectionFormClose = true;
        }

        public bool SelectionFormClose
        {
            get { return _selectionFormClose; }
        }

        private void AddConitionsNameCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (AddConitionsNameCheckBox.CheckState == CheckState.Checked)
            {
                nameGroupBox.Enabled = true;
                addFilterButton.Enabled = true;
            }
            else
            {
                nameGroupBox.Enabled = false;
                textBoxAddLikeName.Text = string.Empty;

                if (AddConitionsDenotationCheckBox.CheckState == CheckState.Unchecked)
                    addFilterButton.Enabled = false;
            }
        }

        private void AddConitionsDenotationCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (AddConitionsDenotationCheckBox.CheckState == CheckState.Checked)
            {
                denotationGroupBox.Enabled = true;
                addFilterButton.Enabled = true;
            }
            else
            {
                denotationGroupBox.Enabled = false;
                textBoxAddLikeDenotation.Text = string.Empty;

                if (AddConitionsNameCheckBox.CheckState == CheckState.Unchecked)
                    addFilterButton.Enabled = false;
            }
        }

        private void AddFilterForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this._haltForm = true;
            _selectionFormClose = true;
        }

        private void chkAllTypes_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllTypes.CheckState == CheckState.Checked)
            {
                chkDocumentation.Checked = true;
                chkComplex.Checked = true;
                chkDetail.Checked = true;
                chkAssembly.Checked = true;
                chkStandart.Checked = true;
                chkOther.Checked = true;
                chkMaterial.Checked = true;
                chkComplement.Checked = true;
                chkComponentProgram.Checked = true;
                chkComplexProgram.Checked = true;
            }
            else
            {
                chkDocumentation.Checked = false;
                chkComplex.Checked = false;
                chkDetail.Checked = false;
                chkAssembly.Checked = false;
                chkStandart.Checked = false;
                chkOther.Checked = false;
                chkMaterial.Checked = false;
                chkComplement.Checked = false;
                chkComponentProgram.Checked = false;
                chkComplexProgram.Checked = false;
            }
        }

        private void chkAllObjects_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllObjects.CheckState == CheckState.Checked)
            {
                chkPosition.Checked = true;
                chkZone.Checked = true;
                chkFormat.Checked = true;
                chkDenotation.Checked = true;
                chkName.Checked = true;
                chkCount.Checked = true;
                chkRemarks.Checked = true;
                chkUnitMeasure.Checked = true;
                chkCountOnKnot.Checked = true;
                chkCountOnProduct.Checked = true;
                checkBoxName.Checked = true;
                checkBoxDenotation.Checked = true;
            }
            else
            {
                chkPosition.Checked = false;
                chkZone.Checked = false;
                chkFormat.Checked = false;
                chkDenotation.Checked = false;
                chkName.Checked = false;
                chkCount.Checked = false;
                chkRemarks.Checked = false;
                chkUnitMeasure.Checked = true;
                chkCountOnKnot.Checked = false;
                chkCountOnProduct.Checked = false;
                checkBoxName.Checked = false;
                checkBoxDenotation.Checked = false;
            }
        }
    }
    public class ReportParameters
    {
        // параметр добавление условия содержит наименование
        public List<string> AddLikeName = new List<string>();
        // параметр добавление условия содержит обозначение
        public List<string> AddLikeDenotation = new List<string>();

        public bool Documentation;
        public bool Complex;
        public bool Detail;
        public bool Assembly;
        public bool Standart;
        public bool Other;
        public bool Materials;
        public bool Complement;
        public bool ComplexProgram;
        public bool ComponentProgram;

        public bool EntranceName;
        public bool EntranceDenotation;
        public bool Positions;
        public bool Zone;
        public bool Format;
        public bool Denotation;
        public bool Name;
        public bool Count;
        public bool Remarks;
        public bool UnitMeasure;
        public bool CountOnKnot;
        public bool CountOnProduct;

    }
}

