using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using DocumentAttributes;
using TFlex.Reporting;
using SpecificationReport.Espd;

namespace DraughtSpecificationESPD
{
    public partial class Attributes : Form
    {
        public Attributes()
        {
            InitializeComponent();
        }
        private System.Windows.Forms.Button makeReportButton;
        private bool _attributesFormClose = false;
        private bool _closeGenerator = false;

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            _attributesFormClose = true;
        }

        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }
        
        public bool AttributesFormClose
        {
            get { return _attributesFormClose; }
        }

        public bool CloseGenerator
        {
            get { return _closeGenerator; }
        }


        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();

            TFDDocument document = new TFDDocument(context.ObjectsInfo[0].ObjectId);
            string nameRow = document.Naimenovanie; 
            int indx = nameRow.IndexOf("Спецификация");
            if (indx >= 0)
                nameRow = nameRow.Remove(indx, 12);
            else
                MessageBox.Show("В наименовании объекта отсутствует, либо неправильно введено слово \"Спецификация\"","Внимание",MessageBoxButtons.OK,MessageBoxIcon.Warning);

            tBoxNameProgram.Text = nameRow.ToUpper();

            textBoxFullName.Text = nameRow.ToLower();

            while (AttributesFormClose == false)
            {
                this.UpdateUI();
            }

            if (CloseGenerator)
            {
                _closeGenerator = false;
                this.Close();
                return;
            }

            DocumentAttributes.DocumentAttributes documentAttr = new DocumentAttributes.DocumentAttributes();
            documentAttr.NameProgram = nameRow;
            documentAttr.SidebarSignHeader1 = (tbSideBarSignHeader1.Text == null) ? "" : tbSideBarSignHeader1.Text;
            documentAttr.SidebarSignHeader2 = (tbSideBarSignHeader2.Text == null) ? "" : tbSideBarSignHeader2.Text;
            documentAttr.SidebarSignHeader3 = (tbSideBarSignHeader3.Text == null) ? "" : tbSideBarSignHeader3.Text;

            documentAttr.SidebarSignName1 = (tbSideBarSignName1.Text == null) ? "" : tbSideBarSignName1.Text;
            documentAttr.SidebarSignName2 = (tbSideBarSignName2.Text == null) ? "" : tbSideBarSignName2.Text;
            documentAttr.SidebarSignName3 = (tbSideBarSignName3.Text == null) ? "" : tbSideBarSignName3.Text;

            documentAttr.DesignedBy = (tbDesignedBy.Text == null) ? "" : tbDesignedBy.Text;
            documentAttr.ControlledBy = (tbControlledBy.Text == null) ? "" : tbControlledBy.Text;
            documentAttr.ApprovedBy = (tbApprovedBy.Text == null) ? "" : tbApprovedBy.Text;
            documentAttr.NControlBy = (tbNControlBy.Text == null) ? "" : tbNControlBy.Text;

            documentAttr.FirstPageNumber = Convert.ToInt32(nudFirstPageNumber.Value);

            if (cbAddChangelogPage.CheckState == CheckState.Checked)
                documentAttr.AddChangelogPage = true;
            else documentAttr.AddChangelogPage = false;

            if (cbFullNames.CheckState == CheckState.Checked)
            {
                documentAttr.FullNames = true;
                documentAttr.NameProgFullName = textBoxFullName.Text;
            }
            else documentAttr.FullNames = false;

            documentAttr.CountEmptyLinesBefore = Convert.ToInt32(nudCountBefore.Value);

            documentAttr.CountEmptyLinesAfter = Convert.ToInt32(nudCountAfter.Value);

            if ((chBoxTitul.CheckState == CheckState.Checked)||(chBoxLU.CheckState == CheckState.Checked))
            {
                documentAttr.Zakaz = tBoxNameZakaz.Text;
                documentAttr.NameProgram = tBoxNameProgram.Text;
                documentAttr.YearMake = tBoxYearMake.Text;
                if (chBoxLU.CheckState == CheckState.Checked)
                {
                    documentAttr.Lu = true;

                    documentAttr.Goev = tBoxGoev.Text;
                    if (chBoxMilitaryChief.CheckState == CheckState.Checked)
                    {
                        documentAttr.MilitaryChief = tBoxMilitaryChief.Text;
                        documentAttr.CheckMilitaryChief = true;
                    }
                    else documentAttr.CheckMilitaryChief = false;

                    documentAttr.Designer = tBoxDesigner.Text;
                    documentAttr.Norm = tBoxNorm.Text;

                    if (chBoxChiefTeam.CheckState == CheckState.Checked)
                    {
                        documentAttr.ChiefTeam = tBoxChiefTeam.Text;
                        documentAttr.CheckChiefTeam = true;
                    }
                    else documentAttr.CheckChiefTeam = false;

                    if (chBoxVPMO.CheckState == CheckState.Checked)
                    {
                        documentAttr.VPMO = tBoxVPMO.Text;
                        documentAttr.CheckVPMO = true;
                    }
                    else documentAttr.CheckVPMO = false;
                    if (chBoxZgkz.CheckState == CheckState.Checked)
                    {
                        documentAttr.SurnameZgkz = tBoxZGKZ.Text;
                        documentAttr.Zgkz = true;
                    }
                    else documentAttr.Zgkz = false;
                }
                if (chBoxTitul.CheckState == CheckState.Checked)
                    documentAttr.TitulList = true;
            }

            documentAttr.Litera = document.Litera;

            this.Close();

            // ТОЧКА ВХОДА МАКРОСА
            SpecificationReport.Espd.EspdSpecificationCadReport.Make(context, documentAttr);
        }

        private void chBoxTitul_CheckedChanged(object sender, EventArgs e)
        {
            if ((chBoxTitul.CheckState == CheckState.Checked)||((chBoxLU.CheckState == CheckState.Checked)))
                grBoxParametr.Enabled = true;
            else
                grBoxParametr.Enabled = false;
        }

        private void chBoxLU_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxLU.CheckState == CheckState.Checked)
            {
                grBoxLU.Enabled = true;
                grBoxParametr.Enabled = true;
            }
            else
            {
                grBoxLU.Enabled = false;
                if (chBoxTitul.CheckState == CheckState.Checked)
                    grBoxParametr.Enabled = true;
                else grBoxParametr.Enabled = false;
            }
        }

        private void chBoxMilitaryChief_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxMilitaryChief.CheckState == CheckState.Checked)
                tBoxMilitaryChief.Enabled = true;
            else tBoxMilitaryChief.Enabled = false;
        }

        private void chBoxChiefTeam_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxChiefTeam.CheckState == CheckState.Checked)
                tBoxChiefTeam.Enabled = true;
            else tBoxChiefTeam.Enabled = false;
        }

        private void chBoxVPMO_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxVPMO.CheckState == CheckState.Checked)
                tBoxVPMO.Enabled = true;
            else tBoxVPMO.Enabled = false;
        }

        private void Attributes_FormClosed(object sender, FormClosedEventArgs e)
        {
            _attributesFormClose = true;
            _closeGenerator = true;
        }

        private void cbFullNames_CheckedChanged(object sender, EventArgs e)
        {

            if (cbFullNames.CheckState == CheckState.Checked)
                textBoxFullName.Enabled = true;
            else textBoxFullName.Enabled = false;
        }

        private void chBoxZgkz_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxZgkz.CheckState == CheckState.Checked)
                tBoxZGKZ.Enabled = true;
            else tBoxZGKZ.Enabled = false;
        }
    }
}
