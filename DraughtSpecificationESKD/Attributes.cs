using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Globus.TFlexDocs; 

namespace DraughtSpecificationESKD
{
    public partial class Attributes : Form
    {
        public Attributes()
        {
            InitializeComponent();
        }

        private void chBoxAddTitul_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAddTitul.CheckState == CheckState.Checked)
            {
                tBoxGoev.Enabled = true;
                tBoxMilitaryChief.Enabled = true;
                cBoxTech.Enabled = true;
                cBoxNorm.Enabled = true;
                tBoxVPMO.Enabled = true;
                tBox45.Enabled = true;
                tBoxCheifTeam.Enabled = true;
                tBoxYearMake.Enabled = true;
                chBoxVPMO.Enabled = true;
                chBoxChiefMillitary.Enabled = true;
            }
            else
            {
                tBoxGoev.Enabled = false;
                tBoxMilitaryChief.Enabled = false;
                cBoxTech.Enabled = false;
                cBoxNorm.Enabled = false;
                tBoxVPMO.Enabled = false;
                tBox45.Enabled = false;
                tBoxCheifTeam.Enabled = false;
                tBoxYearMake.Enabled = false;
                chBoxVPMO.Enabled = false;
                chBoxChiefMillitary.Enabled = false;
            }
        }

        private void Attributes_FormClosed(object sender, FormClosedEventArgs e)
        {
            _attributesFormClose = true;
            _closeGenerator = true;
        }

        private void chBoxNewPageOtherItemsAfterProductsKind_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chBoxChiefMillitary_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxChiefMillitary.CheckState == CheckState.Unchecked)
            {
                tBoxMilitaryChief.Enabled = false;
                tBoxMilitaryChief.Text = String.Empty;
            }
            else tBoxMilitaryChief.Enabled = true;
        }

        private void chBoxVPMO_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxVPMO.CheckState == CheckState.Unchecked)
            {
                tBoxVPMO.Enabled = false;
                tBoxVPMO.Text = String.Empty;
            }
            else tBoxVPMO.Enabled = true;
        }

        private void tbSideBarSignHeader1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void chBoxAddModify_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAddModify.CheckState == CheckState.Checked)
                grBoxModifySpec.Enabled = true;
            else grBoxModifySpec.Enabled = false;
        }

        private void grBoxModifySpec_Enter(object sender, EventArgs e)
        {

        }

        private void tBoxGoev_TextChanged(object sender, EventArgs e)
        {

        }

        private void chBoxForExecution_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chBoxDocVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDocVarPart.Checked)
                nudDocumentVarPart.Enabled = true;
            else nudDocumentVarPart.Enabled = false;
        }

        private void chBoxAssemblyVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAssemblyVarPart.Checked)
                nudAssemblyVarPart.Enabled = true;
            else nudAssemblyVarPart.Enabled = false;
        }

        private void chBoxDetailVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDetailVarPart.Checked)
                nudDetailVarPart.Enabled = true;
            else nudDetailVarPart.Enabled = false;
        }

        private void chBoxStdItemVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxStdItemVarPart.Checked)
                nudStdItemsVarPart.Enabled = true;
            else nudStdItemsVarPart.Enabled = false;
        }

        private void chBoxOtherItemVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxOtherItemVarPart.Checked)
                nudOtherItemsVarPart.Enabled = true;
            else nudOtherItemsVarPart.Enabled = false;
        }

        private void chBoxMaterialVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxMaterialVarPart.Checked)
                nudMaterialVarPart.Enabled = true;
            else nudMaterialVarPart.Enabled = false;
        }

        private void chBoxKomplectVarPart_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxKomplectVarPart.Checked)
                nudKomplektVarPart.Enabled = true;
            else nudKomplektVarPart.Enabled = false;
        }

        private void chBoxAddMake_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAddMake.Checked)
                chBoxThroughNumbering.Enabled = true;
            else chBoxThroughNumbering.Enabled = false;
        }

        private void chBoxThroughNumbering_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbFullNames_CheckedChanged(object sender, EventArgs e)
        {
            if (cbFullNames.CheckState == CheckState.Checked)
                textBoxFullName.Enabled = true;
            else textBoxFullName.Enabled = false;
        }

        private void checkBoxSectionProgramSoft_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
