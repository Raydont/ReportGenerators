using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpecificationESKD.исходный_макрос
{
    public partial class AttributesForm : Form
    {
        public AttributesForm()
        {
            InitializeComponent();
        }

        private void makeReportButton_Click(object sender, EventArgs e)
        {

        }

        private void chBoxDocumentation_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDocumentation.Checked)
                nudDocumentation.Enabled = true;
            else nudDocumentation.Enabled = false;
        }

        private void chBoxAssembly_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAssembly.Checked)
                nudAssembly.Enabled = true;
            else nudAssembly.Enabled = false;
        }

        private void chBoxDetails_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDetails.Checked)
                nudDetails.Enabled = true;
            else nudDetails.Enabled = false;
        }

        private void chBoxStdProducts_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxStdProducts.Checked)
                nudStdProducts.Enabled = true;
            else nudStdProducts.Enabled = false;
        }

        private void chBoxOtherProducts_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxOtherProducts.Checked)
                nudStdProducts.Enabled = true;
            else nudStdProducts.Enabled = false;
        }

        private void chBoxMaterials_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxMaterials.Checked)
                nudMaterials.Enabled = true;
            else nudMaterials.Enabled = false;
        }

        private void chBoxKomplekt_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxKomplekt.Checked)
                nudKomplekt.Enabled = true;
            else nudKomplekt.Enabled = false;
        }
    }
}
