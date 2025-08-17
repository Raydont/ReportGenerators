using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpecificationsReport
{
    public partial class Attributes : Form
    {
        public Attributes()
        {
            InitializeComponent();
        }

        private void chBoxAddModify_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAddModify.CheckState == CheckState.Checked)
                grBoxModifySpec.Enabled = true;
            else grBoxModifySpec.Enabled = false;
        }

        private void Attributes_Load(object sender, EventArgs e)
        {

        }

        private void grBoxModifySpec_Enter(object sender, EventArgs e)
        {

        }

        private void Attributes_FormClosed(object sender, FormClosedEventArgs e)
        {
            closeForm = true;
        }

      
    }
}
