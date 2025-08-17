using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Globus.PlanReportsExcel
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            InitializeComponent();
        }

        public int getProgressBarMax()
        {
            return progressBar1.Maximum;
        }

        public void progressBarParam(int step)
        {
            progressBar1.Step = step;
            progressBar1.PerformStep();
        }

        public void progressFormCaption(string caption)
        {
            this.Text = caption;
        }
    }
}
