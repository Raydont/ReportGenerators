using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ActualComplexesReport
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
    }
}
