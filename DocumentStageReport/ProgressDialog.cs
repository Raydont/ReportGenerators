using System.IO;
using System.Windows.Forms;

namespace Globus.DOCs.Technology.Reports.ServerSide.Dialogs
{
    public partial class ProgressDialog : Form
    {
        public static ProgressDialog instance;

        public string logFile;

        public ProgressDialog()
        {
            InitializeComponent();
            instance = this;

            FormClosing += ProgressDialog_FormClosing;
        }

        private void ProgressDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            instance = null;
        }

        public void SetInfo(string info)
        {
            progressLabel.Text = info;
            if (!string.IsNullOrWhiteSpace(logFile))
            {
                if (!File.Exists(logFile))
                {
                    File.WriteAllText(logFile, "");
                }

                File.AppendAllLines(logFile, new string[] { info });
            }
            Application.DoEvents();
        }

        public void SetProgress(int value, int count)
        {
            progressBar.Minimum = 0;
            progressBar.Maximum = count;
            progressBar.Value = value > count ? count : value;
            Application.DoEvents();
        }

    }
}
