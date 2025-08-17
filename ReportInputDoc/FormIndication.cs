using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportIncomingDoc
{
    public partial class FormIndication : Form
    {
        public FormIndication()
        {
            InitializeComponent();
        }

        public void WriteToLog(string line)
        {
            DateTime now = DateTime.Now;
            string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
            tbLog.AppendText(s);
            Application.DoEvents();
        }
    }
}
