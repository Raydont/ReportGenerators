using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EntrancesObjectsNotInNomReferences
{
    public partial class IndicationForm : Form
    {
        public IndicationForm()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            InitializeComponent();
        }

        public enum Stages
        {
            Initialization,
            DataAcquisition,
            DataProcessing,
            ReportGenerating,
            Done
        }

        int counter;
        public void writeToLog(string line)
        {
            counter++;
            DateTime now = DateTime.Now;
            string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
            LogTextBox.AppendText(s);

            if (counter > 5)
            {
                Application.DoEvents();
                LogTextBox.Invalidate();
                counter = 0;
            }
        }

        public void progressBarParam(int step)
        {

            progressBar.PerformStep();
        }
        public void setStage(Stages stage)
        {
            switch (stage)
            {
                case Stages.Initialization:
                    processinglabel.Text = "Инициализация...";
                    progressBarParam(2);
                    break;
                case Stages.DataAcquisition:
                    processinglabel.Text = "Получение данных";
                    progressBarParam(2);
                    break;
                case Stages.DataProcessing:
                    processinglabel.Text = "Обработка данных";
                    progressBarParam(2);
                    break;
                case Stages.ReportGenerating:
                    processinglabel.Text = "Создание отчёта";
                    progressBarParam(2);
                    break;
                case Stages.Done:
                    progressBarParam(2);
                    break;
                default:
                    System.Diagnostics.Debug.Assert(false);
                    break;
            }
            this.Update();
        }

        private void buttonCansel_Click_1(object sender, EventArgs e)
        {
            IndicationForm.ActiveForm.Close();
        }
    }
}
