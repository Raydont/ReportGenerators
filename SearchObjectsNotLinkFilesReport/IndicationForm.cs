using System;
using System.Windows.Forms;

namespace SearchObjectsNotLinkFilesReport
{
    public partial class IndicationForm : Form
    {
        public IndicationForm()
        {
            Application.EnableVisualStyles();
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

        public void writeToLog(string line)
        {
            DateTime now = DateTime.Now;
            string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
            logTextBox.AppendText(s);
            Application.DoEvents();
        }

        public void progressBarParam(int step)
        {

            progressBarSearchObject.PerformStep();
        }
        public void setStage(Stages stage)
        {
            switch (stage)
            {
                case Stages.Initialization:
                    processingLabel.Text = "Инициализация...";
                    progressBarParam(2);
                    break;
                case Stages.DataAcquisition:
                    processingLabel.Text = "Получение данных";
                    progressBarParam(2);
                    break;
                case Stages.DataProcessing:
                    processingLabel.Text = "Обработка данных";
                    progressBarParam(2);
                    break;
                case Stages.ReportGenerating:
                    processingLabel.Text = "Создание отчёта";
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
    }
}
