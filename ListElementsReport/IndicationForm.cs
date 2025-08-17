using System;
using System.Windows.Forms;

namespace ListElementsReport
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
            CheckData,
            DataProcessing,
            ReportGenerating,
            Done
        }

        public void writeToLog(string line)
        {
            DateTime now = DateTime.Now;
            string s = String.Format("[{0:yyyy-MM-dd_HH:mm:ss}] {1}\r\n", now, line);
            LogTextBox.AppendText(s);
            LogTextBox.Invalidate();
            this.Update();
            System.Windows.Forms.Application.DoEvents();
        }

        public void setStage(Stages stage)
        {
            switch (stage)
            {
                case Stages.Initialization:
                    processingLabel.Text = "Инициализация...";
                    progressBar.PerformStep();
                    break;
                case Stages.DataAcquisition:
                    processingLabel.Text = "Получение данных";
                    progressBar.PerformStep();
                    break;
                case Stages.CheckData:
                    processingLabel.Text = "Проверка данных";
                    progressBar.PerformStep();
                    break;
                case Stages.DataProcessing:
                    processingLabel.Text = "Обработка данных";
                    progressBar.PerformStep();
                    break;          
               case Stages.ReportGenerating:
                    processingLabel.Text = "Создание отчёта";
                    progressBar.PerformStep();
                    break;
                case Stages.Done:
                    progressBar.PerformStep();
                    break;
                default:
                    System.Diagnostics.Debug.Assert(false);
                    break;
            }
            this.Update();
        }
    }
}
