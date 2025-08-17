using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SummaryReport
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.canselButton = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.LogTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Processinglabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // canselButton
            // 
            this.canselButton.Location = new System.Drawing.Point(382, 291);
            this.canselButton.Name = "canselButton";
            this.canselButton.Size = new System.Drawing.Size(75, 23);
            this.canselButton.TabIndex = 0;
            this.canselButton.Text = "Отмена";
            this.canselButton.UseVisualStyleBackColor = true;
            this.canselButton.Click += new System.EventHandler(this.canselButton_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(20, 262);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(437, 23);
            this.progressBar1.TabIndex = 2;
            // 
            // LogTextBox
            // 
            this.LogTextBox.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.LogTextBox.Location = new System.Drawing.Point(21, 46);
            this.LogTextBox.Multiline = true;
            this.LogTextBox.Name = "LogTextBox";
            this.LogTextBox.Size = new System.Drawing.Size(436, 164);
            this.LogTextBox.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Журнал операций:";
            // 
            // Processinglabel
            // 
            this.Processinglabel.AutoSize = true;
            this.Processinglabel.Location = new System.Drawing.Point(23, 235);
            this.Processinglabel.Name = "Processinglabel";
            this.Processinglabel.Size = new System.Drawing.Size(0, 13);
            this.Processinglabel.TabIndex = 5;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(485, 332);
            this.Controls.Add(this.Processinglabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.LogTextBox);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.canselButton);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование сводной ведомости";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button canselButton;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox LogTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label Processinglabel;
    }

    public partial class MainForm : Form
    {

        public enum Stages
        {
            Initialization,
            DataAcquisition,
            DataProcessing,
            ReportGenerating,
            Done
        }

        public MainForm()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            InitializeComponent();
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

        public void progressBarParam(int step)
        {

            progressBar1.PerformStep();
        }
        public void setStage(Stages stage)
        {
            switch (stage)
            {
                case Stages.Initialization:
                    Processinglabel.Text = "Инициализация...";
                    break;
                case Stages.DataAcquisition:
                    Processinglabel.Text = "Получение данных";
                    break;
                case Stages.DataProcessing:
                    Processinglabel.Text = "Обработка данных";
                    break;
                case Stages.ReportGenerating:
                    Processinglabel.Text = "Создание отчёта";
                    break;
                case Stages.Done:
                    break;
                default:
                    System.Diagnostics.Debug.Assert(false);
                    break;
            }
            this.Update();
        }

        private void canselButton_Click(object sender, EventArgs e)
        {
            MainForm.ActiveForm.Close();
        }
    }
}
