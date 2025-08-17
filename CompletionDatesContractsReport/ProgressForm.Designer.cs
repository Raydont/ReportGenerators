namespace CompletionDatesContractsReport
{
    partial class ProgressForm
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
            this.label = new System.Windows.Forms.Label();
            this.progressBarMakeReport = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(13, 11);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(263, 13);
            this.label.TabIndex = 3;
            this.label.Text = "Пожалуйста подождите. Формированиея отчета...";
            // 
            // progressBarMakeReport
            // 
            this.progressBarMakeReport.Location = new System.Drawing.Point(13, 35);
            this.progressBarMakeReport.Maximum = 1000;
            this.progressBarMakeReport.Name = "progressBarMakeReport";
            this.progressBarMakeReport.Size = new System.Drawing.Size(403, 23);
            this.progressBarMakeReport.Step = 1;
            this.progressBarMakeReport.TabIndex = 2;
            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(429, 69);
            this.Controls.Add(this.label);
            this.Controls.Add(this.progressBarMakeReport);
            this.Name = "ProgressForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование отчета";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label label;
        public System.Windows.Forms.ProgressBar progressBarMakeReport;
    }
}