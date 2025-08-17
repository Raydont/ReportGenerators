namespace SearchObjectsNotLinkFilesReport
{
    partial class IndicationForm
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
            this.progressBarSearchObject = new System.Windows.Forms.ProgressBar();
            this.logTextBox = new System.Windows.Forms.TextBox();
            this.processingLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBarSearchObject
            // 
            this.progressBarSearchObject.Location = new System.Drawing.Point(12, 179);
            this.progressBarSearchObject.Name = "progressBarSearchObject";
            this.progressBarSearchObject.Size = new System.Drawing.Size(564, 27);
            this.progressBarSearchObject.TabIndex = 0;
            // 
            // logTextBox
            // 
            this.logTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.logTextBox.Location = new System.Drawing.Point(12, 12);
            this.logTextBox.Multiline = true;
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.Size = new System.Drawing.Size(564, 137);
            this.logTextBox.TabIndex = 1;
            // 
            // processingLabel
            // 
            this.processingLabel.AutoSize = true;
            this.processingLabel.Location = new System.Drawing.Point(12, 158);
            this.processingLabel.Name = "processingLabel";
            this.processingLabel.Size = new System.Drawing.Size(0, 13);
            this.processingLabel.TabIndex = 2;
            // 
            // IndicationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(588, 218);
            this.Controls.Add(this.processingLabel);
            this.Controls.Add(this.logTextBox);
            this.Controls.Add(this.progressBarSearchObject);
            this.Name = "IndicationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Поиск объектов в заказах не связанных с файлами";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBarSearchObject;
        private System.Windows.Forms.TextBox logTextBox;
        private System.Windows.Forms.Label processingLabel;
    }
}