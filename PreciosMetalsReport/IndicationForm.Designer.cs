namespace PreciosMetalsReport
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
            this.LogTextBox = new System.Windows.Forms.TextBox();
            this.processinglabel = new System.Windows.Forms.Label();
            this.buttonCansel = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // LogTextBox
            // 
            this.LogTextBox.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.LogTextBox.Location = new System.Drawing.Point(12, 24);
            this.LogTextBox.Multiline = true;
            this.LogTextBox.Name = "LogTextBox";
            this.LogTextBox.Size = new System.Drawing.Size(412, 110);
            this.LogTextBox.TabIndex = 16;
            // 
            // processinglabel
            // 
            this.processinglabel.AutoSize = true;
            this.processinglabel.Location = new System.Drawing.Point(9, 149);
            this.processinglabel.Name = "processinglabel";
            this.processinglabel.Size = new System.Drawing.Size(0, 13);
            this.processinglabel.TabIndex = 15;
            // 
            // buttonCansel
            // 
            this.buttonCansel.Location = new System.Drawing.Point(322, 202);
            this.buttonCansel.Name = "buttonCansel";
            this.buttonCansel.Size = new System.Drawing.Size(102, 23);
            this.buttonCansel.TabIndex = 14;
            this.buttonCansel.Text = "Отмена";
            this.buttonCansel.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 173);
            this.progressBar.Maximum = 100000;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(412, 23);
            this.progressBar.Step = 1;
            this.progressBar.TabIndex = 13;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Журнал операций";
            // 
            // IndicationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 229);
            this.Controls.Add(this.LogTextBox);
            this.Controls.Add(this.processinglabel);
            this.Controls.Add(this.buttonCansel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.label1);
            this.Name = "IndicationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Формирование отчета";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox LogTextBox;
        private System.Windows.Forms.Label processinglabel;
        private System.Windows.Forms.Button buttonCansel;
        public System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label label1;
    }
}