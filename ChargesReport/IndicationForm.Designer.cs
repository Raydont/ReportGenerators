namespace ChargesReport
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
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.buttonCansel = new System.Windows.Forms.Button();
            this.processinglabel = new System.Windows.Forms.Label();
            this.LogTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(21, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Журнал операций";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(24, 179);
            this.progressBar.Maximum = 102;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(412, 23);
            this.progressBar.Step = 1;
            this.progressBar.TabIndex = 2;
            // 
            // buttonCansel
            // 
            this.buttonCansel.Location = new System.Drawing.Point(334, 208);
            this.buttonCansel.Name = "buttonCansel";
            this.buttonCansel.Size = new System.Drawing.Size(102, 23);
            this.buttonCansel.TabIndex = 3;
            this.buttonCansel.Text = "Отмена";
            this.buttonCansel.UseVisualStyleBackColor = true;
            this.buttonCansel.Click += new System.EventHandler(this.buttonCansel_Click);
            // 
            // processinglabel
            // 
            this.processinglabel.AutoSize = true;
            this.processinglabel.Location = new System.Drawing.Point(21, 155);
            this.processinglabel.Name = "processinglabel";
            this.processinglabel.Size = new System.Drawing.Size(0, 13);
            this.processinglabel.TabIndex = 4;
            // 
            // LogTextBox
            // 
            this.LogTextBox.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.LogTextBox.Location = new System.Drawing.Point(24, 30);
            this.LogTextBox.Multiline = true;
            this.LogTextBox.Name = "LogTextBox";
            this.LogTextBox.Size = new System.Drawing.Size(412, 110);
            this.LogTextBox.TabIndex = 6;
            // 
            // IndicationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(458, 240);
            this.Controls.Add(this.LogTextBox);
            this.Controls.Add(this.processinglabel);
            this.Controls.Add(this.buttonCansel);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.label1);
            this.Name = "IndicationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ведомость начисления премии рабочим";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button buttonCansel;
        private System.Windows.Forms.Label processinglabel;
        private System.Windows.Forms.TextBox LogTextBox;
    }
}