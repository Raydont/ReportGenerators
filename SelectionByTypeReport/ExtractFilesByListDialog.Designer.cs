namespace SelectionByTypeReport
{
    partial class ExtractFilesByListDialog
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
            this.DenotationListTextBox = new System.Windows.Forms.TextBox();
            this.ExtractButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(310, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Список обозначений (одно обозначение на каждой строке):";
            // 
            // DenotationListTextBox
            // 
            this.DenotationListTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DenotationListTextBox.Location = new System.Drawing.Point(12, 25);
            this.DenotationListTextBox.Multiline = true;
            this.DenotationListTextBox.Name = "DenotationListTextBox";
            this.DenotationListTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.DenotationListTextBox.Size = new System.Drawing.Size(401, 256);
            this.DenotationListTextBox.TabIndex = 1;
            // 
            // ExtractButton
            // 
            this.ExtractButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.ExtractButton.Location = new System.Drawing.Point(118, 287);
            this.ExtractButton.Name = "ExtractButton";
            this.ExtractButton.Size = new System.Drawing.Size(89, 23);
            this.ExtractButton.TabIndex = 2;
            this.ExtractButton.Text = "Выгрузить";
            this.ExtractButton.UseVisualStyleBackColor = true;
            this.ExtractButton.Click += new System.EventHandler(this.ExtractButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(213, 287);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(89, 23);
            this.CancelButton.TabIndex = 3;
            this.CancelButton.Text = "Отмена";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ExtractFilesByListDialog
            // 
            this.AcceptButton = this.ExtractButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelButton;
            this.ClientSize = new System.Drawing.Size(425, 322);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.ExtractButton);
            this.Controls.Add(this.DenotationListTextBox);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(441, 360);
            this.Name = "ExtractFilesByListDialog";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Выгрузка файлов по списку обозначений";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button ExtractButton;
        private System.Windows.Forms.Button CancelButton;
        public System.Windows.Forms.TextBox DenotationListTextBox;
    }
}