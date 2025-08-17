namespace AccountingREPReport
{
	partial class SelectionForm
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
			this.dtpStartPeriod = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtpEndPeriod = new System.Windows.Forms.DateTimePicker();
			this.label3 = new System.Windows.Forms.Label();
			this.checkBoxQuarter = new System.Windows.Forms.CheckBox();
			this.comboBoxQuarter = new System.Windows.Forms.ComboBox();
			this.buttonОК = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// dtpStartPeriod
			// 
			this.dtpStartPeriod.Location = new System.Drawing.Point(41, 33);
			this.dtpStartPeriod.Name = "dtpStartPeriod";
			this.dtpStartPeriod.Size = new System.Drawing.Size(140, 20);
			this.dtpStartPeriod.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.label1.Location = new System.Drawing.Point(16, 36);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(14, 13);
			this.label1.TabIndex = 1;
			this.label1.Text = "с";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.label2.Location = new System.Drawing.Point(14, 65);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(21, 13);
			this.label2.TabIndex = 2;
			this.label2.Text = "по";
			// 
			// dtpEndPeriod
			// 
			this.dtpEndPeriod.Location = new System.Drawing.Point(41, 62);
			this.dtpEndPeriod.Name = "dtpEndPeriod";
			this.dtpEndPeriod.Size = new System.Drawing.Size(140, 20);
			this.dtpEndPeriod.TabIndex = 3;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.label3.Location = new System.Drawing.Point(81, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(51, 13);
			this.label3.TabIndex = 4;
			this.label3.Text = "Период";
			// 
			// checkBoxQuarter
			// 
			this.checkBoxQuarter.AutoSize = true;
			this.checkBoxQuarter.Location = new System.Drawing.Point(17, 94);
			this.checkBoxQuarter.Name = "checkBoxQuarter";
			this.checkBoxQuarter.Size = new System.Drawing.Size(68, 17);
			this.checkBoxQuarter.TabIndex = 6;
			this.checkBoxQuarter.Text = "Квартал";
			this.checkBoxQuarter.UseVisualStyleBackColor = true;
			this.checkBoxQuarter.CheckedChanged += new System.EventHandler(this.checkBoxQuarter_CheckedChanged);
			// 
			// comboBoxQuarter
			// 
			this.comboBoxQuarter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxQuarter.FormattingEnabled = true;
			this.comboBoxQuarter.Items.AddRange(new object[] {
            "Квартал 1",
            "Квартал 2",
            "Квартал 3",
            "Квартал 4"});
			this.comboBoxQuarter.Location = new System.Drawing.Point(17, 117);
			this.comboBoxQuarter.Name = "comboBoxQuarter";
			this.comboBoxQuarter.Size = new System.Drawing.Size(164, 21);
			this.comboBoxQuarter.TabIndex = 7;
			this.comboBoxQuarter.SelectedValueChanged += new System.EventHandler(this.comboBoxQuarter_SelectedValueChanged);
			// 
			// buttonОК
			// 
			this.buttonОК.Location = new System.Drawing.Point(57, 154);
			this.buttonОК.Name = "buttonОК";
			this.buttonОК.Size = new System.Drawing.Size(75, 23);
			this.buttonОК.TabIndex = 8;
			this.buttonОК.Text = "ОК";
			this.buttonОК.UseVisualStyleBackColor = true;
			this.buttonОК.Click += new System.EventHandler(this.buttonОК_Click);
			// 
			// SelectionForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(200, 189);
			this.Controls.Add(this.buttonОК);
			this.Controls.Add(this.comboBoxQuarter);
			this.Controls.Add(this.checkBoxQuarter);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtpEndPeriod);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtpStartPeriod);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "SelectionForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Период закупки РЭП";
			this.Load += new System.EventHandler(this.SelectionForm_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DateTimePicker dtpStartPeriod;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.DateTimePicker dtpEndPeriod;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.CheckBox checkBoxQuarter;
		private System.Windows.Forms.ComboBox comboBoxQuarter;
		private System.Windows.Forms.Button buttonОК;
	}
}