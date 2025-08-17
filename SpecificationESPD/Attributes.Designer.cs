using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentAttributes;

namespace SpecificationESPD
{
    partial class Attributes
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
            this.makeReportButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tPdocAttributes = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBoxFullName = new System.Windows.Forms.TextBox();
            this.nudCountAfter = new System.Windows.Forms.NumericUpDown();
            this.nudCountBefore = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.cbFullNames = new System.Windows.Forms.CheckBox();
            this.cbAddChangelogPage = new System.Windows.Forms.CheckBox();
            this.nudFirstPageNumber = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbApprovedBy = new System.Windows.Forms.TextBox();
            this.tbNControlBy = new System.Windows.Forms.TextBox();
            this.tbControlledBy = new System.Windows.Forms.TextBox();
            this.tbDesignedBy = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.labelZagolovok2 = new System.Windows.Forms.Label();
            this.labelZagolovok1 = new System.Windows.Forms.Label();
            this.tbSideBarSignName3 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignName2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignName1 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader3 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader1 = new System.Windows.Forms.TextBox();
            this.tPTitulLU = new System.Windows.Forms.TabPage();
            this.chBoxLU = new System.Windows.Forms.CheckBox();
            this.chBoxTitul = new System.Windows.Forms.CheckBox();
            this.grBoxLU = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.chBoxZgkz = new System.Windows.Forms.CheckBox();
            this.label23 = new System.Windows.Forms.Label();
            this.chBoxVPMO = new System.Windows.Forms.CheckBox();
            this.chBoxChiefTeam = new System.Windows.Forms.CheckBox();
            this.label22 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chBoxMilitaryChief = new System.Windows.Forms.CheckBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.grBoxParametr = new System.Windows.Forms.GroupBox();
            this.tBoxYearMake = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.tBoxNameProgram = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.tBoxNameZakaz = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.comboBoxChiefDesigner = new System.Windows.Forms.ComboBox();
            this.comboBoxChiefVPMO = new System.Windows.Forms.ComboBox();
            this.comboBoxMaker = new System.Windows.Forms.ComboBox();
            this.comboBoxRateChecker = new System.Windows.Forms.ComboBox();
            this.comboBoxChiefTeam = new System.Windows.Forms.ComboBox();
            this.comboBoxRepresentative = new System.Windows.Forms.ComboBox();
            this.comboBoxZGKZ = new System.Windows.Forms.ComboBox();
            this.tabControl1.SuspendLayout();
            this.tPdocAttributes.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tPTitulLU.SuspendLayout();
            this.grBoxLU.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grBoxParametr.SuspendLayout();
            this.SuspendLayout();
            // 
            // makeReportButton
            // 
            this.makeReportButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.makeReportButton.Location = new System.Drawing.Point(476, 466);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(107, 23);
            this.makeReportButton.TabIndex = 3;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tPdocAttributes);
            this.tabControl1.Controls.Add(this.tPTitulLU);
            this.tabControl1.Location = new System.Drawing.Point(12, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(571, 451);
            this.tabControl1.TabIndex = 4;
            // 
            // tPdocAttributes
            // 
            this.tPdocAttributes.Controls.Add(this.groupBox3);
            this.tPdocAttributes.Controls.Add(this.groupBox2);
            this.tPdocAttributes.Controls.Add(this.groupBox1);
            this.tPdocAttributes.Location = new System.Drawing.Point(4, 22);
            this.tPdocAttributes.Name = "tPdocAttributes";
            this.tPdocAttributes.Padding = new System.Windows.Forms.Padding(3);
            this.tPdocAttributes.Size = new System.Drawing.Size(563, 425);
            this.tPdocAttributes.TabIndex = 0;
            this.tPdocAttributes.Text = "Атрибуты документа";
            this.tPdocAttributes.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.textBoxFullName);
            this.groupBox3.Controls.Add(this.nudCountAfter);
            this.groupBox3.Controls.Add(this.nudCountBefore);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.cbFullNames);
            this.groupBox3.Controls.Add(this.cbAddChangelogPage);
            this.groupBox3.Controls.Add(this.nudFirstPageNumber);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Location = new System.Drawing.Point(3, 247);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(548, 133);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Нумерация листов и содержание";
            // 
            // textBoxFullName
            // 
            this.textBoxFullName.Enabled = false;
            this.textBoxFullName.Location = new System.Drawing.Point(13, 101);
            this.textBoxFullName.Name = "textBoxFullName";
            this.textBoxFullName.Size = new System.Drawing.Size(529, 20);
            this.textBoxFullName.TabIndex = 8;
            // 
            // nudCountAfter
            // 
            this.nudCountAfter.Location = new System.Drawing.Point(467, 54);
            this.nudCountAfter.Name = "nudCountAfter";
            this.nudCountAfter.Size = new System.Drawing.Size(66, 20);
            this.nudCountAfter.TabIndex = 7;
            this.nudCountAfter.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // nudCountBefore
            // 
            this.nudCountBefore.Location = new System.Drawing.Point(468, 24);
            this.nudCountBefore.Name = "nudCountBefore";
            this.nudCountBefore.Size = new System.Drawing.Size(66, 20);
            this.nudCountBefore.TabIndex = 6;
            this.nudCountBefore.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(253, 56);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(208, 13);
            this.label11.TabIndex = 5;
            this.label11.Text = "Количество пустых строк после записи";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(253, 26);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(190, 13);
            this.label10.TabIndex = 4;
            this.label10.Text = "Количество пустых строк до записи";
            // 
            // cbFullNames
            // 
            this.cbFullNames.AutoSize = true;
            this.cbFullNames.Location = new System.Drawing.Point(13, 78);
            this.cbFullNames.Name = "cbFullNames";
            this.cbFullNames.Size = new System.Drawing.Size(143, 17);
            this.cbFullNames.TabIndex = 3;
            this.cbFullNames.Text = "Полные наименования";
            this.cbFullNames.UseVisualStyleBackColor = true;
            this.cbFullNames.CheckedChanged += new System.EventHandler(this.cbFullNames_CheckedChanged);
            // 
            // cbAddChangelogPage
            // 
            this.cbAddChangelogPage.AutoSize = true;
            this.cbAddChangelogPage.Checked = true;
            this.cbAddChangelogPage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbAddChangelogPage.Location = new System.Drawing.Point(13, 55);
            this.cbAddChangelogPage.Name = "cbAddChangelogPage";
            this.cbAddChangelogPage.Size = new System.Drawing.Size(234, 17);
            this.cbAddChangelogPage.TabIndex = 2;
            this.cbAddChangelogPage.Text = "Добавлять лист регистрации изменений";
            this.cbAddChangelogPage.UseVisualStyleBackColor = true;
            // 
            // nudFirstPageNumber
            // 
            this.nudFirstPageNumber.Location = new System.Drawing.Point(148, 26);
            this.nudFirstPageNumber.Name = "nudFirstPageNumber";
            this.nudFirstPageNumber.Size = new System.Drawing.Size(45, 20);
            this.nudFirstPageNumber.TabIndex = 1;
            this.nudFirstPageNumber.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(10, 26);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(117, 13);
            this.label9.TabIndex = 0;
            this.label9.Text = "Номер первого листа";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.tbApprovedBy);
            this.groupBox2.Controls.Add(this.tbNControlBy);
            this.groupBox2.Controls.Add(this.tbControlledBy);
            this.groupBox2.Controls.Add(this.tbDesignedBy);
            this.groupBox2.Location = new System.Drawing.Point(3, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(548, 129);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Подписи основной надписи";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(154, 100);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "Утв.";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(154, 74);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 13);
            this.label7.TabIndex = 11;
            this.label7.Text = "Н. контр.";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(154, 48);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(36, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Пров.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(154, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Разраб.";
            // 
            // tbApprovedBy
            // 
            this.tbApprovedBy.Location = new System.Drawing.Point(223, 97);
            this.tbApprovedBy.Name = "tbApprovedBy";
            this.tbApprovedBy.Size = new System.Drawing.Size(156, 20);
            this.tbApprovedBy.TabIndex = 8;
            // 
            // tbNControlBy
            // 
            this.tbNControlBy.Location = new System.Drawing.Point(223, 71);
            this.tbNControlBy.Name = "tbNControlBy";
            this.tbNControlBy.Size = new System.Drawing.Size(156, 20);
            this.tbNControlBy.TabIndex = 7;
            // 
            // tbControlledBy
            // 
            this.tbControlledBy.Location = new System.Drawing.Point(223, 45);
            this.tbControlledBy.Name = "tbControlledBy";
            this.tbControlledBy.Size = new System.Drawing.Size(156, 20);
            this.tbControlledBy.TabIndex = 6;
            // 
            // tbDesignedBy
            // 
            this.tbDesignedBy.Location = new System.Drawing.Point(223, 19);
            this.tbDesignedBy.Name = "tbDesignedBy";
            this.tbDesignedBy.Size = new System.Drawing.Size(156, 20);
            this.tbDesignedBy.TabIndex = 5;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.labelZagolovok2);
            this.groupBox1.Controls.Add(this.labelZagolovok1);
            this.groupBox1.Controls.Add(this.tbSideBarSignName3);
            this.groupBox1.Controls.Add(this.tbSideBarSignName2);
            this.groupBox1.Controls.Add(this.tbSideBarSignName1);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader3);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader2);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader1);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(548, 100);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Визы на поле подшивки";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(277, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Фамилия";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(277, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Фамилия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(277, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Фамилия";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 74);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Заголовок";
            // 
            // labelZagolovok2
            // 
            this.labelZagolovok2.AutoSize = true;
            this.labelZagolovok2.Location = new System.Drawing.Point(46, 48);
            this.labelZagolovok2.Name = "labelZagolovok2";
            this.labelZagolovok2.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok2.TabIndex = 8;
            this.labelZagolovok2.Text = "Заголовок";
            // 
            // labelZagolovok1
            // 
            this.labelZagolovok1.AutoSize = true;
            this.labelZagolovok1.Location = new System.Drawing.Point(46, 22);
            this.labelZagolovok1.Name = "labelZagolovok1";
            this.labelZagolovok1.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok1.TabIndex = 7;
            this.labelZagolovok1.Text = "Заголовок";
            // 
            // tbSideBarSignName3
            // 
            this.tbSideBarSignName3.Location = new System.Drawing.Point(349, 71);
            this.tbSideBarSignName3.Name = "tbSideBarSignName3";
            this.tbSideBarSignName3.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName3.TabIndex = 6;
            // 
            // tbSideBarSignName2
            // 
            this.tbSideBarSignName2.Location = new System.Drawing.Point(349, 45);
            this.tbSideBarSignName2.Name = "tbSideBarSignName2";
            this.tbSideBarSignName2.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName2.TabIndex = 5;
            // 
            // tbSideBarSignName1
            // 
            this.tbSideBarSignName1.Location = new System.Drawing.Point(349, 19);
            this.tbSideBarSignName1.Name = "tbSideBarSignName1";
            this.tbSideBarSignName1.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName1.TabIndex = 4;
            // 
            // tbSideBarSignHeader3
            // 
            this.tbSideBarSignHeader3.Location = new System.Drawing.Point(117, 71);
            this.tbSideBarSignHeader3.Name = "tbSideBarSignHeader3";
            this.tbSideBarSignHeader3.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader3.TabIndex = 3;
            // 
            // tbSideBarSignHeader2
            // 
            this.tbSideBarSignHeader2.Location = new System.Drawing.Point(117, 45);
            this.tbSideBarSignHeader2.Name = "tbSideBarSignHeader2";
            this.tbSideBarSignHeader2.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader2.TabIndex = 2;
            // 
            // tbSideBarSignHeader1
            // 
            this.tbSideBarSignHeader1.Location = new System.Drawing.Point(117, 19);
            this.tbSideBarSignHeader1.Name = "tbSideBarSignHeader1";
            this.tbSideBarSignHeader1.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader1.TabIndex = 1;
            // 
            // tPTitulLU
            // 
            this.tPTitulLU.Controls.Add(this.chBoxLU);
            this.tPTitulLU.Controls.Add(this.chBoxTitul);
            this.tPTitulLU.Controls.Add(this.grBoxLU);
            this.tPTitulLU.Controls.Add(this.grBoxParametr);
            this.tPTitulLU.Location = new System.Drawing.Point(4, 22);
            this.tPTitulLU.Name = "tPTitulLU";
            this.tPTitulLU.Padding = new System.Windows.Forms.Padding(3);
            this.tPTitulLU.Size = new System.Drawing.Size(563, 425);
            this.tPTitulLU.TabIndex = 1;
            this.tPTitulLU.Text = "Титульный лист и Лист утверждения";
            this.tPTitulLU.UseVisualStyleBackColor = true;
            // 
            // chBoxLU
            // 
            this.chBoxLU.AutoSize = true;
            this.chBoxLU.Location = new System.Drawing.Point(322, 14);
            this.chBoxLU.Name = "chBoxLU";
            this.chBoxLU.Size = new System.Drawing.Size(171, 17);
            this.chBoxLU.TabIndex = 2;
            this.chBoxLU.Text = "Добавить лист утверждения";
            this.chBoxLU.UseVisualStyleBackColor = true;
            this.chBoxLU.CheckedChanged += new System.EventHandler(this.chBoxLU_CheckedChanged);
            // 
            // chBoxTitul
            // 
            this.chBoxTitul.AutoSize = true;
            this.chBoxTitul.Location = new System.Drawing.Point(70, 14);
            this.chBoxTitul.Name = "chBoxTitul";
            this.chBoxTitul.Size = new System.Drawing.Size(158, 17);
            this.chBoxTitul.TabIndex = 0;
            this.chBoxTitul.Text = "Добавить титульный лист";
            this.chBoxTitul.UseVisualStyleBackColor = true;
            this.chBoxTitul.CheckedChanged += new System.EventHandler(this.chBoxTitul_CheckedChanged);
            // 
            // grBoxLU
            // 
            this.grBoxLU.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grBoxLU.Controls.Add(this.panel2);
            this.grBoxLU.Controls.Add(this.panel1);
            this.grBoxLU.Enabled = false;
            this.grBoxLU.Location = new System.Drawing.Point(9, 165);
            this.grBoxLU.Name = "grBoxLU";
            this.grBoxLU.Size = new System.Drawing.Size(551, 256);
            this.grBoxLU.TabIndex = 1;
            this.grBoxLU.TabStop = false;
            this.grBoxLU.Text = "Подписи и визы листа утверждения";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.comboBoxZGKZ);
            this.panel2.Controls.Add(this.comboBoxRepresentative);
            this.panel2.Controls.Add(this.comboBoxChiefTeam);
            this.panel2.Controls.Add(this.chBoxZgkz);
            this.panel2.Controls.Add(this.label23);
            this.panel2.Controls.Add(this.chBoxVPMO);
            this.panel2.Controls.Add(this.chBoxChiefTeam);
            this.panel2.Controls.Add(this.label22);
            this.panel2.Controls.Add(this.label20);
            this.panel2.Controls.Add(this.label19);
            this.panel2.Location = new System.Drawing.Point(9, 141);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(529, 105);
            this.panel2.TabIndex = 14;
            // 
            // chBoxZgkz
            // 
            this.chBoxZgkz.AutoSize = true;
            this.chBoxZgkz.Checked = true;
            this.chBoxZgkz.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxZgkz.Location = new System.Drawing.Point(365, 81);
            this.chBoxZgkz.Name = "chBoxZgkz";
            this.chBoxZgkz.Size = new System.Drawing.Size(15, 14);
            this.chBoxZgkz.TabIndex = 23;
            this.chBoxZgkz.UseVisualStyleBackColor = true;
            this.chBoxZgkz.CheckedChanged += new System.EventHandler(this.chBoxZgkz_CheckedChanged);
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(35, 81);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(34, 13);
            this.label23.TabIndex = 21;
            this.label23.Text = "ЗГКЗ";
            // 
            // chBoxVPMO
            // 
            this.chBoxVPMO.AutoSize = true;
            this.chBoxVPMO.Checked = true;
            this.chBoxVPMO.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxVPMO.Location = new System.Drawing.Point(365, 55);
            this.chBoxVPMO.Name = "chBoxVPMO";
            this.chBoxVPMO.Size = new System.Drawing.Size(15, 14);
            this.chBoxVPMO.TabIndex = 20;
            this.chBoxVPMO.UseVisualStyleBackColor = true;
            this.chBoxVPMO.CheckedChanged += new System.EventHandler(this.chBoxVPMO_CheckedChanged);
            // 
            // chBoxChiefTeam
            // 
            this.chBoxChiefTeam.AutoSize = true;
            this.chBoxChiefTeam.Checked = true;
            this.chBoxChiefTeam.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxChiefTeam.Location = new System.Drawing.Point(365, 32);
            this.chBoxChiefTeam.Name = "chBoxChiefTeam";
            this.chBoxChiefTeam.Size = new System.Drawing.Size(15, 14);
            this.chBoxChiefTeam.TabIndex = 19;
            this.chBoxChiefTeam.UseVisualStyleBackColor = true;
            this.chBoxChiefTeam.CheckedChanged += new System.EventHandler(this.chBoxChiefTeam_CheckedChanged);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label22.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label22.Location = new System.Drawing.Point(111, 9);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(269, 13);
            this.label22.TabIndex = 18;
            this.label22.Text = "Визы на поле подшивки листа утверждения";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(35, 55);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(123, 13);
            this.label20.TabIndex = 15;
            this.label20.Text = "Представитель ВП МО";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(35, 32);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(108, 13);
            this.label19.TabIndex = 13;
            this.label19.Text = "Начальник бригады";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.comboBoxRateChecker);
            this.panel1.Controls.Add(this.comboBoxMaker);
            this.panel1.Controls.Add(this.comboBoxChiefVPMO);
            this.panel1.Controls.Add(this.comboBoxChiefDesigner);
            this.panel1.Controls.Add(this.chBoxMilitaryChief);
            this.panel1.Controls.Add(this.label21);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Location = new System.Drawing.Point(9, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(529, 120);
            this.panel1.TabIndex = 13;
            // 
            // chBoxMilitaryChief
            // 
            this.chBoxMilitaryChief.AutoSize = true;
            this.chBoxMilitaryChief.Checked = true;
            this.chBoxMilitaryChief.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxMilitaryChief.Location = new System.Drawing.Point(365, 48);
            this.chBoxMilitaryChief.Name = "chBoxMilitaryChief";
            this.chBoxMilitaryChief.Size = new System.Drawing.Size(15, 14);
            this.chBoxMilitaryChief.TabIndex = 20;
            this.chBoxMilitaryChief.UseVisualStyleBackColor = true;
            this.chBoxMilitaryChief.CheckedChanged += new System.EventHandler(this.chBoxMilitaryChief_CheckedChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label21.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.label21.Location = new System.Drawing.Point(158, 4);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(195, 13);
            this.label21.TabIndex = 17;
            this.label21.Text = "Подписи на листе утверждения";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(40, 25);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(117, 13);
            this.label13.TabIndex = 9;
            this.label13.Text = "Главный конструктор";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(40, 95);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(53, 13);
            this.label16.TabIndex = 15;
            this.label16.Text = "Н. контр.";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(40, 71);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(74, 13);
            this.label15.TabIndex = 13;
            this.label15.Text = "Исполнитель";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(40, 48);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(100, 13);
            this.label14.TabIndex = 11;
            this.label14.Text = "Начальник ВП МО";
            // 
            // grBoxParametr
            // 
            this.grBoxParametr.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grBoxParametr.Controls.Add(this.tBoxYearMake);
            this.grBoxParametr.Controls.Add(this.label18);
            this.grBoxParametr.Controls.Add(this.tBoxNameProgram);
            this.grBoxParametr.Controls.Add(this.label17);
            this.grBoxParametr.Controls.Add(this.tBoxNameZakaz);
            this.grBoxParametr.Controls.Add(this.label12);
            this.grBoxParametr.Enabled = false;
            this.grBoxParametr.Location = new System.Drawing.Point(9, 37);
            this.grBoxParametr.Name = "grBoxParametr";
            this.grBoxParametr.Size = new System.Drawing.Size(548, 122);
            this.grBoxParametr.TabIndex = 0;
            this.grBoxParametr.TabStop = false;
            this.grBoxParametr.Text = "Общие параметры листов";
            // 
            // tBoxYearMake
            // 
            this.tBoxYearMake.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tBoxYearMake.Location = new System.Drawing.Point(174, 94);
            this.tBoxYearMake.Name = "tBoxYearMake";
            this.tBoxYearMake.Size = new System.Drawing.Size(120, 20);
            this.tBoxYearMake.TabIndex = 5;
            // 
            // label18
            // 
            this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(20, 97);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(71, 13);
            this.label18.TabIndex = 4;
            this.label18.Text = "Год выпуска";
            // 
            // tBoxNameProgram
            // 
            this.tBoxNameProgram.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.tBoxNameProgram.Location = new System.Drawing.Point(174, 71);
            this.tBoxNameProgram.Name = "tBoxNameProgram";
            this.tBoxNameProgram.Size = new System.Drawing.Size(334, 20);
            this.tBoxNameProgram.TabIndex = 3;
            // 
            // label17
            // 
            this.label17.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(20, 73);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(145, 13);
            this.label17.TabIndex = 2;
            this.label17.Text = "Наименование программы";
            // 
            // tBoxNameZakaz
            // 
            this.tBoxNameZakaz.Location = new System.Drawing.Point(174, 15);
            this.tBoxNameZakaz.Multiline = true;
            this.tBoxNameZakaz.Name = "tBoxNameZakaz";
            this.tBoxNameZakaz.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.tBoxNameZakaz.Size = new System.Drawing.Size(334, 50);
            this.tBoxNameZakaz.TabIndex = 1;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(20, 18);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(122, 13);
            this.label12.TabIndex = 0;
            this.label12.Text = "Наименование заказа";
            // 
            // comboBoxChiefDesigner
            // 
            this.comboBoxChiefDesigner.FormattingEnabled = true;
            this.comboBoxChiefDesigner.Location = new System.Drawing.Point(161, 22);
            this.comboBoxChiefDesigner.Name = "comboBoxChiefDesigner";
            this.comboBoxChiefDesigner.Size = new System.Drawing.Size(182, 21);
            this.comboBoxChiefDesigner.TabIndex = 21;
            // 
            // comboBoxChiefVPMO
            // 
            this.comboBoxChiefVPMO.FormattingEnabled = true;
            this.comboBoxChiefVPMO.Location = new System.Drawing.Point(161, 45);
            this.comboBoxChiefVPMO.Name = "comboBoxChiefVPMO";
            this.comboBoxChiefVPMO.Size = new System.Drawing.Size(182, 21);
            this.comboBoxChiefVPMO.TabIndex = 22;
            // 
            // comboBoxMaker
            // 
            this.comboBoxMaker.FormattingEnabled = true;
            this.comboBoxMaker.Location = new System.Drawing.Point(161, 68);
            this.comboBoxMaker.Name = "comboBoxMaker";
            this.comboBoxMaker.Size = new System.Drawing.Size(182, 21);
            this.comboBoxMaker.TabIndex = 23;
            // 
            // comboBoxRateChecker
            // 
            this.comboBoxRateChecker.FormattingEnabled = true;
            this.comboBoxRateChecker.Location = new System.Drawing.Point(161, 92);
            this.comboBoxRateChecker.Name = "comboBoxRateChecker";
            this.comboBoxRateChecker.Size = new System.Drawing.Size(182, 21);
            this.comboBoxRateChecker.TabIndex = 24;
            // 
            // comboBoxChiefTeam
            // 
            this.comboBoxChiefTeam.FormattingEnabled = true;
            this.comboBoxChiefTeam.Location = new System.Drawing.Point(161, 29);
            this.comboBoxChiefTeam.Name = "comboBoxChiefTeam";
            this.comboBoxChiefTeam.Size = new System.Drawing.Size(182, 21);
            this.comboBoxChiefTeam.TabIndex = 24;
            // 
            // comboBoxRepresentative
            // 
            this.comboBoxRepresentative.FormattingEnabled = true;
            this.comboBoxRepresentative.Location = new System.Drawing.Point(161, 52);
            this.comboBoxRepresentative.Name = "comboBoxRepresentative";
            this.comboBoxRepresentative.Size = new System.Drawing.Size(182, 21);
            this.comboBoxRepresentative.TabIndex = 25;
            // 
            // comboBoxZGKZ
            // 
            this.comboBoxZGKZ.FormattingEnabled = true;
            this.comboBoxZGKZ.Location = new System.Drawing.Point(161, 77);
            this.comboBoxZGKZ.Name = "comboBoxZGKZ";
            this.comboBoxZGKZ.Size = new System.Drawing.Size(182, 21);
            this.comboBoxZGKZ.TabIndex = 26;
            // 
            // Attributes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(595, 501);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.makeReportButton);
            this.Name = "Attributes";
            this.Text = "Настройки отчета";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Attributes_FormClosed);
            this.tabControl1.ResumeLayout(false);
            this.tPdocAttributes.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tPTitulLU.ResumeLayout(false);
            this.tPTitulLU.PerformLayout();
            this.grBoxLU.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.grBoxParametr.ResumeLayout(false);
            this.grBoxParametr.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tPdocAttributes;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NumericUpDown nudCountAfter;
        private System.Windows.Forms.NumericUpDown nudCountBefore;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.CheckBox cbFullNames;
        private System.Windows.Forms.CheckBox cbAddChangelogPage;
        private System.Windows.Forms.NumericUpDown nudFirstPageNumber;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbApprovedBy;
        private System.Windows.Forms.TextBox tbNControlBy;
        private System.Windows.Forms.TextBox tbControlledBy;
        private System.Windows.Forms.TextBox tbDesignedBy;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelZagolovok2;
        private System.Windows.Forms.Label labelZagolovok1;
        private System.Windows.Forms.TextBox tbSideBarSignName3;
        private System.Windows.Forms.TextBox tbSideBarSignName2;
        private System.Windows.Forms.TextBox tbSideBarSignName1;
        private System.Windows.Forms.TextBox tbSideBarSignHeader3;
        private System.Windows.Forms.TextBox tbSideBarSignHeader2;
        private System.Windows.Forms.TextBox tbSideBarSignHeader1;
        private TabPage tPTitulLU;
        private CheckBox chBoxLU;
        private CheckBox chBoxTitul;
        private GroupBox grBoxLU;
        private Panel panel2;
        private CheckBox chBoxVPMO;
        private CheckBox chBoxChiefTeam;
        private Label label22;
        private Label label20;
        private Label label19;
        private Panel panel1;
        private CheckBox chBoxMilitaryChief;
        private Label label21;
        private Label label13;
        private Label label16;
        private Label label15;
        private Label label14;
        private GroupBox grBoxParametr;
        public TextBox tBoxYearMake;
        private Label label18;
        private TextBox tBoxNameProgram;
        private Label label17;
        private TextBox tBoxNameZakaz;
        private Label label12;
        private CheckBox chBoxZgkz;
        private Label label23;
        private TextBox textBoxFullName;
        public ComboBox comboBoxZGKZ;
        public ComboBox comboBoxRepresentative;
        public ComboBox comboBoxChiefTeam;
        public ComboBox comboBoxRateChecker;
        public ComboBox comboBoxMaker;
        public ComboBox comboBoxChiefVPMO;
        public ComboBox comboBoxChiefDesigner;
    }
}