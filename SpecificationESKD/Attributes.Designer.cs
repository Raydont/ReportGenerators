using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.Reporting;
using SpecificationESKD.Globus.TFlexDocs;
using System.Data.SqlClient;
using TFlex.DOCs.Model.Structure;

namespace SpecificationESKD
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbSideBarSignName3 = new System.Windows.Forms.ComboBox();
            this.cbSideBarSignName1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.labelZagolovok2 = new System.Windows.Forms.Label();
            this.labelZagolovok1 = new System.Windows.Forms.Label();
            this.tbSideBarSignName2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader3 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader2 = new System.Windows.Forms.TextBox();
            this.tbSideBarSignHeader1 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbControlledBy = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbApprovedBy = new System.Windows.Forms.TextBox();
            this.tbNControlBy = new System.Windows.Forms.TextBox();
            this.tbDesignedBy = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chBoxSortByAlphabet = new System.Windows.Forms.CheckBox();
            this.checkBoxSectionProgramSoft = new System.Windows.Forms.CheckBox();
            this.textBoxFullName = new System.Windows.Forms.TextBox();
            this.cbZeroToDash = new System.Windows.Forms.CheckBox();
            this.nudCountAfter = new System.Windows.Forms.NumericUpDown();
            this.nudCountBefore = new System.Windows.Forms.NumericUpDown();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.cbFullNames = new System.Windows.Forms.CheckBox();
            this.cbAddChangelogPage = new System.Windows.Forms.CheckBox();
            this.nudFirstPageNumber = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.makeReportButton = new System.Windows.Forms.Button();
            this.tabReport = new System.Windows.Forms.TabControl();
            this.tabPagesAttributes = new System.Windows.Forms.TabPage();
            this.chBoxProgramDocToLower = new System.Windows.Forms.CheckBox();
            this.checkBoxSpecForPI = new System.Windows.Forms.CheckBox();
            this.chBoxThroughNumbering = new System.Windows.Forms.CheckBox();
            this.chBoxAddMake = new System.Windows.Forms.CheckBox();
            this.tabPagesAddNewPages = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chBoxNewPageOtherItemsAfterProductsKind = new System.Windows.Forms.CheckBox();
            this.groupBoxNewPagesVarPart = new System.Windows.Forms.GroupBox();
            this.nudKomplektVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudMaterialVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudOtherItemsVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudStdItemsVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudDetailVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudAssemblyVarPart = new System.Windows.Forms.NumericUpDown();
            this.nudDocumentVarPart = new System.Windows.Forms.NumericUpDown();
            this.chBoxKomplectVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxMaterialVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxOtherItemVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxStdItemVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxDetailVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxAssemblyVarPart = new System.Windows.Forms.CheckBox();
            this.chBoxDocVarPart = new System.Windows.Forms.CheckBox();
            this.GBNewPages = new System.Windows.Forms.GroupBox();
            this.nudKomplekt = new System.Windows.Forms.NumericUpDown();
            this.nudMaterials = new System.Windows.Forms.NumericUpDown();
            this.nudOtherProducts = new System.Windows.Forms.NumericUpDown();
            this.nudStdProducts = new System.Windows.Forms.NumericUpDown();
            this.nudDetails = new System.Windows.Forms.NumericUpDown();
            this.nudAssembly = new System.Windows.Forms.NumericUpDown();
            this.nudDocumentation = new System.Windows.Forms.NumericUpDown();
            this.chBoxKomplekt = new System.Windows.Forms.CheckBox();
            this.chBoxMaterials = new System.Windows.Forms.CheckBox();
            this.chBoxOtherProducts = new System.Windows.Forms.CheckBox();
            this.chBoxStdProducts = new System.Windows.Forms.CheckBox();
            this.chBoxDetails = new System.Windows.Forms.CheckBox();
            this.chBoxAssembly = new System.Windows.Forms.CheckBox();
            this.chBoxDocumentation = new System.Windows.Forms.CheckBox();
            this.tabPageTitul = new System.Windows.Forms.TabPage();
            this.tBoxYearMake = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.grBoxVisa = new System.Windows.Forms.GroupBox();
            this.cBoxNorm = new System.Windows.Forms.ComboBox();
            this.cBoxTech = new System.Windows.Forms.ComboBox();
            this.chBoxVPMO = new System.Windows.Forms.CheckBox();
            this.tBoxCheifTeam = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.tBoxVPMO = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.tBox45 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.chBoxAddTitul = new System.Windows.Forms.CheckBox();
            this.grBoxApprSign = new System.Windows.Forms.GroupBox();
            this.chBoxChiefMillitary = new System.Windows.Forms.CheckBox();
            this.tBoxMilitaryChief = new System.Windows.Forms.TextBox();
            this.tBoxGoev = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.tabPageAddModify = new System.Windows.Forms.TabPage();
            this.chBoxAddModify = new System.Windows.Forms.CheckBox();
            this.grBoxModifySpec = new System.Windows.Forms.GroupBox();
            this.label25 = new System.Windows.Forms.Label();
            this.textBoxList2 = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.textBoxListPage2 = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.textBoxList = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.textBoxNumberModify = new System.Windows.Forms.TextBox();
            this.textBoxListPage = new System.Windows.Forms.TextBox();
            this.textBoxDate = new System.Windows.Forms.TextBox();
            this.textBoxNumberDoc = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).BeginInit();
            this.tabReport.SuspendLayout();
            this.tabPagesAttributes.SuspendLayout();
            this.tabPagesAddNewPages.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBoxNewPagesVarPart.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudKomplektVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaterialVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudOtherItemsVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudStdItemsVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDetailVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAssemblyVarPart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDocumentVarPart)).BeginInit();
            this.GBNewPages.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudKomplekt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaterials)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudOtherProducts)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudStdProducts)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDetails)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAssembly)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDocumentation)).BeginInit();
            this.tabPageTitul.SuspendLayout();
            this.grBoxVisa.SuspendLayout();
            this.grBoxApprSign.SuspendLayout();
            this.tabPageAddModify.SuspendLayout();
            this.grBoxModifySpec.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.cbSideBarSignName3);
            this.groupBox1.Controls.Add(this.cbSideBarSignName1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.labelZagolovok2);
            this.groupBox1.Controls.Add(this.labelZagolovok1);
            this.groupBox1.Controls.Add(this.tbSideBarSignName2);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader3);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader2);
            this.groupBox1.Controls.Add(this.tbSideBarSignHeader1);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(548, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Визы на поле подшивки";
            // 
            // cbSideBarSignName3
            // 
            this.cbSideBarSignName3.FormattingEnabled = true;
            this.cbSideBarSignName3.Location = new System.Drawing.Point(349, 72);
            this.cbSideBarSignName3.Name = "cbSideBarSignName3";
            this.cbSideBarSignName3.Size = new System.Drawing.Size(156, 21);
            this.cbSideBarSignName3.TabIndex = 14;
            // 
            // cbSideBarSignName1
            // 
            this.cbSideBarSignName1.FormattingEnabled = true;
            this.cbSideBarSignName1.Location = new System.Drawing.Point(349, 47);
            this.cbSideBarSignName1.Name = "cbSideBarSignName1";
            this.cbSideBarSignName1.Size = new System.Drawing.Size(156, 21);
            this.cbSideBarSignName1.TabIndex = 13;
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
            this.label3.Location = new System.Drawing.Point(277, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Фамилия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(277, 51);
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
            this.labelZagolovok2.Location = new System.Drawing.Point(46, 26);
            this.labelZagolovok2.Name = "labelZagolovok2";
            this.labelZagolovok2.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok2.TabIndex = 8;
            this.labelZagolovok2.Text = "Заголовок";
            // 
            // labelZagolovok1
            // 
            this.labelZagolovok1.AutoSize = true;
            this.labelZagolovok1.Location = new System.Drawing.Point(46, 51);
            this.labelZagolovok1.Name = "labelZagolovok1";
            this.labelZagolovok1.Size = new System.Drawing.Size(61, 13);
            this.labelZagolovok1.TabIndex = 7;
            this.labelZagolovok1.Text = "Заголовок";
            // 
            // tbSideBarSignName2
            // 
            this.tbSideBarSignName2.Location = new System.Drawing.Point(349, 23);
            this.tbSideBarSignName2.Name = "tbSideBarSignName2";
            this.tbSideBarSignName2.Size = new System.Drawing.Size(156, 20);
            this.tbSideBarSignName2.TabIndex = 5;
            // 
            // tbSideBarSignHeader3
            // 
            this.tbSideBarSignHeader3.Location = new System.Drawing.Point(117, 71);
            this.tbSideBarSignHeader3.Name = "tbSideBarSignHeader3";
            this.tbSideBarSignHeader3.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader3.TabIndex = 3;
            this.tbSideBarSignHeader3.Text = "Н.контр.";
            // 
            // tbSideBarSignHeader2
            // 
            this.tbSideBarSignHeader2.Location = new System.Drawing.Point(117, 47);
            this.tbSideBarSignHeader2.Name = "tbSideBarSignHeader2";
            this.tbSideBarSignHeader2.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader2.TabIndex = 2;
            this.tbSideBarSignHeader2.Text = "Т.контр.";
            // 
            // tbSideBarSignHeader1
            // 
            this.tbSideBarSignHeader1.Location = new System.Drawing.Point(117, 23);
            this.tbSideBarSignHeader1.Name = "tbSideBarSignHeader1";
            this.tbSideBarSignHeader1.Size = new System.Drawing.Size(130, 20);
            this.tbSideBarSignHeader1.TabIndex = 1;
            this.tbSideBarSignHeader1.Text = "Разраб.";
            this.tbSideBarSignHeader1.TextChanged += new System.EventHandler(this.tbSideBarSignHeader1_TextChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cbControlledBy);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.tbApprovedBy);
            this.groupBox2.Controls.Add(this.tbNControlBy);
            this.groupBox2.Controls.Add(this.tbDesignedBy);
            this.groupBox2.Location = new System.Drawing.Point(6, 112);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(548, 129);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Подписи основной надписи";
            // 
            // cbControlledBy
            // 
            this.cbControlledBy.FormattingEnabled = true;
            this.cbControlledBy.Location = new System.Drawing.Point(223, 45);
            this.cbControlledBy.Name = "cbControlledBy";
            this.cbControlledBy.Size = new System.Drawing.Size(156, 21);
            this.cbControlledBy.TabIndex = 15;
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
            this.tbApprovedBy.TextChanged += new System.EventHandler(this.tbApprovedBy_TextChanged);
            // 
            // tbNControlBy
            // 
            this.tbNControlBy.Location = new System.Drawing.Point(223, 71);
            this.tbNControlBy.Name = "tbNControlBy";
            this.tbNControlBy.Size = new System.Drawing.Size(156, 20);
            this.tbNControlBy.TabIndex = 7;
            this.tbNControlBy.Text = "Шемчак";
            // 
            // tbDesignedBy
            // 
            this.tbDesignedBy.Location = new System.Drawing.Point(223, 19);
            this.tbDesignedBy.Name = "tbDesignedBy";
            this.tbDesignedBy.Size = new System.Drawing.Size(156, 20);
            this.tbDesignedBy.TabIndex = 5;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chBoxSortByAlphabet);
            this.groupBox3.Controls.Add(this.checkBoxSectionProgramSoft);
            this.groupBox3.Controls.Add(this.textBoxFullName);
            this.groupBox3.Controls.Add(this.cbZeroToDash);
            this.groupBox3.Controls.Add(this.nudCountAfter);
            this.groupBox3.Controls.Add(this.nudCountBefore);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.cbFullNames);
            this.groupBox3.Controls.Add(this.cbAddChangelogPage);
            this.groupBox3.Controls.Add(this.nudFirstPageNumber);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Location = new System.Drawing.Point(6, 247);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(548, 150);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Нумерация листов и содержание";
            // 
            // chBoxSortByAlphabet
            // 
            this.chBoxSortByAlphabet.AutoSize = true;
            this.chBoxSortByAlphabet.Location = new System.Drawing.Point(318, 124);
            this.chBoxSortByAlphabet.Name = "chBoxSortByAlphabet";
            this.chBoxSortByAlphabet.Size = new System.Drawing.Size(222, 17);
            this.chBoxSortByAlphabet.TabIndex = 11;
            this.chBoxSortByAlphabet.Text = "Сортировка докуменации по алфавиту";
            this.chBoxSortByAlphabet.UseVisualStyleBackColor = true;
            this.chBoxSortByAlphabet.CheckedChanged += new System.EventHandler(this.chBoxSortByAlphabet_CheckedChanged);
            // 
            // checkBoxSectionProgramSoft
            // 
            this.checkBoxSectionProgramSoft.AutoSize = true;
            this.checkBoxSectionProgramSoft.Checked = true;
            this.checkBoxSectionProgramSoft.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSectionProgramSoft.Location = new System.Drawing.Point(12, 124);
            this.checkBoxSectionProgramSoft.Name = "checkBoxSectionProgramSoft";
            this.checkBoxSectionProgramSoft.Size = new System.Drawing.Size(287, 17);
            this.checkBoxSectionProgramSoft.TabIndex = 10;
            this.checkBoxSectionProgramSoft.Text = "Добавлять подзаголовок \"Программные изделия\"";
            this.checkBoxSectionProgramSoft.UseVisualStyleBackColor = true;
            // 
            // textBoxFullName
            // 
            this.textBoxFullName.Enabled = false;
            this.textBoxFullName.Location = new System.Drawing.Point(12, 96);
            this.textBoxFullName.Name = "textBoxFullName";
            this.textBoxFullName.Size = new System.Drawing.Size(529, 20);
            this.textBoxFullName.TabIndex = 9;
            // 
            // cbZeroToDash
            // 
            this.cbZeroToDash.AutoSize = true;
            this.cbZeroToDash.Checked = true;
            this.cbZeroToDash.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbZeroToDash.Location = new System.Drawing.Point(256, 76);
            this.cbZeroToDash.Name = "cbZeroToDash";
            this.cbZeroToDash.Size = new System.Drawing.Size(295, 17);
            this.cbZeroToDash.TabIndex = 8;
            this.cbZeroToDash.Text = "\"-\" для позиции 0 (в Стандартных, Прочих и Деталях)";
            this.cbZeroToDash.UseVisualStyleBackColor = true;
            this.cbZeroToDash.CheckedChanged += new System.EventHandler(this.cbZeroToDash_CheckedChanged);
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
            this.cbFullNames.Location = new System.Drawing.Point(13, 76);
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
            1,
            0,
            0,
            0});
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 28);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(117, 13);
            this.label9.TabIndex = 0;
            this.label9.Text = "Номер первого листа";
            // 
            // makeReportButton
            // 
            this.makeReportButton.Location = new System.Drawing.Point(472, 485);
            this.makeReportButton.Name = "makeReportButton";
            this.makeReportButton.Size = new System.Drawing.Size(107, 23);
            this.makeReportButton.TabIndex = 3;
            this.makeReportButton.Text = "Создать отчет";
            this.makeReportButton.UseVisualStyleBackColor = true;
            this.makeReportButton.Click += new System.EventHandler(this.makeReportButton_Click);
            // 
            // tabReport
            // 
            this.tabReport.Controls.Add(this.tabPagesAttributes);
            this.tabReport.Controls.Add(this.tabPagesAddNewPages);
            this.tabReport.Controls.Add(this.tabPageTitul);
            this.tabReport.Controls.Add(this.tabPageAddModify);
            this.tabReport.Location = new System.Drawing.Point(12, 5);
            this.tabReport.Name = "tabReport";
            this.tabReport.SelectedIndex = 0;
            this.tabReport.Size = new System.Drawing.Size(567, 474);
            this.tabReport.TabIndex = 4;
            // 
            // tabPagesAttributes
            // 
            this.tabPagesAttributes.Controls.Add(this.chBoxProgramDocToLower);
            this.tabPagesAttributes.Controls.Add(this.checkBoxSpecForPI);
            this.tabPagesAttributes.Controls.Add(this.chBoxThroughNumbering);
            this.tabPagesAttributes.Controls.Add(this.chBoxAddMake);
            this.tabPagesAttributes.Controls.Add(this.groupBox1);
            this.tabPagesAttributes.Controls.Add(this.groupBox2);
            this.tabPagesAttributes.Controls.Add(this.groupBox3);
            this.tabPagesAttributes.Location = new System.Drawing.Point(4, 22);
            this.tabPagesAttributes.Name = "tabPagesAttributes";
            this.tabPagesAttributes.Padding = new System.Windows.Forms.Padding(3);
            this.tabPagesAttributes.Size = new System.Drawing.Size(559, 448);
            this.tabPagesAttributes.TabIndex = 0;
            this.tabPagesAttributes.Text = "Атрибуты документа";
            this.tabPagesAttributes.UseVisualStyleBackColor = true;
            // 
            // chBoxProgramDocToLower
            // 
            this.chBoxProgramDocToLower.AutoSize = true;
            this.chBoxProgramDocToLower.Location = new System.Drawing.Point(19, 425);
            this.chBoxProgramDocToLower.Name = "chBoxProgramDocToLower";
            this.chBoxProgramDocToLower.Size = new System.Drawing.Size(184, 17);
            this.chBoxProgramDocToLower.TabIndex = 75;
            this.chBoxProgramDocToLower.Text = "Нижний регистр для программ";
            this.chBoxProgramDocToLower.UseVisualStyleBackColor = true;
            // 
            // checkBoxSpecForPI
            // 
            this.checkBoxSpecForPI.AutoSize = true;
            this.checkBoxSpecForPI.Location = new System.Drawing.Point(259, 425);
            this.checkBoxSpecForPI.Name = "checkBoxSpecForPI";
            this.checkBoxSpecForPI.Size = new System.Drawing.Size(281, 17);
            this.checkBoxSpecForPI.TabIndex = 7;
            this.checkBoxSpecForPI.Text = "Спецификация для Предварительного извещения";
            this.checkBoxSpecForPI.UseVisualStyleBackColor = true;
            // 
            // chBoxThroughNumbering
            // 
            this.chBoxThroughNumbering.AutoSize = true;
            this.chBoxThroughNumbering.Location = new System.Drawing.Point(259, 403);
            this.chBoxThroughNumbering.Name = "chBoxThroughNumbering";
            this.chBoxThroughNumbering.Size = new System.Drawing.Size(235, 17);
            this.chBoxThroughNumbering.TabIndex = 6;
            this.chBoxThroughNumbering.Text = "Сквозная нумерация позиций (по ГОСТу)";
            this.chBoxThroughNumbering.UseVisualStyleBackColor = true;
            // 
            // chBoxAddMake
            // 
            this.chBoxAddMake.AutoSize = true;
            this.chBoxAddMake.Checked = true;
            this.chBoxAddMake.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxAddMake.Location = new System.Drawing.Point(19, 403);
            this.chBoxAddMake.Name = "chBoxAddMake";
            this.chBoxAddMake.Size = new System.Drawing.Size(208, 17);
            this.chBoxAddMake.TabIndex = 5;
            this.chBoxAddMake.Text = "Получить групповую спецификацию";
            this.chBoxAddMake.UseVisualStyleBackColor = true;
            this.chBoxAddMake.CheckedChanged += new System.EventHandler(this.chBoxAddMake_CheckedChanged);
            // 
            // tabPagesAddNewPages
            // 
            this.tabPagesAddNewPages.Controls.Add(this.groupBox5);
            this.tabPagesAddNewPages.Controls.Add(this.groupBox4);
            this.tabPagesAddNewPages.Controls.Add(this.groupBoxNewPagesVarPart);
            this.tabPagesAddNewPages.Controls.Add(this.GBNewPages);
            this.tabPagesAddNewPages.Location = new System.Drawing.Point(4, 22);
            this.tabPagesAddNewPages.Name = "tabPagesAddNewPages";
            this.tabPagesAddNewPages.Padding = new System.Windows.Forms.Padding(3);
            this.tabPagesAddNewPages.Size = new System.Drawing.Size(559, 448);
            this.tabPagesAddNewPages.TabIndex = 1;
            this.tabPagesAddNewPages.Text = "Резервные страницы";
            this.tabPagesAddNewPages.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart);
            this.groupBox5.Location = new System.Drawing.Point(285, 291);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(256, 56);
            this.groupBox5.TabIndex = 7;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Для переменной части";
            // 
            // chBoxNewPageOtherItemsAfterProductsKindVaiosPart
            // 
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.AutoSize = true;
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.Location = new System.Drawing.Point(5, 24);
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.Name = "chBoxNewPageOtherItemsAfterProductsKindVaiosPart";
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.Size = new System.Drawing.Size(245, 17);
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.TabIndex = 1;
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.Text = "Вид изделий с новой страницы в Проч.изд.";
            this.chBoxNewPageOtherItemsAfterProductsKindVaiosPart.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.chBoxNewPageOtherItemsAfterProductsKind);
            this.groupBox4.Location = new System.Drawing.Point(6, 291);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(256, 56);
            this.groupBox4.TabIndex = 6;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Для постоянной части";
            // 
            // chBoxNewPageOtherItemsAfterProductsKind
            // 
            this.chBoxNewPageOtherItemsAfterProductsKind.AutoSize = true;
            this.chBoxNewPageOtherItemsAfterProductsKind.Location = new System.Drawing.Point(5, 24);
            this.chBoxNewPageOtherItemsAfterProductsKind.Name = "chBoxNewPageOtherItemsAfterProductsKind";
            this.chBoxNewPageOtherItemsAfterProductsKind.Size = new System.Drawing.Size(245, 17);
            this.chBoxNewPageOtherItemsAfterProductsKind.TabIndex = 1;
            this.chBoxNewPageOtherItemsAfterProductsKind.Text = "Вид изделий с новой страницы в Проч.изд.";
            this.chBoxNewPageOtherItemsAfterProductsKind.UseVisualStyleBackColor = true;
            this.chBoxNewPageOtherItemsAfterProductsKind.CheckedChanged += new System.EventHandler(this.chBoxNewPageOtherItemsAfterProductsKind_CheckedChanged);
            // 
            // groupBoxNewPagesVarPart
            // 
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudKomplektVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudMaterialVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudOtherItemsVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudStdItemsVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudDetailVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudAssemblyVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.nudDocumentVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxKomplectVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxMaterialVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxOtherItemVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxStdItemVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxDetailVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxAssemblyVarPart);
            this.groupBoxNewPagesVarPart.Controls.Add(this.chBoxDocVarPart);
            this.groupBoxNewPagesVarPart.Location = new System.Drawing.Point(285, 17);
            this.groupBoxNewPagesVarPart.Name = "groupBoxNewPagesVarPart";
            this.groupBoxNewPagesVarPart.Size = new System.Drawing.Size(256, 268);
            this.groupBoxNewPagesVarPart.TabIndex = 4;
            this.groupBoxNewPagesVarPart.TabStop = false;
            this.groupBoxNewPagesVarPart.Text = "Добавление резервных страниц для переменной части";
            // 
            // nudKomplektVarPart
            // 
            this.nudKomplektVarPart.Enabled = false;
            this.nudKomplektVarPart.Location = new System.Drawing.Point(158, 232);
            this.nudKomplektVarPart.Name = "nudKomplektVarPart";
            this.nudKomplektVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudKomplektVarPart.TabIndex = 13;
            // 
            // nudMaterialVarPart
            // 
            this.nudMaterialVarPart.Enabled = false;
            this.nudMaterialVarPart.Location = new System.Drawing.Point(158, 199);
            this.nudMaterialVarPart.Name = "nudMaterialVarPart";
            this.nudMaterialVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudMaterialVarPart.TabIndex = 12;
            // 
            // nudOtherItemsVarPart
            // 
            this.nudOtherItemsVarPart.Enabled = false;
            this.nudOtherItemsVarPart.Location = new System.Drawing.Point(158, 164);
            this.nudOtherItemsVarPart.Name = "nudOtherItemsVarPart";
            this.nudOtherItemsVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudOtherItemsVarPart.TabIndex = 11;
            // 
            // nudStdItemsVarPart
            // 
            this.nudStdItemsVarPart.Enabled = false;
            this.nudStdItemsVarPart.Location = new System.Drawing.Point(158, 131);
            this.nudStdItemsVarPart.Name = "nudStdItemsVarPart";
            this.nudStdItemsVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudStdItemsVarPart.TabIndex = 10;
            // 
            // nudDetailVarPart
            // 
            this.nudDetailVarPart.Enabled = false;
            this.nudDetailVarPart.Location = new System.Drawing.Point(158, 99);
            this.nudDetailVarPart.Name = "nudDetailVarPart";
            this.nudDetailVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudDetailVarPart.TabIndex = 9;
            // 
            // nudAssemblyVarPart
            // 
            this.nudAssemblyVarPart.Enabled = false;
            this.nudAssemblyVarPart.Location = new System.Drawing.Point(158, 66);
            this.nudAssemblyVarPart.Name = "nudAssemblyVarPart";
            this.nudAssemblyVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudAssemblyVarPart.TabIndex = 8;
            // 
            // nudDocumentVarPart
            // 
            this.nudDocumentVarPart.Enabled = false;
            this.nudDocumentVarPart.Location = new System.Drawing.Point(158, 34);
            this.nudDocumentVarPart.Name = "nudDocumentVarPart";
            this.nudDocumentVarPart.Size = new System.Drawing.Size(44, 20);
            this.nudDocumentVarPart.TabIndex = 7;
            // 
            // chBoxKomplectVarPart
            // 
            this.chBoxKomplectVarPart.AutoSize = true;
            this.chBoxKomplectVarPart.Location = new System.Drawing.Point(16, 233);
            this.chBoxKomplectVarPart.Name = "chBoxKomplectVarPart";
            this.chBoxKomplectVarPart.Size = new System.Drawing.Size(84, 17);
            this.chBoxKomplectVarPart.TabIndex = 6;
            this.chBoxKomplectVarPart.Text = "Комплекты";
            this.chBoxKomplectVarPart.UseVisualStyleBackColor = true;
            this.chBoxKomplectVarPart.CheckedChanged += new System.EventHandler(this.chBoxKomplectVarPart_CheckedChanged);
            // 
            // chBoxMaterialVarPart
            // 
            this.chBoxMaterialVarPart.AutoSize = true;
            this.chBoxMaterialVarPart.Location = new System.Drawing.Point(16, 200);
            this.chBoxMaterialVarPart.Name = "chBoxMaterialVarPart";
            this.chBoxMaterialVarPart.Size = new System.Drawing.Size(84, 17);
            this.chBoxMaterialVarPart.TabIndex = 5;
            this.chBoxMaterialVarPart.Text = "Материалы";
            this.chBoxMaterialVarPart.UseVisualStyleBackColor = true;
            this.chBoxMaterialVarPart.CheckedChanged += new System.EventHandler(this.chBoxMaterialVarPart_CheckedChanged);
            // 
            // chBoxOtherItemVarPart
            // 
            this.chBoxOtherItemVarPart.AutoSize = true;
            this.chBoxOtherItemVarPart.Location = new System.Drawing.Point(16, 165);
            this.chBoxOtherItemVarPart.Name = "chBoxOtherItemVarPart";
            this.chBoxOtherItemVarPart.Size = new System.Drawing.Size(108, 17);
            this.chBoxOtherItemVarPart.TabIndex = 4;
            this.chBoxOtherItemVarPart.Text = "Прочие изделия";
            this.chBoxOtherItemVarPart.UseVisualStyleBackColor = true;
            this.chBoxOtherItemVarPart.CheckedChanged += new System.EventHandler(this.chBoxOtherItemVarPart_CheckedChanged);
            // 
            // chBoxStdItemVarPart
            // 
            this.chBoxStdItemVarPart.AutoSize = true;
            this.chBoxStdItemVarPart.Location = new System.Drawing.Point(16, 132);
            this.chBoxStdItemVarPart.Name = "chBoxStdItemVarPart";
            this.chBoxStdItemVarPart.Size = new System.Drawing.Size(138, 17);
            this.chBoxStdItemVarPart.TabIndex = 3;
            this.chBoxStdItemVarPart.Text = "Стандартные изделия";
            this.chBoxStdItemVarPart.UseVisualStyleBackColor = true;
            this.chBoxStdItemVarPart.CheckedChanged += new System.EventHandler(this.chBoxStdItemVarPart_CheckedChanged);
            // 
            // chBoxDetailVarPart
            // 
            this.chBoxDetailVarPart.AutoSize = true;
            this.chBoxDetailVarPart.Location = new System.Drawing.Point(16, 100);
            this.chBoxDetailVarPart.Name = "chBoxDetailVarPart";
            this.chBoxDetailVarPart.Size = new System.Drawing.Size(64, 17);
            this.chBoxDetailVarPart.TabIndex = 2;
            this.chBoxDetailVarPart.Text = "Детали";
            this.chBoxDetailVarPart.UseVisualStyleBackColor = true;
            this.chBoxDetailVarPart.CheckedChanged += new System.EventHandler(this.chBoxDetailVarPart_CheckedChanged);
            // 
            // chBoxAssemblyVarPart
            // 
            this.chBoxAssemblyVarPart.AutoSize = true;
            this.chBoxAssemblyVarPart.Location = new System.Drawing.Point(16, 67);
            this.chBoxAssemblyVarPart.Name = "chBoxAssemblyVarPart";
            this.chBoxAssemblyVarPart.Size = new System.Drawing.Size(129, 17);
            this.chBoxAssemblyVarPart.TabIndex = 1;
            this.chBoxAssemblyVarPart.Text = "Сборочные единицы";
            this.chBoxAssemblyVarPart.UseVisualStyleBackColor = true;
            this.chBoxAssemblyVarPart.CheckedChanged += new System.EventHandler(this.chBoxAssemblyVarPart_CheckedChanged);
            // 
            // chBoxDocVarPart
            // 
            this.chBoxDocVarPart.AutoSize = true;
            this.chBoxDocVarPart.Location = new System.Drawing.Point(16, 35);
            this.chBoxDocVarPart.Name = "chBoxDocVarPart";
            this.chBoxDocVarPart.Size = new System.Drawing.Size(101, 17);
            this.chBoxDocVarPart.TabIndex = 0;
            this.chBoxDocVarPart.Text = "Документация";
            this.chBoxDocVarPart.UseVisualStyleBackColor = true;
            this.chBoxDocVarPart.CheckedChanged += new System.EventHandler(this.chBoxDocVarPart_CheckedChanged);
            // 
            // GBNewPages
            // 
            this.GBNewPages.Controls.Add(this.nudKomplekt);
            this.GBNewPages.Controls.Add(this.nudMaterials);
            this.GBNewPages.Controls.Add(this.nudOtherProducts);
            this.GBNewPages.Controls.Add(this.nudStdProducts);
            this.GBNewPages.Controls.Add(this.nudDetails);
            this.GBNewPages.Controls.Add(this.nudAssembly);
            this.GBNewPages.Controls.Add(this.nudDocumentation);
            this.GBNewPages.Controls.Add(this.chBoxKomplekt);
            this.GBNewPages.Controls.Add(this.chBoxMaterials);
            this.GBNewPages.Controls.Add(this.chBoxOtherProducts);
            this.GBNewPages.Controls.Add(this.chBoxStdProducts);
            this.GBNewPages.Controls.Add(this.chBoxDetails);
            this.GBNewPages.Controls.Add(this.chBoxAssembly);
            this.GBNewPages.Controls.Add(this.chBoxDocumentation);
            this.GBNewPages.Location = new System.Drawing.Point(6, 17);
            this.GBNewPages.Name = "GBNewPages";
            this.GBNewPages.Size = new System.Drawing.Size(256, 268);
            this.GBNewPages.TabIndex = 3;
            this.GBNewPages.TabStop = false;
            this.GBNewPages.Text = "Добавление резервных страниц для постоянной части";
            // 
            // nudKomplekt
            // 
            this.nudKomplekt.Enabled = false;
            this.nudKomplekt.Location = new System.Drawing.Point(158, 232);
            this.nudKomplekt.Name = "nudKomplekt";
            this.nudKomplekt.Size = new System.Drawing.Size(44, 20);
            this.nudKomplekt.TabIndex = 13;
            // 
            // nudMaterials
            // 
            this.nudMaterials.Enabled = false;
            this.nudMaterials.Location = new System.Drawing.Point(158, 199);
            this.nudMaterials.Name = "nudMaterials";
            this.nudMaterials.Size = new System.Drawing.Size(44, 20);
            this.nudMaterials.TabIndex = 12;
            // 
            // nudOtherProducts
            // 
            this.nudOtherProducts.Enabled = false;
            this.nudOtherProducts.Location = new System.Drawing.Point(158, 164);
            this.nudOtherProducts.Name = "nudOtherProducts";
            this.nudOtherProducts.Size = new System.Drawing.Size(44, 20);
            this.nudOtherProducts.TabIndex = 11;
            // 
            // nudStdProducts
            // 
            this.nudStdProducts.Enabled = false;
            this.nudStdProducts.Location = new System.Drawing.Point(158, 131);
            this.nudStdProducts.Name = "nudStdProducts";
            this.nudStdProducts.Size = new System.Drawing.Size(44, 20);
            this.nudStdProducts.TabIndex = 10;
            // 
            // nudDetails
            // 
            this.nudDetails.Enabled = false;
            this.nudDetails.Location = new System.Drawing.Point(158, 99);
            this.nudDetails.Name = "nudDetails";
            this.nudDetails.Size = new System.Drawing.Size(44, 20);
            this.nudDetails.TabIndex = 9;
            // 
            // nudAssembly
            // 
            this.nudAssembly.Enabled = false;
            this.nudAssembly.Location = new System.Drawing.Point(158, 66);
            this.nudAssembly.Name = "nudAssembly";
            this.nudAssembly.Size = new System.Drawing.Size(44, 20);
            this.nudAssembly.TabIndex = 8;
            // 
            // nudDocumentation
            // 
            this.nudDocumentation.Enabled = false;
            this.nudDocumentation.Location = new System.Drawing.Point(158, 34);
            this.nudDocumentation.Name = "nudDocumentation";
            this.nudDocumentation.Size = new System.Drawing.Size(44, 20);
            this.nudDocumentation.TabIndex = 7;
            // 
            // chBoxKomplekt
            // 
            this.chBoxKomplekt.AutoSize = true;
            this.chBoxKomplekt.Location = new System.Drawing.Point(16, 233);
            this.chBoxKomplekt.Name = "chBoxKomplekt";
            this.chBoxKomplekt.Size = new System.Drawing.Size(84, 17);
            this.chBoxKomplekt.TabIndex = 6;
            this.chBoxKomplekt.Text = "Комплекты";
            this.chBoxKomplekt.UseVisualStyleBackColor = true;
            this.chBoxKomplekt.CheckedChanged += new System.EventHandler(this.chBoxKomplekt_CheckedChanged);
            // 
            // chBoxMaterials
            // 
            this.chBoxMaterials.AutoSize = true;
            this.chBoxMaterials.Location = new System.Drawing.Point(16, 200);
            this.chBoxMaterials.Name = "chBoxMaterials";
            this.chBoxMaterials.Size = new System.Drawing.Size(84, 17);
            this.chBoxMaterials.TabIndex = 5;
            this.chBoxMaterials.Text = "Материалы";
            this.chBoxMaterials.UseVisualStyleBackColor = true;
            this.chBoxMaterials.CheckedChanged += new System.EventHandler(this.chBoxMaterials_CheckedChanged);
            // 
            // chBoxOtherProducts
            // 
            this.chBoxOtherProducts.AutoSize = true;
            this.chBoxOtherProducts.Location = new System.Drawing.Point(16, 165);
            this.chBoxOtherProducts.Name = "chBoxOtherProducts";
            this.chBoxOtherProducts.Size = new System.Drawing.Size(108, 17);
            this.chBoxOtherProducts.TabIndex = 4;
            this.chBoxOtherProducts.Text = "Прочие изделия";
            this.chBoxOtherProducts.UseVisualStyleBackColor = true;
            this.chBoxOtherProducts.CheckedChanged += new System.EventHandler(this.chBoxOtherProducts_CheckedChanged);
            // 
            // chBoxStdProducts
            // 
            this.chBoxStdProducts.AutoSize = true;
            this.chBoxStdProducts.Location = new System.Drawing.Point(16, 132);
            this.chBoxStdProducts.Name = "chBoxStdProducts";
            this.chBoxStdProducts.Size = new System.Drawing.Size(138, 17);
            this.chBoxStdProducts.TabIndex = 3;
            this.chBoxStdProducts.Text = "Стандартные изделия";
            this.chBoxStdProducts.UseVisualStyleBackColor = true;
            this.chBoxStdProducts.CheckedChanged += new System.EventHandler(this.chBoxStdProducts_CheckedChanged);
            // 
            // chBoxDetails
            // 
            this.chBoxDetails.AutoSize = true;
            this.chBoxDetails.Location = new System.Drawing.Point(16, 100);
            this.chBoxDetails.Name = "chBoxDetails";
            this.chBoxDetails.Size = new System.Drawing.Size(64, 17);
            this.chBoxDetails.TabIndex = 2;
            this.chBoxDetails.Text = "Детали";
            this.chBoxDetails.UseVisualStyleBackColor = true;
            this.chBoxDetails.CheckedChanged += new System.EventHandler(this.chBoxDetails_CheckedChanged);
            // 
            // chBoxAssembly
            // 
            this.chBoxAssembly.AutoSize = true;
            this.chBoxAssembly.Location = new System.Drawing.Point(16, 67);
            this.chBoxAssembly.Name = "chBoxAssembly";
            this.chBoxAssembly.Size = new System.Drawing.Size(129, 17);
            this.chBoxAssembly.TabIndex = 1;
            this.chBoxAssembly.Text = "Сборочные единицы";
            this.chBoxAssembly.UseVisualStyleBackColor = true;
            this.chBoxAssembly.CheckedChanged += new System.EventHandler(this.chBoxAssembly_CheckedChanged);
            // 
            // chBoxDocumentation
            // 
            this.chBoxDocumentation.AutoSize = true;
            this.chBoxDocumentation.Location = new System.Drawing.Point(16, 35);
            this.chBoxDocumentation.Name = "chBoxDocumentation";
            this.chBoxDocumentation.Size = new System.Drawing.Size(101, 17);
            this.chBoxDocumentation.TabIndex = 0;
            this.chBoxDocumentation.Text = "Документация";
            this.chBoxDocumentation.UseVisualStyleBackColor = true;
            this.chBoxDocumentation.CheckedChanged += new System.EventHandler(this.chBoxDocumentation_CheckedChanged);
            // 
            // tabPageTitul
            // 
            this.tabPageTitul.Controls.Add(this.tBoxYearMake);
            this.tabPageTitul.Controls.Add(this.label19);
            this.tabPageTitul.Controls.Add(this.grBoxVisa);
            this.tabPageTitul.Controls.Add(this.chBoxAddTitul);
            this.tabPageTitul.Controls.Add(this.grBoxApprSign);
            this.tabPageTitul.Location = new System.Drawing.Point(4, 22);
            this.tabPageTitul.Name = "tabPageTitul";
            this.tabPageTitul.Size = new System.Drawing.Size(559, 448);
            this.tabPageTitul.TabIndex = 2;
            this.tabPageTitul.Text = "Атрибуты титульного листа";
            this.tabPageTitul.UseVisualStyleBackColor = true;
            // 
            // tBoxYearMake
            // 
            this.tBoxYearMake.Enabled = false;
            this.tBoxYearMake.Location = new System.Drawing.Point(249, 309);
            this.tBoxYearMake.Name = "tBoxYearMake";
            this.tBoxYearMake.Size = new System.Drawing.Size(135, 20);
            this.tBoxYearMake.TabIndex = 10;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(162, 312);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(71, 13);
            this.label19.TabIndex = 3;
            this.label19.Text = "Год выпуска";
            // 
            // grBoxVisa
            // 
            this.grBoxVisa.Controls.Add(this.cBoxNorm);
            this.grBoxVisa.Controls.Add(this.cBoxTech);
            this.grBoxVisa.Controls.Add(this.chBoxVPMO);
            this.grBoxVisa.Controls.Add(this.tBoxCheifTeam);
            this.grBoxVisa.Controls.Add(this.label18);
            this.grBoxVisa.Controls.Add(this.tBoxVPMO);
            this.grBoxVisa.Controls.Add(this.label16);
            this.grBoxVisa.Controls.Add(this.label17);
            this.grBoxVisa.Controls.Add(this.tBox45);
            this.grBoxVisa.Controls.Add(this.label14);
            this.grBoxVisa.Controls.Add(this.label15);
            this.grBoxVisa.Location = new System.Drawing.Point(13, 142);
            this.grBoxVisa.Name = "grBoxVisa";
            this.grBoxVisa.Size = new System.Drawing.Size(529, 150);
            this.grBoxVisa.TabIndex = 2;
            this.grBoxVisa.TabStop = false;
            this.grBoxVisa.Text = "Визы на поле подшивки";
            // 
            // cBoxNorm
            // 
            this.cBoxNorm.Enabled = false;
            this.cBoxNorm.FormattingEnabled = true;
            this.cBoxNorm.Location = new System.Drawing.Point(152, 73);
            this.cBoxNorm.Name = "cBoxNorm";
            this.cBoxNorm.Size = new System.Drawing.Size(301, 21);
            this.cBoxNorm.TabIndex = 13;
            // 
            // cBoxTech
            // 
            this.cBoxTech.Enabled = false;
            this.cBoxTech.FormattingEnabled = true;
            this.cBoxTech.Location = new System.Drawing.Point(152, 19);
            this.cBoxTech.Name = "cBoxTech";
            this.cBoxTech.Size = new System.Drawing.Size(301, 21);
            this.cBoxTech.TabIndex = 12;
            // 
            // chBoxVPMO
            // 
            this.chBoxVPMO.AutoSize = true;
            this.chBoxVPMO.Checked = true;
            this.chBoxVPMO.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxVPMO.Enabled = false;
            this.chBoxVPMO.Location = new System.Drawing.Point(463, 104);
            this.chBoxVPMO.Name = "chBoxVPMO";
            this.chBoxVPMO.Size = new System.Drawing.Size(15, 14);
            this.chBoxVPMO.TabIndex = 10;
            this.chBoxVPMO.UseVisualStyleBackColor = true;
            this.chBoxVPMO.CheckedChanged += new System.EventHandler(this.chBoxVPMO_CheckedChanged);
            // 
            // tBoxCheifTeam
            // 
            this.tBoxCheifTeam.Enabled = false;
            this.tBoxCheifTeam.Location = new System.Drawing.Point(152, 124);
            this.tBoxCheifTeam.Name = "tBoxCheifTeam";
            this.tBoxCheifTeam.Size = new System.Drawing.Size(301, 20);
            this.tBoxCheifTeam.TabIndex = 9;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(38, 127);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(108, 13);
            this.label18.TabIndex = 8;
            this.label18.Text = "Начальник бригады";
            // 
            // tBoxVPMO
            // 
            this.tBoxVPMO.Enabled = false;
            this.tBoxVPMO.Location = new System.Drawing.Point(152, 100);
            this.tBoxVPMO.Name = "tBoxVPMO";
            this.tBoxVPMO.Size = new System.Drawing.Size(301, 20);
            this.tBoxVPMO.TabIndex = 7;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(38, 103);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(42, 13);
            this.label16.TabIndex = 5;
            this.label16.Text = "ВП МО";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(38, 77);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(50, 13);
            this.label17.TabIndex = 4;
            this.label17.Text = "Н. контр";
            // 
            // tBox45
            // 
            this.tBox45.Enabled = false;
            this.tBox45.Location = new System.Drawing.Point(152, 48);
            this.tBox45.Name = "tBox45";
            this.tBox45.Size = new System.Drawing.Size(301, 20);
            this.tBox45.TabIndex = 3;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(38, 51);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(53, 13);
            this.label14.TabIndex = 1;
            this.label14.Text = "Отдел 45";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(38, 25);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(54, 13);
            this.label15.TabIndex = 0;
            this.label15.Text = "Технолог";
            // 
            // chBoxAddTitul
            // 
            this.chBoxAddTitul.AutoSize = true;
            this.chBoxAddTitul.Location = new System.Drawing.Point(190, 21);
            this.chBoxAddTitul.Name = "chBoxAddTitul";
            this.chBoxAddTitul.Size = new System.Drawing.Size(158, 17);
            this.chBoxAddTitul.TabIndex = 1;
            this.chBoxAddTitul.Text = "Добавить титульный лист";
            this.chBoxAddTitul.UseVisualStyleBackColor = true;
            this.chBoxAddTitul.CheckedChanged += new System.EventHandler(this.chBoxAddTitul_CheckedChanged);
            // 
            // grBoxApprSign
            // 
            this.grBoxApprSign.Controls.Add(this.chBoxChiefMillitary);
            this.grBoxApprSign.Controls.Add(this.tBoxMilitaryChief);
            this.grBoxApprSign.Controls.Add(this.tBoxGoev);
            this.grBoxApprSign.Controls.Add(this.label13);
            this.grBoxApprSign.Controls.Add(this.label12);
            this.grBoxApprSign.Location = new System.Drawing.Point(13, 44);
            this.grBoxApprSign.Name = "grBoxApprSign";
            this.grBoxApprSign.Size = new System.Drawing.Size(529, 93);
            this.grBoxApprSign.TabIndex = 0;
            this.grBoxApprSign.TabStop = false;
            this.grBoxApprSign.Text = "Подписи на титульном листе";
            // 
            // chBoxChiefMillitary
            // 
            this.chBoxChiefMillitary.AutoSize = true;
            this.chBoxChiefMillitary.Checked = true;
            this.chBoxChiefMillitary.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chBoxChiefMillitary.Enabled = false;
            this.chBoxChiefMillitary.Location = new System.Drawing.Point(463, 57);
            this.chBoxChiefMillitary.Name = "chBoxChiefMillitary";
            this.chBoxChiefMillitary.Size = new System.Drawing.Size(15, 14);
            this.chBoxChiefMillitary.TabIndex = 4;
            this.chBoxChiefMillitary.UseVisualStyleBackColor = true;
            this.chBoxChiefMillitary.CheckedChanged += new System.EventHandler(this.chBoxChiefMillitary_CheckedChanged);
            // 
            // tBoxMilitaryChief
            // 
            this.tBoxMilitaryChief.Enabled = false;
            this.tBoxMilitaryChief.Location = new System.Drawing.Point(152, 54);
            this.tBoxMilitaryChief.Name = "tBoxMilitaryChief";
            this.tBoxMilitaryChief.Size = new System.Drawing.Size(301, 20);
            this.tBoxMilitaryChief.TabIndex = 3;
            this.tBoxMilitaryChief.Text = "Д.Е. Покидышев";
            // 
            // tBoxGoev
            // 
            this.tBoxGoev.Enabled = false;
            this.tBoxGoev.Location = new System.Drawing.Point(152, 22);
            this.tBoxGoev.Name = "tBoxGoev";
            this.tBoxGoev.Size = new System.Drawing.Size(301, 20);
            this.tBoxGoev.TabIndex = 2;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(16, 57);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(100, 13);
            this.label13.TabIndex = 1;
            this.label13.Text = "Начальник ВП МО";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(16, 25);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(117, 13);
            this.label12.TabIndex = 0;
            this.label12.Text = "Главный конструктор";
            // 
            // tabPageAddModify
            // 
            this.tabPageAddModify.Controls.Add(this.chBoxAddModify);
            this.tabPageAddModify.Controls.Add(this.grBoxModifySpec);
            this.tabPageAddModify.Location = new System.Drawing.Point(4, 22);
            this.tabPageAddModify.Name = "tabPageAddModify";
            this.tabPageAddModify.Size = new System.Drawing.Size(559, 448);
            this.tabPageAddModify.TabIndex = 3;
            this.tabPageAddModify.Text = "Добавить изменения";
            this.tabPageAddModify.UseVisualStyleBackColor = true;
            // 
            // chBoxAddModify
            // 
            this.chBoxAddModify.AutoSize = true;
            this.chBoxAddModify.Location = new System.Drawing.Point(197, 32);
            this.chBoxAddModify.Name = "chBoxAddModify";
            this.chBoxAddModify.Size = new System.Drawing.Size(135, 17);
            this.chBoxAddModify.TabIndex = 7;
            this.chBoxAddModify.Text = "Добавить изменения";
            this.chBoxAddModify.UseVisualStyleBackColor = true;
            this.chBoxAddModify.CheckedChanged += new System.EventHandler(this.chBoxAddModify_CheckedChanged);
            // 
            // grBoxModifySpec
            // 
            this.grBoxModifySpec.Controls.Add(this.label25);
            this.grBoxModifySpec.Controls.Add(this.textBoxList2);
            this.grBoxModifySpec.Controls.Add(this.label26);
            this.grBoxModifySpec.Controls.Add(this.textBoxListPage2);
            this.grBoxModifySpec.Controls.Add(this.label24);
            this.grBoxModifySpec.Controls.Add(this.textBoxList);
            this.grBoxModifySpec.Controls.Add(this.label23);
            this.grBoxModifySpec.Controls.Add(this.label22);
            this.grBoxModifySpec.Controls.Add(this.label21);
            this.grBoxModifySpec.Controls.Add(this.label20);
            this.grBoxModifySpec.Controls.Add(this.textBoxNumberModify);
            this.grBoxModifySpec.Controls.Add(this.textBoxListPage);
            this.grBoxModifySpec.Controls.Add(this.textBoxDate);
            this.grBoxModifySpec.Controls.Add(this.textBoxNumberDoc);
            this.grBoxModifySpec.Enabled = false;
            this.grBoxModifySpec.Location = new System.Drawing.Point(21, 67);
            this.grBoxModifySpec.Name = "grBoxModifySpec";
            this.grBoxModifySpec.Size = new System.Drawing.Size(507, 223);
            this.grBoxModifySpec.TabIndex = 6;
            this.grBoxModifySpec.TabStop = false;
            this.grBoxModifySpec.Text = "Листы и информация об изменених";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(23, 162);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(32, 13);
            this.label25.TabIndex = 17;
            this.label25.Text = "Лист";
            // 
            // textBoxList2
            // 
            this.textBoxList2.Location = new System.Drawing.Point(194, 159);
            this.textBoxList2.Name = "textBoxList2";
            this.textBoxList2.Size = new System.Drawing.Size(152, 20);
            this.textBoxList2.TabIndex = 16;
            this.textBoxList2.Text = "Нов.";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(24, 188);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(168, 13);
            this.label26.TabIndex = 15;
            this.label26.Text = "Список листов (через зпт и \"-\" )";
            this.label26.Click += new System.EventHandler(this.label26_Click);
            // 
            // textBoxListPage2
            // 
            this.textBoxListPage2.Location = new System.Drawing.Point(194, 185);
            this.textBoxListPage2.Name = "textBoxListPage2";
            this.textBoxListPage2.Size = new System.Drawing.Size(284, 20);
            this.textBoxListPage2.TabIndex = 14;
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(23, 109);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(32, 13);
            this.label24.TabIndex = 9;
            this.label24.Text = "Лист";
            // 
            // textBoxList
            // 
            this.textBoxList.Location = new System.Drawing.Point(194, 106);
            this.textBoxList.Name = "textBoxList";
            this.textBoxList.Size = new System.Drawing.Size(152, 20);
            this.textBoxList.TabIndex = 8;
            this.textBoxList.Text = "Зам.";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(23, 136);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(168, 13);
            this.label23.TabIndex = 7;
            this.label23.Text = "Список листов (через зпт и \"-\" )";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(23, 84);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(92, 13);
            this.label22.TabIndex = 6;
            this.label22.Text = "Дата изменения";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(23, 58);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(75, 13);
            this.label21.TabIndex = 5;
            this.label21.Text = "№ документа";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(23, 32);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(100, 13);
            this.label20.TabIndex = 4;
            this.label20.Text = "Номер изменения";
            // 
            // textBoxNumberModify
            // 
            this.textBoxNumberModify.Location = new System.Drawing.Point(194, 30);
            this.textBoxNumberModify.Name = "textBoxNumberModify";
            this.textBoxNumberModify.Size = new System.Drawing.Size(152, 20);
            this.textBoxNumberModify.TabIndex = 3;
            // 
            // textBoxListPage
            // 
            this.textBoxListPage.Location = new System.Drawing.Point(194, 133);
            this.textBoxListPage.Name = "textBoxListPage";
            this.textBoxListPage.Size = new System.Drawing.Size(284, 20);
            this.textBoxListPage.TabIndex = 0;
            // 
            // textBoxDate
            // 
            this.textBoxDate.Location = new System.Drawing.Point(194, 81);
            this.textBoxDate.Name = "textBoxDate";
            this.textBoxDate.Size = new System.Drawing.Size(152, 20);
            this.textBoxDate.TabIndex = 1;
            // 
            // textBoxNumberDoc
            // 
            this.textBoxNumberDoc.Location = new System.Drawing.Point(194, 55);
            this.textBoxNumberDoc.Name = "textBoxNumberDoc";
            this.textBoxNumberDoc.Size = new System.Drawing.Size(152, 20);
            this.textBoxNumberDoc.TabIndex = 2;
            // 
            // Attributes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(586, 522);
            this.Controls.Add(this.tabReport);
            this.Controls.Add(this.makeReportButton);
            this.Name = "Attributes";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки отчета";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Attributes_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountAfter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudCountBefore)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudFirstPageNumber)).EndInit();
            this.tabReport.ResumeLayout(false);
            this.tabPagesAttributes.ResumeLayout(false);
            this.tabPagesAttributes.PerformLayout();
            this.tabPagesAddNewPages.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBoxNewPagesVarPart.ResumeLayout(false);
            this.groupBoxNewPagesVarPart.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudKomplektVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaterialVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudOtherItemsVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudStdItemsVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDetailVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAssemblyVarPart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDocumentVarPart)).EndInit();
            this.GBNewPages.ResumeLayout(false);
            this.GBNewPages.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudKomplekt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudMaterials)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudOtherProducts)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudStdProducts)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDetails)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudAssembly)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nudDocumentation)).EndInit();
            this.tabPageTitul.ResumeLayout(false);
            this.tabPageTitul.PerformLayout();
            this.grBoxVisa.ResumeLayout(false);
            this.grBoxVisa.PerformLayout();
            this.grBoxApprSign.ResumeLayout(false);
            this.grBoxApprSign.PerformLayout();
            this.tabPageAddModify.ResumeLayout(false);
            this.tabPageAddModify.PerformLayout();
            this.grBoxModifySpec.ResumeLayout(false);
            this.grBoxModifySpec.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelZagolovok2;
        private System.Windows.Forms.Label labelZagolovok1;
        private System.Windows.Forms.TextBox tbSideBarSignName2;
        private System.Windows.Forms.TextBox tbSideBarSignHeader3;
        private System.Windows.Forms.TextBox tbSideBarSignHeader2;
        private System.Windows.Forms.TextBox tbSideBarSignHeader1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbApprovedBy;
        private System.Windows.Forms.TextBox tbNControlBy;
        private System.Windows.Forms.TextBox tbDesignedBy;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NumericUpDown nudFirstPageNumber;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox cbAddChangelogPage;
        private System.Windows.Forms.CheckBox cbFullNames;
        private System.Windows.Forms.Button makeReportButton;
        private System.Windows.Forms.NumericUpDown nudCountAfter;
        private System.Windows.Forms.NumericUpDown nudCountBefore;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TabControl tabReport;
        private System.Windows.Forms.TabPage tabPagesAddNewPages;
        private System.Windows.Forms.TabPage tabPagesAttributes;


        private void chBoxDocumentation_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDocumentation.Checked)
                nudDocumentation.Enabled = true;
            else nudDocumentation.Enabled = false;
        }

        private void chBoxAssembly_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxAssembly.Checked)
                nudAssembly.Enabled = true;
            else nudAssembly.Enabled = false;
        }

        private void chBoxDetails_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxDetails.Checked)
                nudDetails.Enabled = true;
            else nudDetails.Enabled = false;
        }

        private void chBoxStdProducts_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxStdProducts.Checked)
                nudStdProducts.Enabled = true;
            else nudStdProducts.Enabled = false;
        }

        private void chBoxOtherProducts_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxOtherProducts.Checked)
                nudStdProducts.Enabled = true;
            else nudStdProducts.Enabled = false;
        }

        private void chBoxMaterials_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxMaterials.Checked)
                nudMaterials.Enabled = true;
            else nudMaterials.Enabled = false;
        }

        private void chBoxKomplekt_CheckedChanged(object sender, EventArgs e)
        {
            if (chBoxKomplekt.Checked)
                nudKomplekt.Enabled = true;
            else nudKomplekt.Enabled = false;
        }

        private bool _attributesFormClose = false;
        private bool _closeGenerator = false;

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            _attributesFormClose = true;
        }

        void UpdateUI()
        {
            Update();
            Application.DoEvents();
        }
        
        public bool AttributesFormClose
        {
            get { return _attributesFormClose; }
        }

        public bool CloseGenerator
        {
            get { return _closeGenerator; }
        }

        Guid usersGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        Guid chiefDesignerGuid = new Guid("60934222-B8CD-475E-83C9-E8CF9E5C2894");
        Guid userShortNameGuid = new Guid("00e19b83-9889-40b1-bef9-f4098dd9d783");


        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();



            SqlConnection connection = ApiDocs.GetConnection(true);

            // Справочник "Пользователи"
            var UserReferenceInfo = ReferenceCatalog.FindReference(usersGuid);
            var UserReference = UserReferenceInfo.CreateReference();

            // Получение количества исполнений 
            string countMakeQuery = String.Format(@"SELECT Count(MasterID) 
                                                    FROM NomenclatureBaseVersion
                                                    WHERE SlaveID = {0}",
                                                  context.ObjectsInfo[0].ObjectId);

            SqlCommand countMakeCommand = new SqlCommand(countMakeQuery, connection);
            countMakeCommand.CommandTimeout = 0;

            int countMake = Convert.ToInt32(countMakeCommand.ExecuteScalar());

            if (countMake == 0)
            {
                GBNewPages.SetBounds(150, 25, 277, 268);
                GBNewPages.Text = "Добавление резервных страниц";
                groupBoxNewPagesVarPart.Enabled = false;
                groupBoxNewPagesVarPart.Visible = false;
                groupBox5.Enabled = false;
                groupBox5.Visible = false;
                groupBox4.SetBounds(150, 304, 277, 54);

            }

            // Получение должности пользователя
            ReferenceObject post = UserReference.Find(chiefDesignerGuid);
            // Получение краткого имени пользователя через должность
            ReferenceObject user = post.Children.FirstOrDefault();
            string userName = user.ParameterValues[userShortNameGuid].Value.ToString();
            tBoxGoev.Text = userName.Substring(userName.Length - 4, 4) + " " + userName.Substring(0, userName.Length - 4);
            //Запрос на выборку начальника бригады
            string chiefSearchQuery = String.Format
              (@"SELECT up2.LastName
                 FROM ((UsersHierarchy uh INNER JOIN Users u ON (uh.s_ParentID = u.s_ObjectID)
                 INNER JOIN Users u2 ON (uh.s_ObjectID = u2.s_ObjectID)
                 INNER JOIN UserParameters up ON (up.s_ObjectID = u2.s_ObjectID)) INNER JOIN    
                 (UsersHierarchy uh2 INNER JOIN Users u3 ON (uh2.s_ParentID = u3.s_ObjectID)
                 INNER JOIN Users u4 ON (uh2.s_ObjectID = u4.s_ObjectID)) 
                 ON (u4.s_ObjectID = uh.s_ParentID)) INNER JOIN
                 ((UsersHierarchy uh3 INNER JOIN Users u5 ON (uh3.s_ParentID = u5.s_ObjectID)
                 INNER JOIN Users u6 ON (uh3.s_ObjectID = u6.s_ObjectID)) INNER JOIN
                 (UsersHierarchy uh4 INNER JOIN Users u7 ON (uh4.s_ParentID = u7.s_ObjectID)
                 INNER JOIN Users u8 ON (uh4.s_ObjectID = u8.s_ObjectID)
                 INNER JOIN UserParameters up2 ON (up2.s_ObjectID = u8.s_ObjectID))
                 ON(u6.s_ObjectID = uh4.s_ParentID))
                 ON (u3.s_ObjectID = uh3.s_ParentID)
                 WHERE up.LastName = N'{0}' AND u6.FullName = N'Начальник'
                 UNION ALL
                 SELECT ''", context.GetAutorLastName());

            SqlCommand chiefSearchCommand = new SqlCommand(chiefSearchQuery, connection);
            chiefSearchCommand.CommandTimeout = 0;


            tBoxCheifTeam.Text = chiefSearchCommand.ExecuteScalar().ToString();
            tbApprovedBy.Text = tBoxCheifTeam.Text;

            //Запрос на выборку нормоконтроллёра 
            string allConstrQuery = String.Format
              (@"SELECT up2.LastName
                 FROM ((UsersHierarchy uh INNER JOIN Users u ON (uh.s_ParentID = u.s_ObjectID)
                 INNER JOIN Users u2 ON (uh.s_ObjectID = u2.s_ObjectID)
                 INNER JOIN UserParameters up ON (up.s_ObjectID = u2.s_ObjectID)) INNER JOIN    
                 (UsersHierarchy uh2 INNER JOIN Users u3 ON (uh2.s_ParentID = u3.s_ObjectID)
                 INNER JOIN Users u4 ON (uh2.s_ObjectID = u4.s_ObjectID)) 
                 ON (u4.s_ObjectID = uh.s_ParentID)) INNER JOIN
                 ((UsersHierarchy uh3 INNER JOIN Users u5 ON (uh3.s_ParentID = u5.s_ObjectID)
                 INNER JOIN Users u6 ON (uh3.s_ObjectID = u6.s_ObjectID)) INNER JOIN
                 (UsersHierarchy uh4 INNER JOIN Users u7 ON (uh4.s_ParentID = u7.s_ObjectID)
                 INNER JOIN Users u8 ON (uh4.s_ObjectID = u8.s_ObjectID)
                 INNER JOIN UserParameters up2 ON (up2.s_ObjectID = u8.s_ObjectID))
                 ON(u6.s_ObjectID = uh4.s_ParentID))
                 ON (u3.s_ObjectID = uh3.s_ParentID)
                 WHERE up.LastName = N'{0}' AND u6.FullName = N'Конструктор'", context.GetAutorLastName());

            SqlCommand allConstrCommand = new SqlCommand(allConstrQuery, connection);
            allConstrCommand.CommandTimeout = 0;
            SqlDataReader allConstrReader = allConstrCommand.ExecuteReader();


            if (allConstrReader.HasRows)

                while (allConstrReader.Read())
                {
                    cBoxNorm.Items.Add(allConstrReader.GetSqlString(0)); // combobox нормоконтроллёр

                    cbSideBarSignName3.Items.Add(allConstrReader.GetSqlString(0)); // comboBox нормоконтроллёр
                    cbControlledBy.Items.Add(allConstrReader.GetSqlString(0));  // comboBox Проверил
                }
            allConstrReader.Close();




            tBoxYearMake.Text = DateTime.Now.Year.ToString();

            //Запрос на выборку технологов
            string technologSearchQuery = @"SELECT up.LastName
                                            FROM UsersHierarchy uh INNER JOIN Users u ON (uh.s_ParentID = u.s_ObjectID)
                                                    INNER JOIN Users u2 ON (uh.s_ObjectID = u2.s_ObjectID)
                                                    INNER JOIN UserParameters up ON (up.s_ObjectID = u2.s_ObjectID)
                                            WHERE u.FullName = N'Технолог'
                                            ORDER BY up.LastName";
            SqlCommand technologSearchCommand = new SqlCommand(technologSearchQuery, connection);
            technologSearchCommand.CommandTimeout = 0;
            SqlDataReader technologSearchReader = technologSearchCommand.ExecuteReader();


            if (technologSearchReader.HasRows)

                while (technologSearchReader.Read())
                {
                    cBoxTech.Items.Add(technologSearchReader.GetSqlString(0)); // comboBox технолог 
                    cbSideBarSignName1.Items.Add(technologSearchReader.GetSqlString(0)); // textBox разработчик
                }
            technologSearchReader.Close();


            connection.Close();

         
                TFDDocument document = new TFDDocument(context.ObjectsInfo[0].ObjectId);
                string nameRow = document.Naimenovanie;
     
         
            textBoxFullName.Text = nameRow.ToLower();

            tbDesignedBy.Text = context.GetAutorLastName();
            tbSideBarSignName2.Text = context.GetAutorLastName();

            while (AttributesFormClose == false)
            {
                this.UpdateUI();
            }

            if (CloseGenerator)
            {
                _closeGenerator = false;
                this.Close();
                return;
            }

            DocumentAttributes documentAttr = new DocumentAttributes();
            documentAttr.SidebarSignHeader1 = (tbSideBarSignHeader1.Text == null) ? "" : tbSideBarSignHeader1.Text;
            documentAttr.SidebarSignHeader2 = (tbSideBarSignHeader2.Text == null) ? "" : tbSideBarSignHeader2.Text;
            documentAttr.SidebarSignHeader3 = (tbSideBarSignHeader3.Text == null) ? "" : tbSideBarSignHeader3.Text;

            documentAttr.SidebarSignName1 = (tbSideBarSignName2.Text == null) ? "" : tbSideBarSignName2.Text;
            documentAttr.SidebarSignName2 = (cbSideBarSignName1.Text == null) ? "" : cbSideBarSignName1.Text;
            documentAttr.SidebarSignName3 = (cbSideBarSignName3.Text == null) ? "" : cbSideBarSignName3.Text;

            if (documentAttr.SidebarSignName1.Trim() == string.Empty)
            {
                documentAttr.SidebarSignHeader1 = string.Empty;
            }

            if (documentAttr.SidebarSignName2.Trim() == string.Empty)
            {
                documentAttr.SidebarSignHeader2 = string.Empty;
            }

            if (documentAttr.SidebarSignName3.Trim() == string.Empty)
            {
                documentAttr.SidebarSignHeader3 = string.Empty;
            }

            documentAttr.DesignedBy = (tbDesignedBy.Text == null) ? "" : tbDesignedBy.Text;
            documentAttr.ControlledBy = (cbControlledBy.Text == null) ? "" : cbControlledBy.Text;
            documentAttr.ApprovedBy = (tbApprovedBy.Text == null) ? "" : tbApprovedBy.Text;
            documentAttr.NControlBy = (tbNControlBy.Text == null) ? "" : tbNControlBy.Text;

            documentAttr.FirstPageNumber = Convert.ToInt32(nudFirstPageNumber.Value);

            if (cbAddChangelogPage.CheckState == CheckState.Checked)
                documentAttr.AddChangelogPage = true;
            else documentAttr.AddChangelogPage = false;

            if (cbFullNames.CheckState == CheckState.Checked)
            {
                documentAttr.FullNames = true;
                documentAttr.FullName = textBoxFullName.Text;
            }
            else documentAttr.FullNames = false;

            if (cbZeroToDash.CheckState == CheckState.Checked)
                documentAttr.ZeroToDash = true;
            else documentAttr.ZeroToDash = false;

            if (chBoxAddMake.CheckState == CheckState.Checked)
                documentAttr.AddMake = true;
            else documentAttr.AddMake = false;

            if (chBoxProgramDocToLower.CheckState == CheckState.Checked)
                documentAttr.programToLower = true;
            else documentAttr.programToLower = false;

            if (chBoxSortByAlphabet.CheckState == CheckState.Checked)
                documentAttr.SortByAlphabet = true;
            else documentAttr.SortByAlphabet = false;

            if (chBoxAddModify.CheckState == CheckState.Checked)
                documentAttr.AddModify = true;
            else documentAttr.AddModify = false;

            if (checkBoxSectionProgramSoft.CheckState == CheckState.Checked)
                documentAttr.addSectionProgramSoft = true;
            else documentAttr.addSectionProgramSoft = false;

            if (checkBoxSpecForPI.CheckState == CheckState.Checked)
                documentAttr.specificationForPI = true;
            else documentAttr.specificationForPI = false;

            documentAttr.CountEmptyLinesBefore = Convert.ToInt32(nudCountBefore.Value);

            documentAttr.CountEmptyLinesAfter = Convert.ToInt32(nudCountAfter.Value);

            documentAttr.DocumentationPages = nudDocumentation.Value;
            documentAttr.AssemblyPages = nudAssembly.Value;
            documentAttr.DetailsPages = nudDetails.Value;
            documentAttr.StdProductsPages = nudStdProducts.Value;
            documentAttr.OtherProductsPages = nudOtherProducts.Value;
            documentAttr.MaterialsPages = nudMaterials.Value;
            documentAttr.KomplektPages = nudKomplekt.Value;


            documentAttr.DocumentationPagesVarPart = nudDocumentVarPart.Value;
            documentAttr.AssemblyPagesVarPart = nudAssemblyVarPart.Value;
            documentAttr.DetailsPagesVarPart = nudDetailVarPart.Value;
            documentAttr.StdProductsPagesVarPart = nudStdItemsVarPart.Value;
            documentAttr.OtherProductsPagesVarPart = nudOtherItemsVarPart.Value;
            documentAttr.MaterialsPagesVarPart = nudMaterialVarPart.Value;
            documentAttr.KomplektPagesVarPart = nudKomplektVarPart.Value;
            documentAttr.ThroughNumbering = chBoxThroughNumbering.Checked;

            if (chBoxDocVarPart.CheckState == CheckState.Checked)
                documentAttr.docVarPart = true;
            else documentAttr.docVarPart = false;

            if (chBoxAssemblyVarPart.CheckState == CheckState.Checked)
                documentAttr.assemVarPart = true;
            else documentAttr.assemVarPart = false;

            if (chBoxDetailVarPart.CheckState == CheckState.Checked)
                documentAttr.detVarPart = true;
            else documentAttr.detVarPart = false;

            if (chBoxStdItemVarPart.CheckState == CheckState.Checked)
                documentAttr.stdVarPart = true;
            else documentAttr.stdVarPart = false;

            if (chBoxOtherItemVarPart.CheckState == CheckState.Checked)
                documentAttr.othVarPart = true;
            else documentAttr.othVarPart = false;

            if (chBoxMaterialVarPart.CheckState == CheckState.Checked)
                documentAttr.matVarPart = true;
            else documentAttr.matVarPart = false;

            if (chBoxKomplectVarPart.CheckState == CheckState.Checked)
                documentAttr.komVarPart = true;
            else documentAttr.komVarPart = false;

            if (chBoxNewPageOtherItemsAfterProductsKind.CheckState == CheckState.Checked)
                documentAttr.newPageOIAfterKindProducts = true;
            else documentAttr.newPageOIAfterKindProducts = false;

            if (chBoxDocumentation.CheckState == CheckState.Checked)
                documentAttr.doc = true;
            else documentAttr.doc = false;

            if (chBoxAssembly.CheckState == CheckState.Checked)
                documentAttr.assem = true;
            else documentAttr.assem = false;

            if (chBoxDetails.CheckState == CheckState.Checked)
                documentAttr.det = true;
            else documentAttr.det = false;

            if (chBoxStdProducts.CheckState == CheckState.Checked)
                documentAttr.std = true;
            else documentAttr.std = false;

            if (chBoxOtherProducts.CheckState == CheckState.Checked)
                documentAttr.oth = true;
            else documentAttr.oth = false;

            if (chBoxMaterials.CheckState == CheckState.Checked)
                documentAttr.mat = true;
            else documentAttr.mat = false;

            if (chBoxKomplekt.CheckState == CheckState.Checked)
                documentAttr.kom = true;
            else documentAttr.kom = false;

            if (chBoxNewPageOtherItemsAfterProductsKindVaiosPart.CheckState == CheckState.Checked)
                documentAttr.newPageOIAfterKindProductsVariosPart = true;
            else documentAttr.newPageOIAfterKindProductsVariosPart = false;


            if (chBoxAddTitul.CheckState == CheckState.Checked)
            {
                documentAttr.Goev = tBoxGoev.Text;
                documentAttr.MilitaryChief = tBoxMilitaryChief.Text;
                documentAttr.Tech = cBoxTech.Text;
                documentAttr.Otdel45 = tBox45.Text;
                documentAttr.Norm = cBoxNorm.Text;
                documentAttr.VPMO = tBoxVPMO.Text;
                documentAttr.ChiefTeam = tBoxCheifTeam.Text;
                documentAttr.YearMake = tBoxYearMake.Text;
                documentAttr.AddTitulList = true;

                if (chBoxChiefMillitary.CheckState == CheckState.Checked)
                    documentAttr.checkChiefMillitary = true;
                else documentAttr.checkChiefMillitary = false;

                if (chBoxVPMO.CheckState == CheckState.Checked)
                    documentAttr.checkVPMO = true;
                else documentAttr.checkVPMO = false;

            }
            else documentAttr.AddTitulList = false;


            if (chBoxAddModify.CheckState == CheckState.Checked)
            {
                documentAttr.NumberModify = textBoxNumberModify.Text;
                documentAttr.NumberDocument = textBoxNumberDoc.Text;
                documentAttr.DateModify = textBoxDate.Text;
                documentAttr.ListZam = textBoxList.Text;
                documentAttr.ListNov = textBoxList2.Text;

                int index = 0;
                int indexDash = 0;
                string str = textBoxListPage.Text.Trim();
                string strWithDash = string.Empty;
                int firstPage = 0;
                int secondPage = 0;

                // Преобразование строки с номерами страниц через запятую в список номеров страниц
                while (str != string.Empty)
                {
                    index = str.IndexOf(",");

                    if (index != -1)
                    {
                        strWithDash = str.Substring(0, index).Trim();
                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {

                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = str.Remove(0, str.Substring(0, index - indexDash).Length);

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage.Add(i);

                        }
                        else
                        {

                            documentAttr.ListPage.Add(Convert.ToInt32(str.Substring(0, index).Trim()));
                            str = str.Remove(0, str.Substring(0, index + 1).Length);
                        }
                    }
                    else
                    {

                        index = str.Length;
                        strWithDash = str.Trim();



                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {
                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = string.Empty;

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage.Add(i);

                        }
                        else
                        {
                            documentAttr.ListPage.Add(Convert.ToInt32(str.Trim()));
                            str = string.Empty;
                        }
                    }

                }

                index = 0;
                indexDash = 0;
                str = textBoxListPage2.Text.Trim();
                strWithDash = string.Empty;
                firstPage = 0;
                secondPage = 0;


                // Преобразование строки с номерами страниц через запятую в список номеров страниц listPage2
                while (str != string.Empty)
                {
                    index = str.IndexOf(",");

                    if (index != -1)
                    {
                        strWithDash = str.Substring(0, index).Trim();
                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {

                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = str.Remove(0, str.Substring(0, index - indexDash).Length);

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage2.Add(i);

                        }
                        else
                        {

                            documentAttr.ListPage2.Add(Convert.ToInt32(str.Substring(0, index).Trim()));
                            str = str.Remove(0, str.Substring(0, index + 1).Length);
                        }
                    }
                    else
                    {

                        index = str.Length;
                        strWithDash = str.Trim();



                        indexDash = strWithDash.IndexOf("-");
                        if (indexDash != -1)
                        {
                            firstPage = Convert.ToInt32(strWithDash.Substring(0, indexDash).Trim());
                            str = str.Remove(0, str.Substring(0, indexDash + 1).Length);
                            secondPage = Convert.ToInt32(str.Substring(0, index - (indexDash + 1)).Trim());
                            str = string.Empty;

                            for (int i = firstPage; i <= secondPage; i++)
                                documentAttr.ListPage2.Add(i);

                        }
                        else
                        {
                            documentAttr.ListPage2.Add(Convert.ToInt32(str.Trim()));
                            str = string.Empty;
                        }
                    }

                }


            }

            this.Close();

            // ТОЧКА ВХОДА МАКРОСА
            Globus.TFlexDocs.SpecificationReport.GOST_2_113.CadReportVariantA.Make(context, documentAttr);
        }

        private TabPage tabPageTitul;
        private TextBox tBoxYearMake;
        private Label label19;
        private GroupBox grBoxVisa;
        private TextBox tBoxCheifTeam;
        private Label label18;
        private TextBox tBoxVPMO;
        private Label label16;
        private Label label17;
        private TextBox tBox45;
        private Label label14;
        private Label label15;
        private CheckBox chBoxAddTitul;
        private GroupBox grBoxApprSign;
        private TextBox tBoxMilitaryChief;
        private TextBox tBoxGoev;
        private Label label13;
        private Label label12;
        private CheckBox cbZeroToDash;
        private CheckBox chBoxNewPageOtherItemsAfterProductsKind;
        private CheckBox chBoxVPMO;
        private CheckBox chBoxChiefMillitary;
        private CheckBox chBoxAddMake;
        private TabPage tabPageAddModify;
        private CheckBox chBoxAddModify;
        private GroupBox grBoxModifySpec;
        private Label label24;
        private TextBox textBoxList;
        private Label label23;
        private Label label22;
        private Label label21;
        private Label label20;
        private TextBox textBoxNumberModify;
        private TextBox textBoxListPage;
        private TextBox textBoxDate;
        private TextBox textBoxNumberDoc;
        private ComboBox cBoxTech;
        private Label label25;
        private TextBox textBoxList2;
        private Label label26;
        private TextBox textBoxListPage2;
        private ComboBox cbSideBarSignName3;
        private ComboBox cbSideBarSignName1;
        private ComboBox cbControlledBy;
        private ComboBox cBoxNorm;
        private GroupBox groupBoxNewPagesVarPart;
        private NumericUpDown nudKomplektVarPart;
        private NumericUpDown nudMaterialVarPart;
        private NumericUpDown nudOtherItemsVarPart;
        private NumericUpDown nudStdItemsVarPart;
        private NumericUpDown nudDetailVarPart;
        private NumericUpDown nudAssemblyVarPart;
        private NumericUpDown nudDocumentVarPart;
        private CheckBox chBoxKomplectVarPart;
        private CheckBox chBoxMaterialVarPart;
        private CheckBox chBoxOtherItemVarPart;
        private CheckBox chBoxStdItemVarPart;
        private CheckBox chBoxDetailVarPart;
        private CheckBox chBoxAssemblyVarPart;
        private CheckBox chBoxDocVarPart;
        private GroupBox GBNewPages;
        private NumericUpDown nudKomplekt;
        private NumericUpDown nudMaterials;
        private NumericUpDown nudOtherProducts;
        private NumericUpDown nudStdProducts;
        private NumericUpDown nudDetails;
        private NumericUpDown nudAssembly;
        private NumericUpDown nudDocumentation;
        private CheckBox chBoxKomplekt;
        private CheckBox chBoxMaterials;
        private CheckBox chBoxOtherProducts;
        private CheckBox chBoxStdProducts;
        private CheckBox chBoxDetails;
        private CheckBox chBoxAssembly;
        private CheckBox chBoxDocumentation;
        private CheckBox chBoxThroughNumbering;
        private GroupBox groupBox5;
        private CheckBox chBoxNewPageOtherItemsAfterProductsKindVaiosPart;
        private GroupBox groupBox4;
        private TextBox textBoxFullName;
        private CheckBox checkBoxSectionProgramSoft;
        private CheckBox chBoxSortByAlphabet;
        private CheckBox checkBoxSpecForPI;
        private CheckBox chBoxProgramDocToLower;
    }
}