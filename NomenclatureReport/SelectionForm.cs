using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

namespace NomenclatureReport
{
    public partial class SelectionForm : Form
    {

        NomenclatureReport report = new NomenclatureReport();
        string reportFilePathInForm;
        List<TFDDocument> listDocument;
        Dictionary<int, int> listProductId = new Dictionary<int, int>();
        ReportParameters reportParameters;

        public SelectionForm()
        {
            InitializeComponent();
            // Инициализируем DataGridView-ы
            report.InitializeControls(dgvProducts, dgvAddedProducts);
            // заполнение DataGridView и выделение первого комплекса
            string classTostring = string.Empty;
            if (checkBoxComplement.CheckState == CheckState.Checked && checkBoxComplex.CheckState == CheckState.Checked)
                classTostring = "501,502";
            else
            {
                if (checkBoxComplex.CheckState == CheckState.Checked)
                    classTostring = "502";
                else
                {
                    if (checkBoxComplement.CheckState == CheckState.Checked)
                        classTostring = "501";
                    else
                    {
                        classTostring = "null";
                    }
                }
            }

            report.FillProducts(PurchaseDesignProductOperation.GetProducts(classTostring), dgvProducts);
            if (dgvProducts.Rows.Count != 0) dgvProducts.Rows[0].Selected = true;
        }

        private void chkAllTypes_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllTypes.CheckState == CheckState.Checked)
            {
                chkDetail.Checked = true;
                chkAssembly.Checked = true; 
                chkStandart.Checked = true;
                chkOther.Checked = true;
                chkMaterial.Checked = true;
                chkComplement.Checked = true;
                chkCode.Checked = true;
                chkCodeless.Checked = true;
            }
        }

        public PurchaseDesignProductOperation productOperation = new PurchaseDesignProductOperation();

    
        private void MakeReport()
        {
            addProductButton.Enabled = false;
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            btnMakeReport.Enabled = false;
            progressLabel.Text = "Формирование отчета";

            double step = progressBar1.Maximum / (4 * listProductId.Count);
            int step1 = Convert.ToInt32(Math.Round(step));

            var prodIdAndCount = new Dictionary<int, int>();
            foreach (var productId in listProductId)
            {
                listDocument.AddRange(productOperation.BuildTree(productId.Key,productId.Value, reportParameters));
                progressBar1.Value += step1;
            }

            report.GetExcelReport(listDocument, reportFilePathInForm, reportParameters, productOperation, progressBar1);

            addProductButton.Enabled = true;
            deleteComplexButton.Enabled = true;
            clearDBButton.Enabled = true;
            btnMakeReport.Enabled = true;

            progressBar1.Value = progressBar1.Maximum;

            this.Close();
        }

        private void btnMakeReport_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvAddedProducts.Rows.Count; i++)
                listProductId.Add(Convert.ToInt32(dgvAddedProducts.Rows.SharedRow(i).Cells[0].Value), Convert.ToInt32(dgvAddedProducts.Rows.SharedRow(i).Cells[4].Value));
            

            reportParameters.Detail = (chkDetail.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Assembly = ((chkAssembly.CheckState == CheckState.Checked) || (chkAssembly.CheckState == CheckState.Indeterminate)) ? true : false;
            reportParameters.AssemblyCode = (chkCode.CheckState == CheckState.Checked) ? true : false;
            reportParameters.AssemblyCodeless = (chkCodeless.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Standart = (chkStandart.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Other = (chkOther.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Materials = (chkMaterial.CheckState == CheckState.Checked) ? true : false;
            reportParameters.Complement = (chkComplement.CheckState == CheckState.Checked) ? true : false;

            if (!reportParameters.Assembly && !reportParameters.Detail && !reportParameters.Standart && !reportParameters.Other && !reportParameters.Materials && !reportParameters.Complement)
                MessageBox.Show("Ни один тип не выбран. Выберите один из предложенных типов!", "Внимание!", MessageBoxButtons.OK);
            else
                MakeReport();
        }



        public void DotsEntry(IReportGenerationContext context)
        {
            this.Show();
            reportFilePathInForm = context.ReportFilePath;
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void addProductButton_Click(object sender, EventArgs e)
        {
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            btnMakeReport.Enabled = false;
            addProductButton.Enabled = false;
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            btnMakeReport.Enabled = false;

            // получение DocId выделенного комплекса
            progressBar1.Value = 25;

            progressBar1.Value = 50;

            List<TFDDocument> totalTree = new List<TFDDocument>();

            listDocument = new List<TFDDocument>();

            reportParameters = new ReportParameters();
            progressBar1.Value = 75;
            report.AddProduct(dgvAddedProducts,(Convert.ToInt32(dgvProducts.SelectedRows[0].Cells[0].Value)));

            progressBar1.Value = 0;

            deleteComplexButton.Enabled = true;
            clearDBButton.Enabled = true;
            btnMakeReport.Enabled = true;
            addProductButton.Enabled = true;
        }

        private void chkCodeless_CheckedChanged(object sender, EventArgs e)
        {
            chkAssembly.ThreeState = true;
            if (chkCodeless.CheckState == CheckState.Checked)
            {
                if (chkCode.CheckState == CheckState.Checked)
                {
                    chkAssembly.ThreeState = false;
                    chkAssembly.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAssembly.ThreeState = true;
                    chkAssembly.CheckState = CheckState.Indeterminate;
                }
            }
            else
            {
                if (chkCode.CheckState == CheckState.Checked)
                {
                    chkAssembly.ThreeState = true;
                    chkAssembly.CheckState = CheckState.Indeterminate;
                }
                else
                {
                    chkAssembly.ThreeState = false;
                    chkAssembly.CheckState = CheckState.Unchecked;
                }
            }
        }

        private void chkAssembly_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAssembly.CheckState == CheckState.Checked)
            {
                chkCode.CheckState = CheckState.Checked;
                chkCodeless.CheckState = CheckState.Checked;
                chkCodeless.Enabled = true;
                chkCode.Enabled = true;
            }
            else
            {
                chkCode.CheckState = CheckState.Unchecked;
                chkCodeless.CheckState = CheckState.Unchecked;
                chkCodeless.Enabled = false;
                chkCode.Enabled = false;
            }
        }

        private void chkCode_CheckedChanged(object sender, EventArgs e)
        {
            chkAssembly.ThreeState = true;
            if (chkCode.CheckState == CheckState.Checked)
            {
                if (chkCodeless.CheckState == CheckState.Checked)
                {
                    chkAssembly.ThreeState = false ;
                    chkAssembly.CheckState = CheckState.Checked;
                }
                else
                {
                    chkAssembly.ThreeState = true;
                    chkAssembly.CheckState = CheckState.Indeterminate;
                }
            }
            else
            {
                if (chkCodeless.CheckState == CheckState.Checked)
                {
                    chkAssembly.ThreeState = true;
                    chkAssembly.CheckState = CheckState.Indeterminate;
                }
                else
                {
                    chkAssembly.ThreeState = false;
                    chkAssembly.CheckState = CheckState.Unchecked;
                }
            }
        }

        private void deleteComplexButton_Click(object sender, EventArgs e)
        {
            if (dgvAddedProducts.SelectedRows.Count > 0)
                dgvAddedProducts.Rows.RemoveAt(dgvAddedProducts.SelectedRows[0].Index);  

            if (dgvAddedProducts.RowCount == 0)
            {
                clearDBButton.Enabled = false;
                deleteComplexButton.Enabled = false;
                btnMakeReport.Enabled = false;
            }
        }

        private void clearDBButton_Click(object sender, EventArgs e)
        {     
            dgvAddedProducts.Rows.Clear();
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
        }

        private void checkBoxComplex_CheckedChanged(object sender, EventArgs e)
        {
            dgvProducts.Rows.Clear();

            string classTostring = string.Empty;
            if (checkBoxComplement.CheckState == CheckState.Checked && checkBoxComplex.CheckState == CheckState.Checked)
                classTostring = "501,502";
            else
            {
                if (checkBoxComplex.CheckState == CheckState.Checked)
                    classTostring = "502";
                else
                {
                    if (checkBoxComplement.CheckState == CheckState.Checked)
                        classTostring = "501";
                    else
                        classTostring = "null";
                }
            }
            report.FillProducts(PurchaseDesignProductOperation.GetProducts(classTostring), dgvProducts);
            dgvProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void checkBoxComplement_CheckedChanged(object sender, EventArgs e)
        {
            dgvProducts.Rows.Clear();

            string classTostring = string.Empty;
           
            if (checkBoxComplement.CheckState == CheckState.Checked && checkBoxComplex.CheckState == CheckState.Checked)
                classTostring = "501,502";
            else
            {
                if (checkBoxComplex.CheckState == CheckState.Checked)
                    classTostring = "502";
                else
                {
                    if (checkBoxComplement.CheckState == CheckState.Checked)
                        classTostring = "501";
                    else
                        classTostring = "null";
                }
            }
            report.FillProducts(PurchaseDesignProductOperation.GetProducts(classTostring), dgvProducts);
            dgvProducts.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }
    }
}
