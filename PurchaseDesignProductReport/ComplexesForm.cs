using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

using PurchaseDesignProductReport;

namespace PurchaseDesignProductReportForm
{
    public partial class ComplexesForm : Form
    {
        PurchaseProductReport report = new PurchaseProductReport();
        bool isPKIReport;
        string reportFilePathInForm;  

        public ComplexesForm()
        {
            InitializeComponent();

            // Инициализируем DataGridView-ы
            report.InitializeControls(dgvProducts, dgvAddedProducts);
        }    

        private void addProductButton_Click(object sender, EventArgs e)
        {
            addProductButton.Enabled = false;
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            reportButton.Enabled = false;

            // получение DocId выделенного комплекса
            progressBar1.Value = 1;

            // выбор типа отчета
            if (rbPurchaseProducts.Checked) isPKIReport = true;
            else isPKIReport = false;

            reportPanel.Enabled = false; // запрет изменения типа отчета

            // добавление заказа (построение дерева объектов и входимостей)
            progressBar1.Value = 2;

            int objectId = 0;
            objectId = Convert.ToInt32(dgvProducts.SelectedRows[0].Cells[0].Value);

            PurchaseDesignProductOperation.BuildTree(objectId, isPKIReport);

            // расчет количеств для отчета
            progressBar1.Value = 3;

            PurchaseDesignProductOperation.Calculate(objectId);
            progressBar1.Value = 4;

            // считываем данные последнего добавленного заказа (в списке он первый, поскольку сортировка заказов по убыванию дат) в DataGridView
            report.AddProduct(PurchaseDesignProductOperation.GetAddedProducts()[0], dgvAddedProducts);

            if (dgvAddedProducts.Rows.Count != 0) dgvAddedProducts.Rows[0].Selected = true;

            progressBar1.Value = 0;

            deleteComplexButton.Enabled = true;
            clearDBButton.Enabled = true;
            reportButton.Enabled = true;
            addProductButton.Enabled = true;
        }

        private void clearDBButton_Click(object sender, EventArgs e)
        {
            DialogResult attention = MessageBox.Show("Вы уверены, что хотите удалить все обсчитанные заказы?", "Внимание!", MessageBoxButtons.YesNo);
            if (attention == DialogResult.Yes)
            {
                PurchaseDesignProductOperation.DeleteAllAddedProducts();
                dgvAddedProducts.Rows.Clear();
            }
        }     

        private void reportButton_Click(object sender, EventArgs e)
        {
            addProductButton.Enabled = false;
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            reportButton.Enabled = false;

            List<int> productIds = new List<int>();

            for (int row = 0; row < dgvAddedProducts.Rows.Count; row++)
            {
                // productId
                if ((bool)dgvAddedProducts.Rows[row].Cells[0].Value == true)
                {
                    productIds.Add((int)dgvAddedProducts.Rows[row].Cells[1].Value);
                }
            }

            progressBar1.Value = 2;

            report.GetExcelReport(productIds, isPKIReport, reportFilePathInForm);

            addProductButton.Enabled = true;
            deleteComplexButton.Enabled = true;
            clearDBButton.Enabled = true;
            reportButton.Enabled = true;

            progressBar1.Value = 4;
//            System.Diagnostics.Process.Start(reportFilePathInForm);
            _attributesFormClose = true;
            this.Close();
        }

        private void ComplexesForm_Load(object sender, EventArgs e)
        {
            // очистка БД отчета
            PurchaseDesignProductOperation.DeleteAllAddedProducts();

            // заполнение DataGridView и выделение первого комплекса
            report.FillProducts(PurchaseDesignProductOperation.GetProducts(), dgvProducts);
            if (dgvProducts.Rows.Count != 0) dgvProducts.Rows[0].Selected = true;

            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            reportButton.Enabled = false;
        }

        private void dgvProducts_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                dgvProducts.Rows[e.RowIndex].Selected = true;
        }

        private void dgvAddedProducts_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                dgvAddedProducts.Rows[e.RowIndex].Selected = true;
        }

        private void deleteComplexButton_Click(object sender, EventArgs e)
        {
            if (dgvAddedProducts.Rows.Count != 0)
            {
                PurchaseDesignProductOperation.DeleteProductStructure((int)dgvAddedProducts.SelectedRows[0].Cells[1].Value);
                dgvAddedProducts.Rows.RemoveAt(dgvAddedProducts.SelectedRows[0].Index);
            }
        }
    }
}
