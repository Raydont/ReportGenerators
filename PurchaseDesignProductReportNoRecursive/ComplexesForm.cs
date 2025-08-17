using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using PurchaseDesignProductReportNoRecursive;

namespace PurchaseDesignProductReportForm
{
    public partial class ComplexesForm : Form
    {
        PurchaseProductReport report = new PurchaseProductReport();
        bool isPKIReport;
        bool isFirstUnitCodeDevices;
        bool isOccurence;
        string reportFilePathInForm;

        public string userFullName = string.Empty;

        public ComplexesForm()
        {
            InitializeComponent();

            // Инициализируем DataGridView-ы
            report.InitializeControls(dgvProducts, dgvAddedProducts);
        }

        private void addProductButton_Click(object sender, EventArgs e)
        {
            int objectId = Convert.ToInt32(dgvProducts.SelectedRows[0].Cells[0].Value);
            addOrder(objectId, 1);

          //  return;
          //  var listObg = new Dictionary<int, int>();
          //  listObg.Add(129700, 10);   //БФМИ.461271.026
          //  //listObg.Add(120133, 10);   //БФМИ.461271.026
          //  //listObg.Add(68536, 6);     //БФМИ.461272.035-06
          //  //listObg.Add(326934, 1);    //БФМИ.461271.128
          //  //listObg.Add(121344, 8);    //БФМИ.461272.046
          //  //listObg.Add(97301, 1);     //БФМИ.461271.057-01
          //  //listObg.Add(352165, 1);    //БФМИ.468929.055
          //  //listObg.Add(344979, 1);    //БФМИ.461271.133
          //  //listObg.Add(346529, 1);    //БФМИ.461271.132
          ////   listObg.Add(351609, 1);    //БФМИ.468929.226
          ////  listObg.Add(340262, 1);    //БФМИ.461271.130
          ////   listObg.Add(350325, 1);    //БФМИ.468929.216
          //  //listObg.Add(75699, 3);     //БФМИ.461271.082
          //  //listObg.Add(120898, 3);    //БФМИ.461271.090
          //  //listObg.Add(334722, 1);    //БФМИ.461271.056-06.01
          //  //listObg.Add(334893, 1);    //БФМИ.461272.035-06.01
          //  //listObg.Add(98970, 1);     //БФМИ.461271.070
          //  //listObg.Add(119424, 2);    //БФМИ.461271.099

          //  foreach (var obj in listObg)
          //  {
          //      addOrder(obj.Key, obj.Value);
          //  }
           
        }
        public void addOrder(int id, int count)
        {
            progressBar1.Value = 1;
            // выбор типа отчета
            if (rbPurchaseProducts.Checked) isPKIReport = true;
            else isPKIReport = false;

            if (cbFirstUnitCodeDevices.Checked) isFirstUnitCodeDevices = true;
            else isFirstUnitCodeDevices = false;

            if (cbOccurence.Checked) isOccurence = true;
            else isOccurence = false;

            if (isOccurence == false && isFirstUnitCodeDevices == false)
            {
                MessageBox.Show("Задайте для отчета 'Ведомость группового запуска' хотя бы одну из опций для списка входимостей");
                return;
            }

            // запрет изменения типа отчета
            reportPanel.Enabled = false;
            cbOccurence.Enabled = false;
            cbFirstUnitCodeDevices.Enabled = false;
            addProductButton.Enabled = false;
            deleteComplexButton.Enabled = false;
            clearDBButton.Enabled = false;
            reportButton.Enabled = false;

            progressBar1.Value = 2;

            // добавление заказа (построение дерева объектов и входимостей)
            var workingUser = PurchaseDesignProductOperation.BuildTree(Convert.ToInt32(id), isPKIReport, isFirstUnitCodeDevices, this.userFullName, count);
            if (workingUser)
                this.Close();

            // расчет количеств для отчета
            progressBar1.Value = 3;

            // PurchaseDesignProductOperation.Calculate(objectId);
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

            report.GetExcelReport(productIds, isPKIReport, isOccurence, isFirstUnitCodeDevices, reportFilePathInForm);

            addProductButton.Enabled = true;
            deleteComplexButton.Enabled = true;
            clearDBButton.Enabled = true;
            reportButton.Enabled = true;

            progressBar1.Value = 4;
            _attributesFormClose = true;
            PurchaseDesignProductOperation.DeleteUser();
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

        private void rbPurchaseProducts_CheckedChanged(object sender, EventArgs e)
        {
            if (rbPurchaseProducts.Checked)
            {
                cbOccurence.CheckState = CheckState.Checked;
                cbFirstUnitCodeDevices.CheckState = CheckState.Unchecked;
                cbOccurence.Enabled = false;
                cbFirstUnitCodeDevices.Enabled = false;
            }
        }

        private void rbGroupStart_CheckedChanged(object sender, EventArgs e)
        {
            if (rbGroupStart.Checked)
            {
                cbOccurence.Enabled = true;
                cbFirstUnitCodeDevices.Enabled = true;
            }
        }

        private void ComplexesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            PurchaseDesignProductOperation.DeleteUser();
        }
    }
}
