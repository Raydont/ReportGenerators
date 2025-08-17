using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;

namespace EntrancesObjectsNotInNomReferences
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public ReportParameters reportParameters;
        public SelectionForm()
        {
            InitializeComponent();
        }
        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void buttonMakeReport_Click_1(object sender, EventArgs e)
        {
            reportParameters = new ReportParameters(radioButtonStandartItems.Checked, radioButtonOtherItems.Checked, radioButtonMaterials.Checked);
            reportParameters.MissingObjects = comboBoxTypeReport.Text == "Отсутствующие объекты справочника" ? true : false;
            reportParameters.MissingObjectsWithEntrances = comboBoxTypeReport.Text == "Отсутствующие объекты справочника и их входимости" ? true : false;
            reportParameters.ObjectsEqualRefNomWithRemarks = comboBoxTypeReport.Text == "Объекты справочника, соответствующие номенклатурным с примечанием" ? true : false;
            reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarks = comboBoxTypeReport.Text == "Объекты справочника и их входимости, соответствующие номенклатурным с примечанием" ? true : false;
            reportParameters.OrderWithRemarksStdItems = comboBoxTypeReport.Text == "Заказы со списком прямых вхождений Стандартных изделий с примечаниями" ? true : false;
            reportParameters.ObjectsEqualWithEntrancesRefNomWithRemarksWithOrder = comboBoxTypeReport.Text == "Объекты справочника и их входимости, соответствующие номенклатурным с примечанием (со списокм заказов)" ? true : false;
            reportParameters.PathForOrdersFiles = textBoxPathFolder.Text;

            DialogResult = DialogResult.OK;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            comboBoxTypeReport.Text = "Отсутствующие объекты справочника";
            textBoxPathFolder.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString();
        }

        private void textBoxPathFolder_TextChanged(object sender, EventArgs e)
        {
          

        }

        private void comboBoxTypeReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxTypeReport.Text == "Заказы со списком прямых вхождений Стандартных изделий с примечаниями")
            {
                labelChangeFolder.Visible = true;
                textBoxPathFolder.Visible = true;
                buttonChangeFolder.Visible = true;

            }
            else
            {
                labelChangeFolder.Visible = false;
                textBoxPathFolder.Visible = false;
                buttonChangeFolder.Visible = false;
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void buttonChangeFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxPathFolder.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }

    public class ReportParameters
    {
        public bool StandartItemsReference = false;
        public bool OtherItemsReference = false;
        public bool MaterialsReference = false;
        public string ObjectTypes = string.Empty;
        public bool MissingObjects = false;
        public bool MissingObjectsWithEntrances = false;
        public bool ObjectsEqualRefNomWithRemarks = false;
        public bool ObjectsEqualWithEntrancesRefNomWithRemarks = false;
        public bool OrderWithRemarksStdItems = false;
        public bool ObjectsEqualWithEntrancesRefNomWithRemarksWithOrder = false;
        public bool OrdersWithObjectsEqualWithEntrancesRefNomWithRemarks = false;
        public string PathForOrdersFiles = string.Empty;

        public ReportParameters(bool standartItemsReference, bool otherItemsReference, bool materialsReference)
        {
            StandartItemsReference = standartItemsReference;
            OtherItemsReference = otherItemsReference;
            MaterialsReference = materialsReference;
        }
    }
}
