using System;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.Reporting;

namespace ChargesReport
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

        private void MainForm_Load(object sender, EventArgs e)
        {
            comboBoxMethodPayment.Text = "сдельная";
            comboBoxTypeReport.Text = "По табельному времени";

            for (int i = 0; i <= 10; i++)
            {
                comboBoxYear.Items.Add((DateTime.Now.Year + 1 - i).ToString());
            }

            ReferenceObject folderObject;
            ReferenceInfo referenceChargesInfo = ReferenceCatalog.FindReference(ChargesParameters.ReferenceGuid.Charges);
            Reference referenceCharges = referenceChargesInfo.CreateReference();
            var refObject = referenceCharges.Find(_context.ObjectsInfo[0].ObjectId);
            //Находим тип «Папка»
            ClassObject classFolder = referenceChargesInfo.Classes.Find(ChargesParameters.ReferenceTypesGuid.Folder);

            if (refObject.Class.IsInherit(ChargesParameters.ReferenceTypesGuid.Charge))
            {
                folderObject = refObject.Parent;
            }
            else
            {
                if (refObject.Children[0].Class.IsInherit(classFolder))
                {
                    folderObject = null;
                }
                else
                {
                    folderObject = refObject;
                }
            }

            if (folderObject != null)
            {
                comboBoxDepartment.Text = folderObject.ToString();
                int i = 0;
                foreach (var month in comboBoxMonth.Items)
                {
                    if (folderObject.Parent.ToString().Trim().ToLower().Contains(month.ToString().Trim().ToLower()))
                    {
                        comboBoxMonth.SelectedIndex = i;
                        break;
                    }
                    i++;
                }
                comboBoxYear.Text = folderObject.Parent.Parent.ToString().Trim();
            }
            else
            {
                comboBoxDepartment.Text = "Цех 75";

                if (DateTime.Now.Month != 0)
                {
                    comboBoxMonth.SelectedIndex = DateTime.Now.Month - 1;
                }
                else
                {
                    comboBoxMonth.SelectedIndex = 12;
                }
                comboBoxYear.Text = DateTime.Now.Year.ToString();
            }
            SelectMonthHours(comboBoxMonth.Text);
        }

        private void SelectMonthHours(string nameMonth)
        {
            switch (nameMonth.Trim().ToLower())
            {
                case "январь":
                    textBoxCountWorkHours.Text = "135,25";
                    break;
                case "февраль":
                    textBoxCountWorkHours.Text = "160";
                    break;
                case "март":
                    textBoxCountWorkHours.Text = "167,25";
                    break;
                case "апрель":
                    textBoxCountWorkHours.Text = "174,25";
                    break;
                case "май":
                    textBoxCountWorkHours.Text = "144,75";
                    break;
                case "июнь":
                    textBoxCountWorkHours.Text = "151,50";
                    break;
                case "июль":
                    textBoxCountWorkHours.Text = "184,75";
                    break;
                case "август":
                    textBoxCountWorkHours.Text = "167";
                    break;
                case "сентябрь":
                    textBoxCountWorkHours.Text = "176,50";
                    break;
                case "октябрь":
                    textBoxCountWorkHours.Text = "183,50";
                    break;
                case "ноябрь":
                    textBoxCountWorkHours.Text = "150,75";
                    break;
                case "декабрь":
                    textBoxCountWorkHours.Text = "176,50";
                    break;
                default:
                    break;
            }
        }

        private void makeReportButton_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }


        private bool MakeReportParameters()
        {
            reportParameters = new ReportParameters();
            reportParameters.Month = comboBoxMonth.Text;
            reportParameters.MonthInt = Convert.ToInt32(comboBoxMonth.SelectedIndex + 1);
            reportParameters.Year = Convert.ToInt32(comboBoxYear.Text.Trim());
            reportParameters.Department = comboBoxDepartment.Text;
            reportParameters.TypeReport = comboBoxTypeReport.Text;

            try
            {
                reportParameters.CountWorkHours = Convert.ToDouble(textBoxCountWorkHours.Text);
            }
            catch
            {
                MessageBox.Show("Количество рабочих часов указано в неверном формате. Сформировать ведомость нельзя.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (comboBoxMethodPayment.Text == "сдельная")
                reportParameters.MethodPayment = "сдельщикам";
            else
                reportParameters.MethodPayment = "повременщикам";
            
            return true;
        }

        private void comboBoxMethodPayment_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxMethodPayment.Text == "сдельная")
            {
                comboBoxTypeReport.Enabled = false;
                comboBoxTypeReport.Text = "По табельному времени";
            }
            else
            {
                comboBoxTypeReport.Enabled = true; 
            }
        }

        private void comboBoxMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectMonthHours(comboBoxMonth.Text);
        }
    }
}
