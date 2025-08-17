using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Procedures;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;

namespace ExecutorProcessReport
{
    public partial class SelectionForm : Form
    {
        private ReferenceObject selectedObject;
        private List<ProcedureReferenceObject> proceduresByObject  = new List<ProcedureReferenceObject>();
        public ReportParameters reportParameters = new ReportParameters();
        public SelectionForm()
        {
            InitializeComponent();
        }
        public void Init(ReferenceObject reportObject, List<ProcedureReferenceObject> procedures)
        {
            selectedObject = reportObject;
            reportParameters.IsTypeOrder = reportObject.Class.IsInherit(Guids.Ркк.Типы.НарядЗаказ);
            proceduresByObject.AddRange(procedures);
        }

        private void buttonMakeReport_Click_1(object sender, EventArgs e)
        {
            if (comboBoxState.Text == "Текущее состояние")
            {
                reportParameters.IsCurrentState = true;
                reportParameters.Chiefs = radioButtonChiefs.Checked;
                SetSettingsObjects();
            }
            else
            {
                if (comboBoxState.Text == "Все состояния")
                {
                    reportParameters.IsAllState = true;
                    SetSettingsObjects();
                    foreach (var states in StatesByProc.Values)
                    {
                        reportParameters.StatesObject.AddRange(states.Where(t => t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Работа) || t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Согласование)));
                    }
                }
                else
                {
                    reportParameters.IsSelectedState = true;
                    reportParameters.StatesObject.Add(StatesByProc.First(t => t.Key.Name == comboBoxProcess.Text).Value.OrderBy(t => t.ToString()).First(t => t.ToString() == comboBoxState.Text));
                }
            }

            reportParameters.ProcRefObj = StatesByProc.Where(t => t.Key.Name == comboBoxProcess.Text).First().Key;
            reportParameters.BeginPeriod = dateTimePickerBegin.Value.Date;
            reportParameters.EndPeriod = dateTimePickerEnd.Value.Date;
            reportParameters.IsControlState = checkBoxIsState.Checked;
            reportParameters.CountControlDays = Convert.ToInt32(NumericUDCountControlDays.Value);
            DialogResult = DialogResult.OK;
        }

        private void SetSettingsObjects()
        {
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id);
            var reference = referenceInfo.CreateReference();

            if (radioButtonAllObjects.Checked || radioButtonChiefs.Checked)
            {
                reportParameters.IsAllObjects = true;
            }
            else
            {
                if (radioButtonObjectsByDepartment.Checked)
                {
                    var filterDepartment = new Filter(ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id));
                    filterDepartment.Terms.AddTerm("[" + Guids.ГруппыИПользователи.Поля.Наименование + "]", ComparisonOperator.Equal, comboBoxDepartment.Text.Trim());
                    reportParameters.Department = reference.Find(filterDepartment).FirstOrDefault();
                    reportParameters.IsObjectsByDepartment = true;
                }
                else
                {
                    if (radioButtonObjectsByAuthor.Checked)
                    {
                        var filterUser = new Filter(ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id));
                        filterUser.Terms.AddTerm("[" + Guids.ГруппыИПользователи.Поля.КороткоеИмя + "]", ComparisonOperator.Equal, comboBoxAuthor.Text.Trim());
                        reportParameters.Author = reference.Find(filterUser).FirstOrDefault();
                        reportParameters.IsObjectsByAuthor = true;
                    }
                }
            }
        }

        Dictionary<ProcedureReferenceObject, List<ReferenceObject>> StatesByProc = new Dictionary<ProcedureReferenceObject, List<ReferenceObject>>();

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            dateTimePickerBegin.Value = DateTime.Now.AddMonths(-1).Date;
            comboBoxProcess.Items.AddRange(proceduresByObject.Select(t => t.Name.ToString()).OrderBy(t => t).ToArray());

            if (comboBoxProcess.Items.Count > 0)
            {
                comboBoxProcess.Text = comboBoxProcess.Items[0].ToString();
            }
            var textComboBoxProcess = proceduresByObject.Where(t => t.Name.ToString() == selectedObject.Class.ToString());

            if (textComboBoxProcess.Count() > 0)
            {
                comboBoxProcess.Text = textComboBoxProcess.First().ToString();
            }

            foreach (var proc in proceduresByObject)
            {
                StatesByProc.Add(proc, proc.GetObjects(Guids.Процедуры.СписокОбъектов.Состояния.Id).Where(t=>t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Работа) || t.Class.IsInherit(Guids.Процедуры.СписокОбъектов.Состояния.Типы.Согласование)).ToList());
            }

            comboBoxState.Items.Add("Текущее состояние");
            comboBoxState.Items.Add("Все состояния");

            if (ServerGateway.Connection.ClientView.GetUser().Class.IsInherit(Guids.ГруппыИПользователи.Типы.Администратор))
            {
                comboBoxState.Items.AddRange(StatesByProc.Where(t => t.Key.Name == comboBoxProcess.Text).First().Value.Select(t => t.ToString()).OrderBy(t => t.ToString()).Distinct().ToArray());
            }
            comboBoxState.Text = comboBoxState.Items[0].ToString();

            SetDeprtments();
            SetAuthor();
        }

        private void SetAuthor()
        {
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id);
            var reference = referenceInfo.CreateReference();
            var Сотрудник = referenceInfo.Classes.Find(Guids.ГруппыИПользователи.Типы.Сотрудник);
            var Администратор = referenceInfo.Classes.Find(Guids.ГруппыИПользователи.Типы.Администратор);
            var filterUsersByClass = new Filter(ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id));
            filterUsersByClass.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsOneOf, new List<ClassObject> { Сотрудник, Администратор});

            var сотрудники = reference.Find(filterUsersByClass).ToList();
            comboBoxAuthor.Items.AddRange(сотрудники.Select(t => t[Guids.ГруппыИПользователи.Поля.КороткоеИмя].ToString()).OrderBy(t => t).ToArray());
            comboBoxAuthor.Text = ServerGateway.Connection.ClientView.GetUser()[Guids.ГруппыИПользователи.Поля.КороткоеИмя].ToString();
        }


        private void SetDeprtments()
        {
            var текущийПользователь = ServerGateway.Connection.ClientView.GetUser();
            var подразделенияТекущегоПользователя = ПолучитьПодразделения(new List<ReferenceObject> { текущийПользователь }).Distinct();
            var referenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id);
            var reference = referenceInfo.CreateReference();
            var производственноеПодразделение = referenceInfo.Classes.Find(Guids.ГруппыИПользователи.Типы.ПроизводственноеПодразделение);
            var filterAllObjectByClass = new Filter(ServerGateway.Connection.ReferenceCatalog.Find(Guids.ГруппыИПользователи.Id));
            filterAllObjectByClass.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class], ComparisonOperator.IsInheritFrom, производственноеПодразделение);
            filterAllObjectByClass.Terms.AddTerm("[" + Guids.ГруппыИПользователи.Поля.Номер + "]", ComparisonOperator.IsNotEmptyString, null);

            var подразделения = reference.Find(filterAllObjectByClass).ToList();
            comboBoxDepartment.Items.AddRange(подразделения.Select(t => t[Guids.ГруппыИПользователи.Поля.Наименование].ToString()).OrderBy(t => t).ToArray());
            comboBoxDepartment.Text = подразделенияТекущегоПользователя.First().ToString();
        }

        private List<ReferenceObject> ПолучитьПодразделения(List<ReferenceObject> родительскиеОбъекты)
        {
            var объектыПодразделения = new List<ReferenceObject>();
            foreach (var родитель in родительскиеОбъекты)
            {
                if (родитель.Class.IsInherit(Guids.ГруппыИПользователи.Типы.ПроизводственноеПодразделение) && 
                    !string.IsNullOrEmpty(родитель[Guids.ГруппыИПользователи.Поля.Номер].ToString()))
                {
                    объектыПодразделения.Add(родитель);
                }
                объектыПодразделения.AddRange(ПолучитьПодразделения(родитель.Parents.ToList()));
            }
            return объектыПодразделения;
        }

        private void comboBoxProcess_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxProcess.Text.Trim() != string.Empty && StatesByProc.Count > 0)
            {
                var selectedProc = StatesByProc.Where(t => t.Key.Name == comboBoxProcess.Text).First();
                comboBoxState.Items.Clear();
                comboBoxState.Items.Add("Текущее состояние");
                comboBoxState.Items.Add("Все состояния");
                if (ServerGateway.Connection.ClientView.GetUser().Class.IsInherit(Guids.ГруппыИПользователи.Типы.Администратор))
                {
                    comboBoxState.Items.AddRange(selectedProc.Value.Select(t => t.ToString()).OrderBy(t => t.ToString()).Distinct().ToArray());
                }

                comboBoxState.Text = comboBoxState.Items[0].ToString();
            }
        }

        private void comboBoxState_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxState.Text == "Текущее состояние")
            {
                groupBoxSelectPeriod.Visible = false;
                groupBoxSelectObjects.Visible = true;
                groupBoxSelectPeriod.Top = 178;
                buttonMakeReport.Top = 283;
                Height = 357;
            }
            else
            {
                groupBoxSelectPeriod.Visible = true;
                if (comboBoxState.Text == "Все состояния")
                {
                    groupBoxSelectPeriod.Top = 287;
                    buttonMakeReport.Top = 368;
                    Height = 445;
                }
                else
                {    
                    groupBoxSelectObjects.Visible = false;
                    groupBoxSelectPeriod.Top = 178;
                    buttonMakeReport.Top = 283;
                    Height = 357;
                }
            }
        }

        private void radioButtonObjectsMyDepartment_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxDepartment.Enabled = true;
            comboBoxAuthor.Enabled = false;
        }

        private void radioButtonMyObjects_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxDepartment.Enabled = false;
            comboBoxAuthor.Enabled = true;
        }

        private void radioButtonAllObjects_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxDepartment.Enabled = false;
            comboBoxAuthor.Enabled = false;
        }
    }

    public class ReportParameters
    {
        public DateTime BeginPeriod;
        public DateTime EndPeriod;
        public List<ReferenceObject> StatesObject = new List<ReferenceObject>();
        public ProcedureReferenceObject ProcRefObj = null;
        public bool IsTypeOrder = false;
        public bool IsAllObjects = false;
        public bool IsObjectsByDepartment = false;
        public bool IsObjectsByAuthor = false;
        public ReferenceObject Author;
        public ReferenceObject Department;
        public bool IsAllState = false;
        public bool IsCurrentState = false;
        public bool IsSelectedState = false;
        public bool IsControlState = false;
        public int CountControlDays;
        public bool Chiefs = false;
    }
}
