using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;

namespace OfficeEquipmentReport
{
    public partial class SelectionForm : Form
    {
        private IReportGenerationContext _context;
        public SelectionForm()
        {
            InitializeComponent();
        }
        public ReportParameters reportParameters = new ReportParameters();

        public void Init(IReportGenerationContext context)
        {
            _context = context;
        }

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            //Создаем ссылку на справочник
            ReferenceInfo info = ReferenceCatalog.FindReference(Guids.ReferencesGuid.DepartmentsGuid);
            Reference reference = info.CreateReference();
            //Находим тип «Производственное подразделение»
            ClassObject classObject = info.Classes.Find(Guids.DepartmentTypes.DepartmentTypeGuid);
            //Создаем фильтр
            Filter filter = new Filter(info);
            //Добавляем условие поиска – «Тип порожден от Подразделение»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.IsInheritFrom, classObject);
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр 
            var departaments = reference.Find(filter).OrderBy(t => t.ToString().Trim()).ToArray();
            comboBoxDepartments.Items.AddRange(departaments);
           // var deparmentObject = FindDepartment(ClientView.Current.GetUser());
            var userDepartment = departaments.Where(t => t.ToString().ToLower().Trim() == reference.Find(_context.ObjectsInfo[0].ObjectId).ToString().ToLower().Trim()).ToList().FirstOrDefault();
            comboBoxDepartments.SelectedItem = userDepartment == null? comboBoxDepartments.Items[0]:userDepartment;
        }

        //public ReferenceObject FindDepartment(ReferenceObject refObj)
        //{
        //    foreach (var parent in refObj.Parents)
        //    {
        //        if (parent[Guids.ReferencesGuid.GroupAndUserGuid].GetString() == "АО РКБ Глобус")
        //        {
        //            if (refObj[Guids.ReferencesGuid.GroupAndUserGuid].ToString() != "Получатели рассылок" && refObj[Guids.ReferencesGuid.GroupAndUserGuid].ToString() != "КИЦ")
        //                return refObj;
        //        }
        //        var department = FindDepartment(parent);
        //        if (department != null && department[Guids.ReferencesGuid.GroupAndUserGuid].ToString() != "Получатели рассылок" && refObj[Guids.ReferencesGuid.GroupAndUserGuid].ToString() != "КИЦ") return department;
        //    }

        //    return null;
        //}

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }

        private bool MakeReportParameters()
        {
            reportParameters = new ReportParameters();

            reportParameters.Department = (ReferenceObject)comboBoxDepartments.SelectedItem;
            reportParameters.Computer = checkBoxComputer.CheckState == CheckState.Checked ? true : false;
            reportParameters.Monitor = checkBoxMonitor.CheckState == CheckState.Checked ? true : false;
            reportParameters.Printer = checkBoxPrinter.CheckState == CheckState.Checked ? true : false;
            reportParameters.Scanner = checkBoxScanner.CheckState == CheckState.Checked ? true : false;
            reportParameters.Storage = checkBoxStorage.CheckState == CheckState.Checked ? true : false;
            reportParameters.Notebook = checkBoxNotebook.CheckState == CheckState.Checked ? true : false;
            reportParameters.MFU = checkBoxMFU.CheckState == CheckState.Checked ? true : false;
            reportParameters.OtherDevices = checkBoxOtherDevices.CheckState == CheckState.Checked ? true : false;

            return true;
        }
    }
}
