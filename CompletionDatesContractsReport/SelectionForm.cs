using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.Reporting;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.Model.References.Users;

namespace CompletionDatesContractsReport
{
    public partial class SelectionForm : Form
    {
        public SelectionForm()
        {
            InitializeComponent();
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        //GUID справочника «Группы и пользователи»
        private Guid GroupAndUsersGuid = new Guid ("8ee861f3-a434-4c24-b969-0896730b93ea");
        //GUID типа «Производственное подразделение»
        private Guid IndustrialDeptTypeGuid = new Guid("c2fb14b6-58a5-43e2-953e-a41166bff848");
        //GUID параметра «Наименование»
        private Guid FullNameGuid = new Guid ("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
        //Guid группы пользователей активные пользователи ОРД
        private Guid ActiveUserORD = new Guid("dd1e40ca-eb2e-4f04-b3c1-0ad4728e89b1");

        //Guid группы пользователей активные пользователи ОРД
        private Guid WatchDZUSZNO = new Guid("fb6444f6-b476-4bf7-acfd-2856c604fa34");
        //GUID типа "Пользователь"
        public Guid UserTypeGuid = new Guid("e280763e-ce5a-4754-9a18-dd17554b0ffd");
        //GUID номер подразделения
        public Guid NumberDepartamentsGuid = new Guid("1ff481a8-2d7f-4f41-a441-76e83728e420");

        private void SelectionForm_Load(object sender, EventArgs e)
        {

            TFlex.DOCs.UI.Objects.TFlexDOCsClientUI.Initialize();
            //Создаем ссылку на справочник
            ReferenceInfo info = ReferenceCatalog.FindReference(GroupAndUsersGuid);
            Reference reference = info.CreateReference();
            //Находим тип «Производственное подразделение»
            ClassObject classObject = info.Classes.Find(IndustrialDeptTypeGuid);
            //Создаем фильтр
            Filter filter = new Filter(info);
            //Добавляем условие поиска – «Тип порожден от Производственное подразделение»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.IsInheritFrom, classObject);
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр 
            var departaments = reference.Find(filter).OrderBy(t => t.ToString()).ToArray();
            comboBoxSpecificDepartament.Items.AddRange(departaments);
            var deparmentObject = FindDepartment(ClientView.Current.GetUser());
            comboBoxSpecificDepartament.SelectedItem = departaments.Where(t => t.ToString() == deparmentObject.ToString()).ToList()[0];
        
            //Проверка скорости работы отчета о сроках действия договора 
            //Предусмотреть быструю загрузка объектов 
            //Находим тип «Пользователь»
            ClassObject classObjectUser = info.Classes.Find(UserTypeGuid);
            //Создаем фильтр
            Filter filterUser = new Filter(info);
            //Добавляем условие поиска – «Тип = пользователь»
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, classObjectUser);
            //Получаем список объектов, в качестве условия поиска – сформированный фильтр  
            var users = reference.Find(filterUser).Where(t => t.Parents.Contains(reference.Find(WatchDZUSZNO))).OrderBy(t => t.ToString()).ToArray().ToList();
            if (!users.Contains(ClientView.Current.GetUser()))
                users.Add(ClientView.Current.GetUser());
            comboBoxSpecificMaker.Items.AddRange(users.ToArray());
           // comboBoxSpecificMaker.SelectedItem = users.Where(t => t.ToString() == ClientView.Current.GetUser().ToString()).ToList()[0];
            comboBoxOrderSum.SelectedIndex = 0;
        }
        public ReferenceObject FindDepartment(ReferenceObject refObj)
        {
            foreach (var parent in refObj.Parents)
            {
                if (parent[GroupAndUsersGuid].GetString() == "АО РКБ Глобус")
                {
                    if (refObj[GroupAndUsersGuid].ToString() != "Получатели рассылок" && refObj[GroupAndUsersGuid].ToString() != "КИЦ")
                        return refObj;
                }
                var department = FindDepartment(parent);
                if (department != null && department[GroupAndUsersGuid].ToString() != "Получатели рассылок" && refObj[GroupAndUsersGuid].ToString() != "КИЦ") return department;
            }

            return null;
        }
        public void Init(IReportGenerationContext context)
        {
            GetReportParameters getReportParameters = new GetReportParameters();
            AttributeReport attributeReport = new AttributeReport();
            attributeReport.AllJournalReport = radioButtonAllJournal.Checked;
            System.Windows.Forms.Application.DoEvents();
            if (radioButtonSpecificDepartament.Checked)
            {
                attributeReport.Maker = null;
                attributeReport.Department = null;
            }
            
            if (radioButtonSpecificDepartament.Checked)
            {
                attributeReport.DepartmentReport = radioButtonSpecificDepartament.Checked;
                attributeReport.Department = (UserReferenceObject)comboBoxSpecificDepartament.SelectedItem;
            }

            if (radioButtonSpecificMaker.Checked)
            {
                attributeReport.MakerReport = radioButtonSpecificMaker.Checked;
                attributeReport.Maker = (User)comboBoxSpecificMaker.SelectedItem;                
            }

            attributeReport.SumOrder500000 = false;
            if (comboBoxOrderSum.SelectedIndex == 1)
                attributeReport.SumOrder500000 = true;
           
            getReportParameters.FillRKK(context, attributeReport);
        }

        private void radioButtonSpecificDepartament_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSpecificDepartament.Checked == true)
            {
                comboBoxSpecificDepartament.Enabled = true;
                comboBoxSpecificMaker.Enabled = false;
            }
            else
            {
                comboBoxSpecificDepartament.Enabled = false;
                comboBoxSpecificMaker.Enabled = true;
            }
        }

        private void radioButtonSpecificMaker_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonSpecificMaker.Checked == true)
            {
                comboBoxSpecificDepartament.Enabled = false;
                comboBoxSpecificMaker.Enabled = true;
            }
            else
            {
                comboBoxSpecificDepartament.Enabled = true;
                comboBoxSpecificMaker.Enabled = false;
            }
        }

        private void radioButtonAllJournal_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxSpecificDepartament.Enabled = false;
            comboBoxSpecificMaker.Enabled = false;
        }
    }
}
