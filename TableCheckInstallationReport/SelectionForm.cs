using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.Classes;
using TFlex.DOCs.Model.Structure;
using TFlex.Reporting;
using System.Text.RegularExpressions;

namespace TableCheckInstallationReport
{
    public partial class SelectionForm : Form
    {

        // Гуид справочника "Группы и пользователи"
        private static readonly Guid UsersAndGroupsReferenceGuid = new Guid("8ee861f3-a434-4c24-b969-0896730b93ea");
        //Тип справочника "Группы и пользователи" должность
        private static readonly Guid TypePostUsersGuid = new Guid("ea7c8581-c329-48bc-9e54-9becdccc5ae7");
        //параметр Наименование объекта справочника "Группыи пользователи"
        private static readonly Guid NameParametrGuid = new Guid("beb2d7d1-07ef-40aa-b45b-23ef3d72e5aa");
        //параметр Короткое имя справочника "Группы и пользователи"
        private static readonly Guid LastNameGuid = new Guid("76a97c36-f2d6-49ad-abb0-f5fdc91c071b");

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
            //Создаем ссылку на справочник Группы и пользователи
            ReferenceInfo info = ReferenceCatalog.FindReference(UsersAndGroupsReferenceGuid);
            Reference reference = info.CreateReference();

            //Находим тип «Должность»
            ClassObject classObject = info.Classes.Find(TypePostUsersGuid);
            //Создаем фильтр
            Filter filterPost = new Filter(info);
            //Добавляем условие поиска – «Тип = Должность»
            filterPost.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.Class],
                ComparisonOperator.Equal, classObject);

            //Получение списка конструкторов из своего подразделения
            var myParentDesign = ClientView.Current.GetUser().Parents.Where(t=>t[NameParametrGuid].Value.ToString() == "Конструктор").ToList();
            if (myParentDesign.Count() == 1)
            {
                var listMyDepartament = myParentDesign[0].Children.OrderBy(t => t[LastNameGuid].Value).Select(t => t[LastNameGuid].Value.ToString()).ToArray();
                cbSideBarSignName2.Items.AddRange(listMyDepartament);
                cbControlledBy.Items.AddRange(listMyDepartament);
                cbSideBarSignName2.Text = listMyDepartament[0];
                cbControlledBy.Text = listMyDepartament[0];
            }

            tbSideBarSignName1.Text = ClientView.Current.GetUser()[LastNameGuid].Value.ToString();
            tbDesignedBy.Text = ClientView.Current.GetUser()[LastNameGuid].Value.ToString();
   
            var myParents = ClientView.Current.GetUser().Parents;

            foreach (var parent in myParents)
            {
                foreach (var parentParent in parent.Parents)
                {

                    foreach (var child in parentParent.Children)
                    {
                        if (child[NameParametrGuid].Value.ToString() == "Начальник")
                        {
                            tbApprovedBy.Text = child.Children[0][LastNameGuid].Value.ToString();
                        }
                    }
                }
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
            try
            {
                reportParameters.X = Convert.ToInt32(Convert.ToDouble(comboBoxX.Text.Trim().Replace(".",","))*1000);
                reportParameters.Y = Convert.ToInt32(Convert.ToDouble(textBoxY.Text.Trim().Replace(".", ",")) * 1000);
            }
            catch
            {
                MessageBox.Show("Неправильный формат ввода X, Y! ","Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            reportParameters.DenotationDevice = textBoxDenotationDevice.Text.Trim();
            reportParameters.FirstUse = textBoxFirstUse.Text.Trim();

            if (radioButtonMPP.Checked && !radioButtonOMPP.Checked)
            {
                reportParameters.MPP = true;
            }
            else
            {
                reportParameters.MPP = false;
            }

         

            reportParameters.SignHeader1 = tbSideBarSignHeader1.Text.Trim();
            reportParameters.SignHeader2 = tbSideBarSignHeader2.Text.Trim();
            reportParameters.SignName1 = tbSideBarSignName1.Text.Trim();
            reportParameters.SignName2 = cbSideBarSignName2.Text.Trim();

            reportParameters.AuthorBy = tbDesignedBy.Text.Trim();
            reportParameters.MusteredBy = cbControlledBy.Text.Trim();
            reportParameters.NControlBy = tbNControlBy.Text.Trim();
            reportParameters.ApprovedBy = tbApprovedBy.Text.Trim();

            reportParameters.SignData1 = dtpDataSign1.Value;
            reportParameters.SignData2 = dtpDataSign2.Value;

            if (cbLetter.Text.Trim() != "-" && cbLetter.Text.Length > 0)
                reportParameters.Litera1 = cbLetter.Text[0];

            if (cbLetter.Text.Trim() != "-" && cbLetter.Text.Length > 1)
                reportParameters.Litera2 = cbLetter.Text[1];

            if (checkBoxLetter2.Checked)
            {
                if ((cbLetter.Text.Trim() == "-" || cbLetter.Text.Trim() == string.Empty) && cbLetter2.Text.Trim() != "-" && cbLetter2.Text.Length > 0)
                {
                    reportParameters.Litera1 = cbLetter2.Text[0];

                    if (cbLetter2.Text.Trim() != "-" && cbLetter2.Text.Length > 1)
                        reportParameters.Litera2 = cbLetter2.Text[1];
                }
                else
                {
                    if (cbLetter.Text.Trim().Length == 1)
                    {
                        if (cbLetter2.Text.Trim() != "-" && cbLetter2.Text.Length > 0)
                            reportParameters.Litera2 = cbLetter2.Text[0];

                        if (cbLetter2.Text.Trim() != "-" && cbLetter2.Text.Length > 1)
                            reportParameters.Litera3 = cbLetter2.Text[1];
                    }
                    else
                    {
                        if (cbLetter.Text.Trim().Length == 2 && cbLetter2.Text.Trim() != "-" && cbLetter2.Text.Length > 0)
                            reportParameters.Litera3 = cbLetter2.Text[0];
                    }
                }
            }
        
            reportParameters.NumberSol = tbNumberSol.Text.Trim();
            reportParameters.Code = tbCode.Text.Trim();

            reportParameters.AuthorCheck = (checkBoxAuthor.CheckState == CheckState.Checked) ? true : false;
            reportParameters.MusteredBCheck = (checkBoxMasteredBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.NControlByCheck = (checkBoxNControllBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.ApprovedByCheck = (checkBoxApprovedBy.CheckState == CheckState.Checked) ? true : false;
            reportParameters.DateStamp1 = (checkBoxDate1.CheckState == CheckState.Checked) ? true : false;
            reportParameters.DateStamp2 = (checkBoxDate2.CheckState == CheckState.Checked) ? true : false;
            reportParameters.PageUD = (checkBoxPageUD.CheckState == CheckState.Checked) ? true : false;

            reportParameters.AuthorByDate = dTPAuthor.Value;
            reportParameters.MusteredByDate = dTPMasteredBy.Value;
            reportParameters.NControlByDate = dTPNControlBy.Value;
            reportParameters.ApprovedByDate = dTPApprovedBy.Value;

            if (textBoxFileName.Text.Trim().ToLower().Contains(".ipc"))
            {
                try
                {
                  //  StreamReader streamReader = File.OpenText(textBoxFileName.Text.Trim());

                   // string read = null;

                    //  while ((read = streamReader.ReadLine()) != null)

                    var lines = File.ReadAllLines(textBoxFileName.Text.Trim(), Encoding.GetEncoding("Windows-1251"));

                    foreach (var read in lines)
                    {
                        reportParameters.FileData.Add(read);
                    }
                    reportParameters.newFileName = textBoxFileName.Text.Replace(openFileDialog.SafeFileName, "") + "NET_" + openFileDialog.SafeFileName.ToString();
                    var fileStream = File.Create(reportParameters.newFileName);
                    fileStream.Close();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Выберите корректный файл!\n" + e.ToString(), "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            if (textBoxFileName.Text.Trim().ToLower().Contains(".net"))
            {
                try
                {
                   // StreamReader streamReader = File.OpenText(textBoxFileName.Text.Trim());

                    //string read = null;

                    var listStr = new List<string>();
                    bool readSignals = false;

                    var lines = File.ReadAllLines(textBoxFileName.Text.Trim(), Encoding.GetEncoding("Windows-1251"));

                    foreach(var read in lines)
                    {
                        if (readSignals && read != "% END")
                            listStr.Add(read);
                        if (read == "% SIGNALS")
                            readSignals = true;
                       
                    }
                    string str = string.Empty;

                    for (int i = 0; i < listStr.Count; i++)
                    {

                        if (i != 0)
                        {
                            var prewListStr = listStr[i - 1].Trim();
                            if (prewListStr[prewListStr.Length - 1] == ',')
                            {
                                reportParameters.FileDataNet[reportParameters.FileDataNet.Count -1] += listStr[i].Trim();
                            }
                            else
                            {
                                reportParameters.FileDataNet.Add(listStr[i].Trim());
                            }
                        }
                        else
                        {
                            reportParameters.FileDataNet.Add(listStr[i].Trim());
                        }
                    }
                    reportParameters.newFileName = textBoxFileName.Text.Replace(openFileDialog.SafeFileName, "") + "NET_" + openFileDialog.SafeFileName.ToString();
                    var fileStream = File.Create(reportParameters.newFileName);
                    fileStream.Close();
                   // Clipboard.SetText(string.Join("\r\n*", reportParameters.FileDataNet));
                }
                catch (Exception e)
                {
                    MessageBox.Show("Выберите корректный файл!\n" + e.ToString(), "Ошибка!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            return true;

        }


        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBoxApprovedBy_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBoxApprovedBy.CheckState == CheckState.Checked)
                dTPApprovedBy.Visible = true;
            else
                dTPApprovedBy.Visible = false;
        }

        private void checkBoxNControllBy_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxNControllBy.CheckState == CheckState.Checked)
                dTPNControlBy.Visible = true;
            else
                dTPNControlBy.Visible = false;
        }

        private void checkBoxMasteredBy_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxMasteredBy.CheckState == CheckState.Checked)
                dTPMasteredBy.Visible = true;
            else
                dTPMasteredBy.Visible = false;
        }

        private void checkBoxAuthor_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAuthor.CheckState == CheckState.Checked)
                dTPAuthor.Visible = true;
            else
                dTPAuthor.Visible = false;
        }

        private void buttonLoadFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFileName.Text = openFileDialog.FileName;
                if (textBoxFileName.Text.ToLower().Contains(".ipc") || textBoxFileName.Text.ToLower().Contains(".net"))
                {
                    makeReportButton.Enabled = true;
                }
            }
        }

        private void checkBoxLetter2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxLetter2.Checked)
            {
                cbLetter2.Enabled = true;
            }
            else
            {
                cbLetter2.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
