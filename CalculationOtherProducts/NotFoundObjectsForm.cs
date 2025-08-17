using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.Mail;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Users;
using TFlex.DOCs.Model.Structure;

namespace CalculationOtherProducts
{
    public partial class NotFoundObjectsForm : Form
    {
        public NotFoundObjectsForm()
        {
            InitializeComponent();
        }

        public Dictionary<string, List<StructureTableData>> StructuredData = new Dictionary<string, List<StructureTableData>>();
        public List<TableData> addedTableData = new List<TableData>();

        private void buttonMoving_Click(object sender, EventArgs e)
        {
            int errorID = 0;
            try
            {
                //  MessageBox.Show(string.Join("\r\n", StructuredData.Select(t=> t.Key + "\r\n" + string.Join("\r\n", t.Value.Select(r=>"       -" + r.NameSections)))));
              // MessageBox.Show(string.Join("\r\n", StructuredData.Select(t => t.Key + "\r\n" + string.Join("\r\n", t.Value.Select(r => "       -" + r.NameSections)))));
                foreach (var firstSection in StructuredData)
                {
                    errorID = 1;
                    //Если для переноса выбран корневой заголовок 
                    if (comboBoxSections.Text == firstSection.Key)
                    {
                        errorID = 2;
                        // Если элементы раздела отсутствуют
                        if (firstSection.Value == null || firstSection.Value.Count == 0)
                        {
                            errorID = 3;
                            //  MessageBox.Show("1 vALUE.cOUNT = 0        ");
                            StructureTableData strData = new StructureTableData();
                            strData.NameSections = string.Empty;
                            List<TableData> data = strData.Data;
                            List<TableData> addedData = new List<TableData>();
                            errorID = 4;
                            foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
                            {
                                errorID = 5;
                                data.Add((TableData)obj.Tag);
                                addedData.Add((TableData)obj.Tag);
                            }

                            errorID = 6;
                           // strData.Data = new List<TableData>();

                            strData.Data.AddRange(SumDublicates(data));
                            firstSection.Value.Add(strData);
                        }
                        else
                        {
                            errorID = 7;
                            var emptySection = firstSection.Value.Where(t => t.NameSections.Trim() == string.Empty || t.NameSections.Trim() == null);

                            if (emptySection == null || emptySection.Count() == 0)
                            {
                                StructureTableData strData = new StructureTableData();
                                strData.NameSections = string.Empty;
                                List<TableData> data = new List<TableData>();
                                foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
                                {
                                    data.Add((TableData)obj.Tag);
                                }
                                strData.Data.AddRange(SumDublicates(data));
                                firstSection.Value.Add(strData);
                                errorID = 8;
                            }
                            else
                            {
   
                                //ReferenceObject ttu = null;

                                //ReferenceObject tt = ttu.Links.ToMany[new Guid("")].Objects.OrderByDescending(t => t.SystemFields.CreationDate).ToList().Last();

                                //Если пустой раздел есть
                                List<TableData> data = new List<TableData>();
                                foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
                                {
                                    data.Add((TableData)obj.Tag);
                                }

                                var sortData = firstSection.Value.Where(t => t.NameSections.Trim() == string.Empty).ToList()[0].Data;
                                sortData.AddRange(data);
                                firstSection.Value.Where(t => t.NameSections.Trim() == string.Empty).ToList()[0].Data = new List<TableData>();
                                firstSection.Value.Where(t => t.NameSections.Trim() == string.Empty).ToList()[0].Data.AddRange(SumDublicates(sortData));
                                errorID = 9;
                            }
                        }
                    }

                    foreach (var secondSections in firstSection.Value)
                    {
                        errorID = 10;
                        if (comboBoxSections.Text == secondSections.NameSections)
                        {
                            errorID = 11;
                            List<TableData> data = new List<TableData>();
                            string regularExpression = string.Empty;

                            if (secondSections.Data.Count > 0)
                            {
                                regularExpression = secondSections.Data[0].RegularExpression;
                            }

                            foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
                            {
                                var tableData = (TableData)obj.Tag;
                                tableData.ShortName = Regex.Match(tableData.Name, regularExpression).Value.ToString();

                                if (tableData.ShortName == string.Empty)
                                {
                                    tableData.ShortName = tableData.Name;
                                }
                                data.Add(tableData);
                            }
                            secondSections.Data.AddRange(SumDublicates(data));
                        }                    
                    }      
                }

                if (StructuredData.Count(t => t.Value.Count(k => k.NameSections == comboBoxSections.Text) == 0) == StructuredData.Count() &&
                    StructuredData.Count(t=>t.Key == comboBoxSections.Text)==0)
                {
                    MessageBox.Show("Внимание ошибка! Секция " + comboBoxSections.Text + " не найдена.Обратитесь в отдел 911", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
                {
                    errorID = 12;
                    listViewNotFoundObjects.Items.Remove(obj);
                    TableData item = (TableData)obj.Tag;
                    item.Section = comboBoxSections.Text;
                    addedTableData.Add(item);
                }
                CalcIndex();
                buttonMoving.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Внимание ошибка! Обратитесь в отдел 911. Код ошибки 238.\r\n" + ex + "\r\n" + errorID, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Рассылка уведомления о корректировке рег выражений 
        private void Mailing()
        {
            MailMessage message = new MailMessage();
            var usersReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Users);
            var usersReference = usersReferenceInfo.CreateReference();
            User userGAV = ((User)usersReference.Find(Guids.GAVGuid));
            var orderAddedTableData = new List<TableData>();
            orderAddedTableData.AddRange(addedTableData.OrderBy(t => t.Section).ThenBy(t => t.Denotation).ThenBy(t => t.Name).ToList());

            //Тема сообщения
            message.Subject = "Изменение рег выражений для справочника Категории прочих изделий для " + (ReferenceObject)this.Tag;
            //Текст сообщения
            message.Body = string.Join("\r\n", orderAddedTableData.Select(t=>t.Name + " " + t.Denotation).ToList()) + "\r\n\r\n" +
                string.Join("\r\n", orderAddedTableData.Select(t => t.ShortName).ToList()) + "\r\n\r\n" +
                string.Join("\r\n", orderAddedTableData.Select(t => t.Section).ToList());

            message.To.Add(new MailUser(userGAV));
            message.Send();        
        }

        static public List<TableData> SumDublicates(List<TableData> data)
        {

            for (int mainID = 0; mainID < data.Count; mainID++)
            {
                for (int slaveID = mainID + 1; slaveID < data.Count; slaveID++)
                {
                    //если документы имеют одинаковые обозначения и наименования - удаляем повторку
                    if (data[mainID].ShortName.Trim() == data[slaveID].ShortName.Trim())
                    {
                        data[mainID].Count += data[slaveID].Count;
                        data.RemoveAt(slaveID);
                        slaveID--;
                    }
                }
            }
            return data;
        }

        private void CalcIndex()
        {
            int i = 0;
            foreach (ListViewItem item in listViewNotFoundObjects.Items)
            {
                i++;
                item.SubItems[0].Text = i.ToString();
            }
        }

        private void listViewNotFoundObjects_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var modifyNameForm = new ModifyNameForm();
            modifyNameForm.Visible = false;

            modifyNameForm.labelNameObject.Text = listViewNotFoundObjects.SelectedItems[0].SubItems[1].Text;
            modifyNameForm.textBoxShortName.Text = listViewNotFoundObjects.SelectedItems[0].SubItems[2].Text;

            if (modifyNameForm.ShowDialog() == DialogResult.OK)
            {
                listViewNotFoundObjects.SelectedItems[0].SubItems[2].Text = modifyNameForm.textBoxShortName.Text;
                TableData item = (TableData)listViewNotFoundObjects.SelectedItems[0].Tag;
                item.ShortName = modifyNameForm.textBoxShortName.Text;
                listViewNotFoundObjects.SelectedItems[0].Tag = item;
            }
        }

        private void buttonAddObjects_Click(object sender, EventArgs e)
        {
            Mailing();
            DialogResult = DialogResult.OK;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void listViewNotFoundObjects_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            buttonMoving.Enabled = true;
        }

        private void listViewNotFoundObjects_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            buttonMoving.Enabled = true;
        }

        private void listViewNotFoundObjects_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((listViewNotFoundObjects.Items.Count != 0) && (listViewNotFoundObjects.SelectedItems.Count != 0) && (listViewNotFoundObjects.FocusedItem.Selected))
            {
                buttonMoving.Enabled = true;

            }
            else
            {
                buttonMoving.Enabled = false;
            }
        }
    }
}
