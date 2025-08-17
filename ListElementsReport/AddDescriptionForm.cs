using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References;

namespace ListElementsReport
{
    public partial class AddDescriptionForm : Form
    {
        public AddDescriptionForm()
        {
            InitializeComponent();
        }

        public static readonly Guid NameOtherItemsGuid = new Guid("4d3a600e-0dd2-4741-b045-c37b78fe88ae");
        public static readonly Guid DenotationItemsGuid = new Guid("27fe7568-2650-453e-a34c-a431ea0f5f00");
        private void listViewObjectsReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            buttonWriteObject.Enabled = true;
            buttonDeleteName.Enabled = true;
            comboBoxName.Items.Clear();
            var pki = (ElementPki)listViewObjectsReport.FocusedItem.Tag;
            label1.Text = pki.Name;

            label1.Tag = listViewObjectsReport.FocusedItem.Tag;
            var textNorm = new TextNormalizer();

            List<KeyValuePair<string, ReferenceObject>> findObjs;

            //Если выбран флажок Выбор ТКС, отбо идет по наименованию
            if (selectTKS)
                findObjs = normalRefObjects.Where(t => t.Key.Contains(
             textNorm.GetNormalForm(pki.Name.Replace("*", "").Replace(pki.dbElements[0].TU, "")))
                && t.Key.Contains(textNorm.GetNormalForm(pki.dbElements[0].TU))).ToList();
            else
                findObjs = normalRefObjects.Where(t => t.Key.Contains(
              textNorm.GetNormalForm(pki.dbElements[0].Comment + " " + pki.dbElements[0].Description.Replace("*", ""))) 
              && t.Key.Contains(textNorm.GetNormalForm(pki.dbElements[0].TU))).ToList();
          

            List < string> namesOtherItems = new List<string>();
            foreach(var obj in findObjs)
            {
                if (otherItems)
                    namesOtherItems.Add(obj.Value[NameOtherItemsGuid].Value + " " + obj.Value[DenotationItemsGuid].Value);
                else
                    namesOtherItems.Add(obj.Value[ParametersInfo.Name].Value + " " + obj.Value[ParametersInfo.Denotation].Value);
            }
            if (namesOtherItems.Count > 0)
            {
                comboBoxName.Items.AddRange(namesOtherItems.OrderBy(t => t).ToArray());
                comboBoxName.Text = comboBoxName.Items[0].ToString();
            }
        }

        public List<ElementPki> elementPKIs = new List<ElementPki>();
        public bool otherItems = false;
        public bool selectTKS = false;
        public void CopyItemsListView()
        {
            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                elementPKIs.Add((ElementPki)item.Tag);
            }
        }

        private void listViewObjectsReport_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            label1.Text = listViewObjectsReport.SelectedItems[0].SubItems[1].Text;
        }

        private void buttonWriteObject_Click(object sender, EventArgs e)
        {
            if (label1.Text.Trim() == string.Empty)
                return;
            ElementPki pki = (ElementPki)label1.Tag;
            buttonWriteObject.Enabled = false;
            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                if (((ElementPki)item.Tag).PosDenotation == pki.PosDenotation)
                {
                    foreach (ListViewItem it in listViewObjectsReport.SelectedItems)
                    {
                        //   pki.OldName = listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text;
                        pki.OldName = listViewObjectsReport.Items[it.Index].SubItems[1].Text;
                        listViewObjectsReport.Items[it.Index].SubItems[1].Text = comboBoxName.Text;
                        //    pki.Name = comboBoxName.Text;
                        //   pki.PosDenotation = ((ElementPKI)it.Tag).PosDenotation;

                        //  listViewObjectsReport.Items[it.Index].Tag = pki;
                        ((ElementPki)listViewObjectsReport.Items[it.Index].Tag).OldName = listViewObjectsReport.Items[it.Index].SubItems[1].Text;
                        ((ElementPki)listViewObjectsReport.Items[it.Index].Tag).Name = comboBoxName.Text;
                        // it.Tag = pki;
                      //  MessageBox.Show(((ElementPKI)listViewObjectsReport.Items[it.Index].Tag).PosDenotation + " " + ((ElementPKI)listViewObjectsReport.Items[it.Index].Tag).Name);

                    }
                    return;
                }
            }
        }

        private void buttonWriteToReport_Click(object sender, EventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }
        public Dictionary<string, ReferenceObject> normalRefObjects = new Dictionary<string, ReferenceObject>();
        private void AddDescriptionForm_Load(object sender, EventArgs e)
        {

        }

        private void AddDescriptionForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }

        private void buttonDeleteName_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).OldName;
            ElementPki pki = (ElementPki)label1.Tag;
            pki.Name = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).OldName;
            label1.Tag = pki;
            listViewObjectsReport.FocusedItem.Tag = pki;
            buttonDeleteName.Enabled = false;
        }
    }
}
