using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ListElementsReport
{
    public partial class AddAdmissionForm : Form
    {
        public AddAdmissionForm()
        {
            InitializeComponent();
        }

        private void listViewObjectsReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonWriteObject.Enabled = true;
            buttonDeleteAdmission.Enabled = true;
            label1.Text = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).Name;
            label1.Tag = listViewObjectsReport.FocusedItem.Tag;
            comboBoxAdmissions.Items.AddRange(((ElementPki)listViewObjectsReport.FocusedItem.Tag).dbElements[0].RangeValue.ToArray());
        }

        private void listViewObjectsReport_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            label1.Text = listViewObjectsReport.SelectedItems[0].SubItems[1].Text;
        }

        public List<ElementPki> elementPKIs = new List<ElementPki>();
        public void CopyItemsListView()
        {
            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                elementPKIs.Add((ElementPki)item.Tag);
            }
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
                    pki.OldName = listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text;
                    listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text = listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text.Replace(pki.Admission, comboBoxAdmissions.Text);
                    pki.Name = listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text;
                    item.Tag = pki;
                    return;
                }
            }
            
        }

        private void buttonWriteToReport_Click(object sender, EventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }

        private void AddAdmissionForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }

        private void buttonDeleteAdmission_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[1].Text = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).OldName;
            ElementPki pki = (ElementPki)label1.Tag;
            pki.Name = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).OldName;
            label1.Tag = pki;
            listViewObjectsReport.FocusedItem.Tag = pki;
            buttonDeleteAdmission.Enabled = false;
        }
    }
}
