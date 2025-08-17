using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ListElementsReport
{
    public partial class AddCommentForm : Form
    {
        public AddCommentForm()
        {
            InitializeComponent();
        }

        private void listViewObjectsReport_MouseDoubleClick(object sender, MouseEventArgs e)
        {        
            label1.Text = listViewObjectsReport.SelectedItems[0].SubItems[1].Text;
        }

        private void SelectionRezCondForm_Load(object sender, EventArgs e)
        {

        }
        public List<ElementPki> elementPKIs = new List<ElementPki>();
        public void CopyItemsListView()
        {         
            foreach(ListViewItem item in listViewObjectsReport.Items)
            {
                var pki = (ElementPki)item.Tag;              
                elementPKIs.Add(pki);
            }
        }

        private void SelectionRezCondForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }

        private void buttonWriteToReport_Click(object sender, EventArgs e)
        {
            CopyItemsListView();
            this.DialogResult = DialogResult.OK;
        }

        private void listViewObjectsReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonDeleteComment.Enabled = true;
            buttonWriteObject.Enabled = true;
            label1.Text = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).Name;
            label1.Tag = listViewObjectsReport.FocusedItem.Tag;
            comboBoxRangeValue.Items.Clear();
            var range = ((ElementPki)listViewObjectsReport.FocusedItem.Tag).dbElements[0].RangeValue.ToArray();
            comboBoxRangeValue.Items.AddRange(range);
            if (range.Length > 0)
                comboBoxRangeValue.Text = range[0].ToString();
        }

        private void buttonDeleteComment_Click(object sender, EventArgs e)
        {
            listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[3].Text = string.Empty;
            ElementPki pki = (ElementPki)label1.Tag;
            pki.Comment = string.Empty;
            label1.Tag = pki;
            listViewObjectsReport.FocusedItem.Tag = pki; 
            buttonDeleteComment.Enabled = false;
        }

        private void buttonWriteObject_Click(object sender, EventArgs e)
        {
            if (label1.Text.Trim() == string.Empty)
                return;         
            ElementPki pki = (ElementPki)label1.Tag;

            if (pki.Comment.Trim() == string.Empty)
            {
                pki.Comment = comboBoxRangeValue.Text.Trim();
            }
            else
            {
                pki.Comment += "; " +comboBoxRangeValue.Text.Trim();
            }

            foreach (ListViewItem item in listViewObjectsReport.Items)
            {
                if(((ElementPki)item.Tag).Name == pki.Name)
                {
                    item.Tag = pki;
                    listViewObjectsReport.Items[listViewObjectsReport.FocusedItem.Index].SubItems[3].Text = pki.Comment;
                  //  comboBoxRangeValue.Text = string.Empty;
                    return;
                }
            }
        }

        private void SelectionRezCondForm_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void buttonSeparator_Click(object sender, EventArgs e)
        {
            
        }

        private void buttonSeparator_Click_1(object sender, EventArgs e)
        {
            comboBoxRangeValue.Text = comboBoxRangeValue.Text + "; ";
        }
    }
}
