using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.UI.Common;

namespace SelectionByTypeReport
{
    public partial class ExtractFilesByListDialog : Form
    {
        public string folderToSaveFiles;

        public ExtractFilesByListDialog()
        {
            InitializeComponent();
        }
       // public bool notFolderDialog = false;
        private void ExtractButton_Click(object sender, EventArgs e)
        {
            //if (notFolderDialog)
            //{
            //    DialogResult = DialogResult.Abort;
            //    return;
            //}
                
            var selectFolderDialog = new SelectFolderDialog();
            if (selectFolderDialog.ShowDialog(this) == DialogOpenResult.Ok)
            {
                folderToSaveFiles = selectFolderDialog.SelectedPath;

                DialogResult = DialogResult.OK;
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }
    }
}
