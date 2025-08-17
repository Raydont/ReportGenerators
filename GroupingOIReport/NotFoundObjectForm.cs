using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model.References.Nomenclature;

namespace GroupingOIReport
{
    public partial class NotFoundObjectForm : Form
    {
        public NotFoundObjectForm()
        {
            InitializeComponent();
        }

        public List<DeviceByGroups> SortedData = new List<DeviceByGroups>();
       

        private void buttonMoving_Click(object sender, EventArgs e)
        {
            var selectedDevice = (DeviceByGroups)listViewNotFoundObjects.FocusedItem.Tag;
            var selectedObject = selectedDevice.NotFoundedObject.Where(t => t.NomObject[Guids.NomenclatureParameters.NameGuid].Value.ToString() == listViewNotFoundObjects.FocusedItem.SubItems[1].Text &&
            t.NomObject[Guids.NomenclatureParameters.DenotationGuid].Value.ToString() == listViewNotFoundObjects.FocusedItem.SubItems[2].Text).FirstOrDefault();

            string nameGroup = comboBoxSections.Text;
            int errorId = 0;
            try
            {
                errorId = 1;
                if (GroupsReferenceNames.supplyGroup.Keys.ToList().Contains(nameGroup))
                {
                    errorId = 2;
                    foreach (var obj in SortedData)
                    {
                        errorId = 3;
                        //Устройство
                        if (obj.Device == selectedDevice.Device)
                        {
                            errorId = 4;
                            bool foundedObj = false;
                            //Группы и списки изделий
                            for (int i = 0; i < obj.Groups.Count; i++)
                            {
                                errorId = 5;
                                if (obj.Groups.Keys.ToList()[i] == nameGroup)
                                {
                                    errorId = 6;
                                    //Список изделий
                                    for (int j = 0; j < obj.Groups[nameGroup].Count; j++)
                                    {
                                        errorId = 7;
                                        if (obj.Groups[nameGroup][j].NomObject == selectedObject.NomObject)
                                        {
                                            errorId = 8;
                                            obj.Groups[nameGroup][j].Count += selectedObject.Count;
                                            errorId = 9;
                                            if (selectedObject.Remarks.Count > 0)
                                                obj.Groups[nameGroup][j].AddRemarks(selectedObject.Remarks);
                                            foundedObj = true;
                                            break;
                                        }
                                    }
                                }
                              
                            }
                            if (!foundedObj)
                            {
                                errorId = 10;
                                obj.Groups.Where(t=>t.Key == nameGroup).ToList()[0].Value.Add(selectedObject);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("54 Внимание ошибка! Сформированный отчет будет некорректен. Обратитесь в отдел 911. Код ошибки " + errorId + " \r\n " + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            foreach (ListViewItem obj in listViewNotFoundObjects.SelectedItems)
            {
                listViewNotFoundObjects.Items.Remove(obj);
            }
            CalcIndex();
            buttonMoving.Enabled = false;
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

        private void buttonAddObjects_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
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
