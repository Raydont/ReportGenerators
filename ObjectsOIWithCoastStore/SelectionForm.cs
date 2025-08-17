using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using TFlex.DOCs.Model;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Nomenclature;
using TFlex.DOCs.Model.Search;
using TFlex.DOCs.Model.Structure;
using TFlex.DOCs.UI.Common;
using TFlex.Reporting;

namespace ObjectsOIWithCoastStore
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

        ToolStripLabel dateTimeLabel;
        ToolStripLabel infoLabel;
        ToolStripProgressBar tsProgressBar;

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            dateTimeLabel = new ToolStripLabel();
            infoLabel = new ToolStripLabel();
            tsProgressBar = new ToolStripProgressBar();

            statusStrip.Items.Add(dateTimeLabel);
            statusStrip.Items.Add(infoLabel);
            statusStrip.Items.Add(tsProgressBar);

            var nomReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomReferenceInfo.CreateReference();

            var filter = new Filter(nomReferenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, _context.ObjectsInfo[0].ObjectId);

            NomenclatureObject findedObj = (NomenclatureObject)reference.Find(filter).FirstOrDefault();

            radioButtonForSelectedObject.Text = "Для " + findedObj.ToString();
        }

        private void buttonMakeReport_Click(object sender, EventArgs e)
        {
            if (MakeReportParameters())
                DialogResult = DialogResult.OK;
        }

        private bool CheckAndLoadObjectFailure()
        {
            var dateBegin = dateTimePickerBegin.Value.Date;
            var dateEnd = dateTimePickerEnd.Value.Date;

            if ((dateEnd - dateBegin).Days <= 0)
            {
                MessageBox.Show("Некорректно задан период!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            ReferenceInfo failureOIReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.Failure.RefId);
            Reference failureOIReference = failureOIReferenceInfo.CreateReference();

            ReferenceInfo storeObjectReferenceInfo = ServerGateway.Connection.ReferenceCatalog.Find(Guids.References.NomStoreGuid);
            Reference nomStoreReference = storeObjectReferenceInfo.CreateReference();

            var filterByPeriod = new Filter(failureOIReferenceInfo);
            filterByPeriod.Terms.AddTerm("[" + Guids.Failure.Params.DateSZ + "]", ComparisonOperator.GreaterThanOrEqual, dateBegin);
            filterByPeriod.Terms.AddTerm("[" + Guids.Failure.Params.DateSZ + "]", ComparisonOperator.LessThanOrEqual, dateEnd);
            var failureObjectsByPeriods = failureOIReference.Find(filterByPeriod);
            if (failureObjectsByPeriods == null || failureObjectsByPeriods.Count() == 0)
            {
                MessageBox.Show("Объектов за указанный период не найдено!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            var storeObjects = new List<StoreObject>();

            foreach(var failure in failureObjectsByPeriods.OrderBy(t=>t[Guids.Failure.Params.DateSZ].GetDateTime()))
            {
                foreach (var store in failure.GetObjects(Guids.Failure.Links.OtherItemsListObjects))
                {
                    var filterNomStore = new Filter(storeObjectReferenceInfo);
                    filterNomStore.Terms.AddTerm("[" + Guids.NomStore.NameObjectStore + "]", 
                        ComparisonOperator.Equal, 
                        store[Guids.Failure.Params.NameOtherItems].GetString());
                    var storeObject = nomStoreReference.Find(filterNomStore).FirstOrDefault();
                    if (storeObject != null)
                    {
                        var newStoreObject = new StoreObject(storeObject);
                        newStoreObject.Count = store[Guids.Failure.Params.Count].GetInt32();
                        newStoreObject.DateSZ = failure[Guids.Failure.Params.DateSZ].GetDateTime().Date;
                        newStoreObject.NumberSZ = failure[Guids.Failure.Params.NumberSZ].GetString();
                        storeObjects.Add(newStoreObject);
                    }
                }
            }

            reportParameters.StoreObjects.AddRange(storeObjects);
          //  MessageBox.Show(string.Join("\r\n", storeObjects.Select(t => t.Name + " " + t.Count)));

            return true;
        }

        private bool MakeReportParameters()
        {
            reportParameters.SelectedDevice = radioButtonForSelectedObject.Checked ? true : false;
            reportParameters.SelectedDevices = radioButtonForListDevice.Checked ? true : false;
            reportParameters.SelectedDevicesDifferentFiles = radioButtonForListDevicesDifferentFiles.Checked ? true : false;
            reportParameters.SelectedStoreObjects = radioButtonForListPKI.Checked ? true : false;
            reportParameters.SelectedNomObjects = radioButtonForListNomenclature.Checked ? true : false;
            reportParameters.SelectedDevicesWithKoef = radioButtonForListDeviceKoef.Checked ? true : false;
            reportParameters.ObjectsReferenceFailure = radioButtonForOIFailureByPeriod.Checked ? true : false;

            if (reportParameters.SelectedDevicesDifferentFiles)
            {
                var selectFolderDialog = new SelectFolderDialog();
                if (selectFolderDialog.ShowDialog(this) == DialogOpenResult.Ok)
                {
                    reportParameters.SelectedPath = selectFolderDialog.SelectedPath;
                }
            }

            var nomReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomReferenceInfo.CreateReference();

            var filter = new Filter(nomReferenceInfo);
            filter.Terms.AddTerm(reference.ParameterGroup[SystemParameterType.ObjectId], ComparisonOperator.Equal, _context.ObjectsInfo[0].ObjectId);

            NomenclatureObject findedObj = (NomenclatureObject)reference.Find(filter).FirstOrDefault();
            reportParameters.Device = findedObj;

            //Загружаем и проверяем объекты из справочника Отказы КРЭ и ПКИ, если выбран соответсвующий radiobutton
            if (radioButtonForOIFailureByPeriod.Checked)
            {
                reportParameters.BeginPeriod = dateTimePickerBegin.Value.Date;
                reportParameters.EndPeriod = dateTimePickerEnd.Value.Date;
                return CheckAndLoadObjectFailure();
            }

            return true;
        }

        public List<StoreObject> storeObjects = new List<StoreObject>();

        //Загрузка файла с объектами ПКИ Номенклатуры склада
        private void buttonLoadFile1_Click(object sender, EventArgs e)
        {
            buttonMakeReport.Enabled = false;
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference("Номенклатура складов");
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
                textBoxFileNameListPKI.Text = openFileDialog1.FileName;
                if (!textBoxFileNameListPKI.Text.ToLower().Contains(".xls"))
                    return;
    
                int errorID = 0;
                int numberStr = 0;

                try
                {              
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialog1.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    var countRecord = 0;
                    for (int i = 2; xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        countRecord++;
                    }

                    tsProgressBar.Maximum = countRecord;
                    tsProgressBar.Minimum = 0;
                    tsProgressBar.Step = 1;

                    for (int i = 2; xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        infoLabel.Text = "Количество считанных объектов " + (i+1).ToString() + " из " + countRecord;
                        dateTimeLabel.Text = DateTime.Now.ToLongTimeString();
                        tsProgressBar.PerformStep();
                        statusStrip.Update();
                        numberStr = i;
                        errorID = 3;
                        var name = xls[2, i].Text.Trim();
                        errorID = 4;
                        int count = Convert.ToInt32(xls[4, i].Text.Trim());
                        errorID = 5;
                        string measureUnit = xls[3, i].Text.Trim();

                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Наименование", ComparisonOperator.Equal, name);
                        var refObject = reference.Find(filter).FirstOrDefault();

                        if (refObject != null)
                        {
                            var newObj = new StoreObject(name, measureUnit, count, refObject);
                            reportParameters.StoreObjects.Add(newObj);
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(name + " строка № " + i);
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                        Clipboard.SetText(notFoundObjToStr);
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr +  "\r\n" +openFileDialog1.FileName + "\r\nКод ошибки:" + errorID+ "\r\n" +  ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                buttonMakeReport.Enabled = true;

            }
            else
                return;
        }


        private void radioButtonForListPKI_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForListPKI.Checked)
            {
                textBoxFileNameListPKI.Enabled = true;
                buttonLoadFileListPKI.Enabled = true;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;

                labelBegin.Enabled = false;
                labelEnd.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBegin.Enabled = false;
            }

        }

        private void radioButtonForSelectedObject_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForSelectedObject.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;

                labelBegin.Enabled = false;
                labelEnd.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBegin.Enabled = false;
            }

        }

        private void radioButtonForListDevice_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForListDevice.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListDevice.Enabled = true;
                buttonLoadFileListDevices.Enabled = true;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;

                labelBegin.Enabled = false;
                labelEnd.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBegin.Enabled = false;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForListNomenclature.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = true;
                buttonLoadFileListNomenclature.Enabled = true;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;
            }
        }

        private void radioButtonForListDeviceKoef_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForListDeviceKoef.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = true;
                buttonLoadFileListDevicesKoef.Enabled = true;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;

                labelBegin.Enabled = false;
                labelEnd.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBegin.Enabled = false;
            }
        }

        private void buttonLoadFileListDevices_Click(object sender, EventArgs e)
        {
            buttonMakeReport.Enabled = false;
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();

                if (radioButtonForListDevice.Checked)
                {
                    textBoxFileNameListDevice.Text = openFileDialog1.FileName;
                    if (!textBoxFileNameListDevice.Text.ToLower().Contains(".xls"))
                        return;
                }


                if (radioButtonForListDevicesDifferentFiles.Checked)
                {
                    textBoxFileNameListDevicesDifferentFiles.Text = openFileDialog1.FileName;
                    if (!textBoxFileNameListDevicesDifferentFiles.Text.ToLower().Contains(".xls"))
                        return;
                }

                int errorID = 0;
                int numberStr = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialog1.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    var countRecord = 0;
                    for (int i = 2; xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        countRecord++;
                    }

                    tsProgressBar.Maximum = countRecord;
                    tsProgressBar.Minimum = 0;
                    tsProgressBar.Step = 1;

                    for (int i = 2; xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        infoLabel.Text = "Количество считанных объектов " + (i+1).ToString() + " из " + countRecord;
                        dateTimeLabel.Text = DateTime.Now.ToLongTimeString();
                        tsProgressBar.PerformStep();
                        statusStrip.Update();

                        numberStr = i;
                        errorID = 3;
                        var denotation = xls[2, i].Text.Trim();
                        int count = Convert.ToInt32(xls[3, i].Text.Trim());

                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                        var refObject = reference.Find(filter).FirstOrDefault();

                        if (refObject != null)
                        {
                            if (!reportParameters.Devices.Keys.Contains(refObject))
                            {
                                reportParameters.Devices.Add((NomenclatureObject)refObject, count);
                            }
                            else
                            {
                                errorID = 8;
                                errorXlsObjects.Add(denotation + " строка № " + i + " объект с таким Обозначением уже добавлен");
                            }
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(denotation + " строка № " + i);
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                        Clipboard.SetText(notFoundObjToStr);
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr + "\r\n" + openFileDialog1.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                buttonMakeReport.Enabled = true;

            }
            else
                return;
        }

        private void buttonLoadFileListNomenclature_Click(object sender, EventArgs e)
        {
            buttonMakeReport.Enabled = false;
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
                textBoxFileNameListNomenclature.Text = openFileDialog1.FileName;
                if (!textBoxFileNameListNomenclature.Text.ToLower().Contains(".xls"))
                    return;

                int errorID = 0;
                int numberStr = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialog1.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    var countRecord = 0;

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty || xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        countRecord++;
                    }

                    tsProgressBar.Maximum = countRecord;
                    tsProgressBar.Minimum = 0;
                    tsProgressBar.Step = 1;

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty || xls[2, i].Text.Trim() != string.Empty; i++)
                    {
                        infoLabel.Text =  "Количество считанных объектов " + (i-1).ToString() + " из " + countRecord;
                        dateTimeLabel.Text = DateTime.Now.ToLongTimeString();
                        tsProgressBar.PerformStep();
                        statusStrip.Update();
                        numberStr = i;
                        errorID = 3;
                        var name = xls[2, i].Text.Trim();
                        var denotation = xls[1, i].Text.Trim();
                        var count = Convert.ToDouble(xls[3, i].Text.Trim());
                        if (Math.Abs(Math.Round(count) - count) > 0)
                        {
                            errorXlsObjects.Add(name + " " + denotation + " строка № " + i + " объект с некорректным количеством " + count + " был округлен до " + Math.Ceiling(count));
                        }
                        

                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Наименование", ComparisonOperator.Equal, name);
                        filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                        var refObject = (NomenclatureObject)reference.Find(filter).FirstOrDefault();

                        if (refObject != null)
                        {
                            if (reportParameters.NomObjects.Count(t => t.NomObj.SystemFields.Id == refObject.SystemFields.Id) == 0)
                            {
                                try
                                {
                                    var nomObject = new NomObject(refObject);
                                    try
                                    {
                                        nomObject.Count = Math.Ceiling(count);
                                        reportParameters.NomObjects.Add(nomObject);
                                    }
                                    catch
                                    {
                                        errorXlsObjects.Add(name + " строка № " + i + " Некорректное количество - " + count );
                                    }
                                }
                                catch
                                {
                                    errorXlsObjects.Add(name + " строка № " + i + " Нет связи с объектом склада");
                                }
                               
                            }
                            else
                            {
                                errorID = 8;
                                errorXlsObjects.Add(name + " строка № " + i + " объект с таким Наименованием уже добавлен");
                            }
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(name + " строка № " + i);
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                        Clipboard.SetText(notFoundObjToStr);
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr + "\r\n" + openFileDialog1.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                buttonMakeReport.Enabled = true;

            }
            else
                return;
        }

        private void buttonLoadFileListDevicesKoef_Click(object sender, EventArgs e)
        {
            buttonMakeReport.Enabled = false;
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
                textBoxFileNameDevicesKoef.Text = openFileDialog1.FileName;
                if (!textBoxFileNameDevicesKoef.Text.ToLower().Contains(".xls"))
                    return;

                int errorID = 0;
                int numberStr = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialog1.FileName, false);
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    var countRecord = 0;
                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        countRecord++;
                    }

                    tsProgressBar.Maximum = countRecord;
                    tsProgressBar.Minimum = 0;
                    tsProgressBar.Step = 1;

                    for (int i = 2; xls[1, i].Text.Trim() != string.Empty; i++)
                    {
                        infoLabel.Text = "Количество считанных объектов " + (i+1).ToString() + " из " + countRecord;
                        dateTimeLabel.Text = DateTime.Now.ToLongTimeString();
                        tsProgressBar.PerformStep();
                        statusStrip.Update();

                        numberStr = i;
                        errorID = 3;
                        var denotation = xls[1, i].Text.Trim();
                        int count = Convert.ToInt32(xls[2, i].Text.Trim());
                        errorID = 39;
                        double koef = Convert.ToDouble(xls[3, i].Text.Trim());
                        errorID = 33;

                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                        var refObject = (NomenclatureObject)reference.Find(filter).FirstOrDefault();
                       
                        if (refObject != null)
                        {
                            errorID = 4;
                            if (reportParameters.DevicesWithKoef.Count(t => t.NomObj.SystemFields.Id == refObject.SystemFields.Id) == 0)
                            {
                                errorID = 5;
                                reportParameters.DevicesWithKoef.Add(new DeviceWithKoef(refObject, count, koef));
                            }
                            else
                            {
                                errorID = 8;
                                errorXlsObjects.Add(denotation + " строка № " + i + " объект с таким Наименованием уже добавлен");
                            }
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(denotation + " строка № " + i + " не найден в Номенклатуре");
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                        Clipboard.SetText(notFoundObjToStr);
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr + "\r\n" + openFileDialog1.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                buttonMakeReport.Enabled = true;

            }
            else
                return;
        }

        private void radioButtonForListDevicesDifferentFiles_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForListDevicesDifferentFiles.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListDevicesDifferentFiles.Enabled = true;
                buttonLoadListDevicesDifferentFiles.Enabled = true;

                labelBegin.Enabled = false;
                labelEnd.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                dateTimePickerBegin.Enabled = false;
            }
        }

    /*    public List<NomObject> ReadData(NomenclatureObject nomObject)
        {
            List<NomObject> objects = new List<NomObject>();

            var nomReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            var reference = nomReferenceInfo.CreateReference();

            var hierarchyLinks = nomObject.Children.RecursiveLoadHierarchyLinks();
            var hierarchy = new Dictionary<int, Dictionary<Guid, ComplexHierarchyLink>>();
            var objectsDictionary = new Dictionary<int, NomenclatureObject>();
            objectsDictionary[nomObject.SystemFields.Id] = nomObject;

            foreach (var hierarchyLink in hierarchyLinks)
            {
                var nhl = (NomenclatureHierarchyLink)hierarchyLink;
                objectsDictionary[hierarchyLink.ChildObjectId] = (NomenclatureObject)hierarchyLink.ChildObject;

                Dictionary<Guid, ComplexHierarchyLink> parentHierarchy = null;
                if (!hierarchy.TryGetValue(hierarchyLink.ParentObjectId, out parentHierarchy))
                {
                    parentHierarchy = new Dictionary<Guid, ComplexHierarchyLink>();
                    hierarchy[hierarchyLink.ParentObjectId] = parentHierarchy;
                }
                parentHierarchy[hierarchyLink.Guid] = hierarchyLink;
            }

            var objectTypes = new List<Guid>();
            objectTypes.Add(Guids.NomenclatureTypes.OtherItem);

            foreach (var nomObj in objectsDictionary.Values.Where(t => objectTypes.Any(t1 => t.Class.IsInherit(t1)))
                .OrderBy(t => t.Denotation.GetString()).ThenBy(t => t.Name.GetString()))
            {
                var doc = new NomObject(nomObj);
                objects.Add(doc);
            }

            List<NomObject> objectsWithCount = new List<NomObject>();

            foreach (var obj in objects)
            {
                var newObj = obj;
                objectsWithCount.Add(newObj);
            }

            return objectsWithCount;
        }*/

        private void buttonLoadListDevicesDifferentFiles_Click(object sender, EventArgs e)
        {
            buttonMakeReport.Enabled = false;
            openFileDialog1.Filter = "(Excel-файлы)|*.xls;*.xlsx";
            List<string> errorXlsObjects = new List<string>();

            var nomenclatureReferenceInfo = ReferenceCatalog.FindReference(SpecialReference.Nomenclature);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var xls = new Xls();
  

                if (radioButtonForListDevice.Checked)
                {
                    textBoxFileNameListDevice.Text = openFileDialog1.FileName;
                    if (!textBoxFileNameListDevice.Text.ToLower().Contains(".xls"))
                        return;
                }


                if (radioButtonForListDevicesDifferentFiles.Checked)
                {
                    textBoxFileNameListDevicesDifferentFiles.Text = openFileDialog1.FileName;
                    if (!textBoxFileNameListDevicesDifferentFiles.Text.ToLower().Contains(".xls"))
                        return;
                }

                int errorID = 0;
                int numberStr = 0;

                try
                {
                    xls.Application.DisplayAlerts = false;
                    xls.Open(openFileDialog1.FileName, false);
                    xls.SelectWorksheet("Структура ОКА-МБ(МП)");
                    errorID = 1;

                    var reference = nomenclatureReferenceInfo.CreateReference();
                    errorID = 2;

                    var countRecord = 1053;


                    tsProgressBar.Maximum = countRecord;
                    tsProgressBar.Minimum = 0;
                    tsProgressBar.Step = 1;

                    for (int i = 3; i < 1056; i++)
                    {
                        infoLabel.Text = "Количество считанных объектов " + (i + 1).ToString() + " из " + countRecord;
                        dateTimeLabel.Text = DateTime.Now.ToLongTimeString();
                        tsProgressBar.PerformStep();
                        statusStrip.Update();

                        numberStr = i;
                        errorID = 3;
                        string denotation = xls[10, i].Text.Trim();

                        if (denotation == string.Empty)
                            continue;

                        int count = 1;
                        try
                        {
                            count = Convert.ToInt32(xls[5, i].Text.Trim());
                        }
                        catch
                        {
                            errorXlsObjects.Add(i.ToString() + " строка - некорректное количество");
                        }
                        var filter = new Filter(nomenclatureReferenceInfo);
                        filter.Terms.AddTerm("Обозначение", ComparisonOperator.Equal, denotation);
                        var refObject = reference.Find(filter).FirstOrDefault();



                        if (refObject != null)
                        {
                            if (!reportParameters.Devices.Keys.Contains(refObject))
                            {
                                reportParameters.Devices.Add((NomenclatureObject)refObject, 1);
                               
                            }
                            else
                            {
                               // errorID = 8;
                              //  errorXlsObjects.Add(i.ToString() + " строка " + denotation  + " объект с таким Обозначением уже добавлен");
                            }

                            var deviceByRow = new DeviceWithCountByRow((NomenclatureObject)refObject, count, i);
                            reportParameters.DevicesByRow.Add(deviceByRow);
                        }
                        else
                        {
                            errorID = 9;
                            errorXlsObjects.Add(i.ToString() + " строка - " + denotation + " объекта не существует ");
                        }
                    }

                    if (errorXlsObjects.Count > 0)
                    {
                        var notFoundObjToStr = string.Join("\r\n", errorXlsObjects);
                        Clipboard.SetText(notFoundObjToStr);
                        MessageBox.Show("1. Отчет будет неполным, так как не найдены следущие объекты:\r\n" + string.Join("\r\n", errorXlsObjects), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("2. Внимание ошибка! Сформированный отчет будет не корректен. Обратитесь в отдел 911. Возникла ошибка при работе с файлом в строке " + numberStr + "\r\n" + openFileDialog1.FileName + "\r\nКод ошибки:" + errorID + "\r\n" + ex, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    xls.Quit(true);
                    xls.Close();
                }
                buttonMakeReport.Enabled = true;

              //  var allcountPKI = reportParameters.Devices.Keys.Sum(t => ReadData(t).Count);

              //  MessageBox.Show("Всего найдено " + reportParameters.Devices.Count() + " объектов " + allcountPKI);

            }
            else
                return;
        }

        private void radioButtonForOIFailureByPeriod_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonForOIFailureByPeriod.Checked)
            {
                textBoxFileNameListPKI.Enabled = false;
                buttonLoadFileListPKI.Enabled = false;

                textBoxFileNameListDevice.Enabled = false;
                buttonLoadFileListDevices.Enabled = false;

                textBoxFileNameListNomenclature.Enabled = false;
                buttonLoadFileListNomenclature.Enabled = false;

                textBoxFileNameDevicesKoef.Enabled = false;
                buttonLoadFileListDevicesKoef.Enabled = false;

                textBoxFileNameListDevicesDifferentFiles.Enabled = false;
                buttonLoadListDevicesDifferentFiles.Enabled = false;

                labelBegin.Enabled = true;
                labelEnd.Enabled = true;
                dateTimePickerEnd.Enabled = true;
                dateTimePickerBegin.Enabled = true;
            }
        }
    }
}
