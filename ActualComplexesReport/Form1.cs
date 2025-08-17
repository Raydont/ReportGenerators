using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ActualComplexesReport
{
    public class SaveFileForm : Form
    {
        #region Форма

        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.objectLabel = new System.Windows.Forms.Label();
            this.objectTextBox = new System.Windows.Forms.TextBox();
            this.selectFileLabel = new System.Windows.Forms.Label();
            this.selectFileTextBox = new System.Windows.Forms.TextBox();
            this.selectFileButton = new System.Windows.Forms.Button();
            this.saveButton = new System.Windows.Forms.Button();
            this.serviceTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // objectLabel
            // 
            this.objectLabel.AutoSize = true;
            this.objectLabel.Location = new System.Drawing.Point(13, 24);
            this.objectLabel.Name = "objectLabel";
            this.objectLabel.Size = new System.Drawing.Size(48, 13);
            this.objectLabel.TabIndex = 0;
            this.objectLabel.Text = "Объект:";
            // 
            // objectTextBox
            // 
            this.objectTextBox.Location = new System.Drawing.Point(64, 21);
            this.objectTextBox.Name = "objectTextBox";
            this.objectTextBox.ReadOnly = true;
            this.objectTextBox.Size = new System.Drawing.Size(334, 20);
            this.objectTextBox.TabIndex = 1;
            // 
            // selectFileLabel
            // 
            this.selectFileLabel.AutoSize = true;
            this.selectFileLabel.Location = new System.Drawing.Point(12, 85);
            this.selectFileLabel.Name = "selectFileLabel";
            this.selectFileLabel.Size = new System.Drawing.Size(98, 13);
            this.selectFileLabel.TabIndex = 2;
            this.selectFileLabel.Text = "Выбранный файл:";
            // 
            // selectFileTextBox
            // 
            this.selectFileTextBox.Location = new System.Drawing.Point(15, 101);
            this.selectFileTextBox.Name = "selectFileTextBox";
            this.selectFileTextBox.ReadOnly = true;
            this.selectFileTextBox.Size = new System.Drawing.Size(350, 20);
            this.selectFileTextBox.TabIndex = 3;
            // 
            // selectFileButton
            // 
            this.selectFileButton.Location = new System.Drawing.Point(371, 99);
            this.selectFileButton.Name = "selectFileButton";
            this.selectFileButton.Size = new System.Drawing.Size(28, 23);
            this.selectFileButton.TabIndex = 4;
            this.selectFileButton.Text = "...";
            this.selectFileButton.UseVisualStyleBackColor = true;
            this.selectFileButton.Click += new System.EventHandler(this.selectFileButton_Click);
            // 
            // saveButton
            // 
            this.saveButton.Location = new System.Drawing.Point(103, 127);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(228, 23);
            this.saveButton.TabIndex = 5;
            this.saveButton.Text = "Сохранить";
            this.saveButton.UseVisualStyleBackColor = true;
            this.saveButton.Click += new System.EventHandler(this.saveButton_Click);
            // 
            // serviceTextBox
            // 
            this.serviceTextBox.Location = new System.Drawing.Point(204, 56);
            this.serviceTextBox.Name = "serviceTextBox";
            this.serviceTextBox.Size = new System.Drawing.Size(192, 20);
            this.serviceTextBox.TabIndex = 6;
            // 
            // SaveFileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(410, 158);
            this.Controls.Add(this.serviceTextBox);
            this.Controls.Add(this.saveButton);
            this.Controls.Add(this.selectFileButton);
            this.Controls.Add(this.selectFileTextBox);
            this.Controls.Add(this.selectFileLabel);
            this.Controls.Add(this.objectTextBox);
            this.Controls.Add(this.objectLabel);
            this.Name = "SaveFileForm";
            this.Text = "Добавление файла структуры в хранилище КД";
            this.Load += new System.EventHandler(this.SaveFileForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label objectLabel;
        private System.Windows.Forms.TextBox objectTextBox;
        private System.Windows.Forms.Label orderLabel;
        private System.Windows.Forms.TextBox orderTextBox;
        private System.Windows.Forms.Label selectFileLabel;
        private System.Windows.Forms.TextBox selectFileTextBox;
        private System.Windows.Forms.Button selectFileButton;
        private System.Windows.Forms.Button saveButton;

        FileReference fileReference;

        public FileObject fileObject = null;
        public ReferenceObject objectToAttach;
        public int attachedFilesCount;
        public string newFileName = string.Empty;

        // Связь со справочником "Аппаратура"
        public static Guid complexLinkGuid = new Guid("f22aa93e-a7cc-4938-9d9b-3d9f90718697");
        // Наименование из справочника "Аппаратура"        
        public static Guid complexNameGuid = new Guid("7a70c6f4-f586-4739-b84a-cadef6454f49");
        // Обозначение из справочника "Аппаратура"                
        public static Guid complexDenotationGuid = new Guid("bc5f447c-8e7e-4cd1-b43c-da5920e9019a");
        // Номер заказа из справочника "Заказы"
        public static Guid orderNumberGuid = new Guid("027fcfd8-4bb6-4b09-9423-7f2ffc0ae0dc");
        // Связь со справочником "Файлы" справочника "Заказы"        
        public static Guid structureLinkGuid = new Guid("eb6b67f0-03ba-406a-88dd-bdf1d45cc864");

        public SaveFileForm()
        {
            InitializeComponent();
        }

        public SaveFileForm(ReferenceObject refObject)
        {
            InitializeComponent();
            objectToAttach = refObject;
        }


        private void SaveFileForm_Load(object sender, EventArgs e)
        {
            objectTextBox.Text = "[" + objectToAttach.Links.ToOne[complexLinkGuid].LinkedObject[complexNameGuid].Value.ToString() + "  "
                                 + objectToAttach.Links.ToOne[complexLinkGuid].LinkedObject[complexDenotationGuid].Value.ToString() + "]  заказ "
                                 + objectToAttach[orderNumberGuid].Value.ToString();

            // Сколько файлов присоединено к объекту
            var attachedFilesLink = objectToAttach.Links.ToMany[structureLinkGuid];
            attachedFilesCount = attachedFilesLink.Objects.Count();
        }

        private void selectFileButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Все разрешенные (zip,ZIP)|*.zip; |ZIP|*.zip";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                selectFileTextBox.Text = openFileDialog.FileName;
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            if (selectFileTextBox.Text.Trim() == string.Empty || !File.Exists(selectFileTextBox.Text))
            {
                MessageBox.Show("Файл по указанному пути не существует");
                return;
            }

            fileReference = new FileReference();

            TFlex.DOCs.Model.References.Files.FolderObject folderObject = null;

            try
            {
                folderObject = FindOrCreateFolder();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                DialogResult = DialogResult.Cancel;
                return;
            }

            FileInfo fileInfo = new FileInfo(selectFileTextBox.Text);
            string fileExtension = fileInfo.Extension;
            newFileName = "Структура " + objectToAttach.Links.ToOne[complexLinkGuid].LinkedObject[complexDenotationGuid].Value.ToString() + " заказ " + objectToAttach[orderNumberGuid].Value.ToString();

            if (attachedFilesCount != 0)
            {
                string fileVersion = "_B" + attachedFilesCount;             // если у объекта уже есть присоединенные файлы, то сформировать новую версию
                newFileName = newFileName + fileVersion + fileExtension;
            }
            else
            {
                newFileName = newFileName + fileExtension;
            }

            newFileName = fileNameCorrector(newFileName);

            var appDataFileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), newFileName);
            File.SetAttributes(selectFileTextBox.Text, FileAttributes.Normal);
            File.Copy(selectFileTextBox.Text, appDataFileName, true);

            // ищем, не присоединен ли уже файл с сформированным нами именем к объекту
            FileObject searchfileObject = null;
            ReferenceInfo fileReferenceInfo = ReferenceCatalog.FindReference(fileReference.Id);
            using (Filter filter = new Filter(fileReferenceInfo))
            {
                filter.Terms.AddTerm(fileReference.ParameterGroup.OneToOneParameters.FindByName("Наименование"), ComparisonOperator.Equal, newFileName); // Наименование файла
                searchfileObject = fileReference.Find(filter).FirstOrDefault() as FileObject;
            }

            if (searchfileObject != null)
            {
                DialogResult replaceAction = MessageBox.Show("Файл с таким именем " + newFileName + " уже существует в файловом Хранилище и присоединен к выбранному объекту.\n\n Заменить его выбранным Вами файлом?", "ДОКУМЕНТ УЖЕ СУЩЕСТВУЕТ !!!", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if (replaceAction == DialogResult.No)
                {
                    DialogResult = DialogResult.Cancel;
                    return;
                }
            }

            if (!AddFile(folderObject, appDataFileName))
                DialogResult = DialogResult.Cancel;
            else
                DialogResult = DialogResult.OK;
        }

        private bool AddFile(FolderObject folderObject, string appDataFileName)
        {
            fileObject = fileReference.AddFile(appDataFileName, folderObject);
            try
            {
                Desktop.CheckIn(fileObject, "Добавление файла", false);

                File.SetAttributes(appDataFileName, FileAttributes.Normal);
                File.Delete(appDataFileName);
            }
            catch
            {
                MessageBox.Show("Не удалось сохранить файл в файловом хранилище. Обратитесь к системному администратору.", "Ошибка сервера", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try
                {
                    Desktop.UndoCheckOut(new[] { fileObject });
                }
                catch { }
                return false;
            }

            return true;
        }

        // Замена некорректных символов в имени файла
        private string fileNameCorrector(string fileName)
        {

            foreach (char invalidChar in System.IO.Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(invalidChar, Convert.ToChar("_"));
            }

            return fileName;
        }

        // Поиск папки, в которую помещается файл
        private TFlex.DOCs.Model.References.Files.FolderObject FindOrCreateFolder()
        {
            string baseFolderPath = @"Файлы документов\Структуры аппаратур";

            TFlex.DOCs.Model.References.Files.FolderObject baseFolder = (TFlex.DOCs.Model.References.Files.FolderObject)fileReference.FindByPath(baseFolderPath);

            string orderFolderPath = "Аппаратура " + objectToAttach.Links.ToOne[complexLinkGuid].LinkedObject[complexDenotationGuid].Value.ToString() + " заказ " + objectToAttach[orderNumberGuid].Value.ToString();

            orderFolderPath = fileNameCorrector(orderFolderPath);
            baseFolderPath += "\\" + orderFolderPath;


            if ((TFlex.DOCs.Model.References.Files.FolderObject)fileReference.FindByPath(baseFolderPath) == null)
            {
                var folder = baseFolder.CreateFolder("", orderFolderPath);
                Desktop.CheckIn(folder, "Создание папки " + orderFolderPath, false);
            }

            return (TFlex.DOCs.Model.References.Files.FolderObject)fileReference.FindByPath(baseFolderPath);
        }

        private TextBox serviceTextBox;
    }
}
