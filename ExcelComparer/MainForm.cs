using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ExcelComparer
{
    public partial class MainForm : Form
    {
        #region FIELDS

        private readonly string fileFilter = "Excel(XLSX)|*.xlsx|Excel(XLSB)|*.xlsb|Excel(XLSM)|*.xlsm|Excel(XLS)|*.xls|Все файлы|*.*";
        private readonly List<string> fileExtentions = new List<string>() { ".xlsx", ".xlsb", ".xlsm", ".xls" };

        private string file;
        private string directory;
        private string[] files;

        private Thread workThread = null;

        #endregion

        public MainForm()
        {
            InitializeComponent();

            this.directory = @"d:\My Coding\KWORK\Files\Dir\";
            this.file = @"d:\My Coding\KWORK\Files\file1.xlsx";

            this.lblDataBaseFile.Text = this.file;
            this.lblDirectoryPath.Text = this.directory;
        }

        #region CONTROL EVENTS

        private void btnStart_Click(object sender, EventArgs e)
        {
            if(LogForm.Instance.Visible == false)
            { 
                LogForm.Instance.Show();
            }
            try
            {
                GetFileList();
                if (!IsValidSettings())
                {
                    string message = "Неверные настройки программы. Смотри логи работы программы";
                    LogWriter.Write(message);
                    MessageBox.Show(message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelDbComparer comparer = new ExcelDbComparer(this.file, this.files);
                this.workThread = new Thread(comparer.RunWork) { IsBackground = true };
                this.workThread.Start(this);
            }
            catch (Exception ex)
            {
                LogWriter.Write("Обработка прекращена! Причина: " + ex.Message);
                MessageBox.Show("В процессе обработки возникли ошибки. Смотри логи работы!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            try
            {
                if (this.workThread != null)
                {
                    this.workThread.Abort();
                    this.workThread = null;
                }
            }
            catch (Exception) { e.Cancel = true; }
            base.OnClosing(e);
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            try
            {
                this.file = GetDataBaseFileName();
                lblDataBaseFile.Text = this.file;
            }
            catch (ComparerException ceex)
            {
                LogWriter.Write(ceex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnSelectDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                this.directory = GetDataBasesDirectory();
                lblDirectoryPath.Text = directory;
            }
            catch (ComparerException ceex)
            {
                LogWriter.Write(ceex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new AboutBox().ShowDialog();
        }

        #endregion

        #region METHODS

        public void StartWork()
        {
            if (this.InvokeRequired)
            {
                Action action = new Action(StartWork);
                this.Invoke(action, new object[0]);
            }
            else
            {
                btnStart.Enabled = false;
                btnCancel.Enabled = false;
                btnSelectDirectory.Enabled = false;
                btnSelectFile.Enabled = false;
                linkLabel1.Enabled = false;
                Text = "Excel Comparer [Идёт обработка...]";
            }
        }

        public void EndWork()
        {
            if (this.InvokeRequired)
            {
                Action action = new Action(EndWork);
                this.Invoke(action, new object[0]);
            }
            else
            {
                btnStart.Enabled = true;
                btnCancel.Enabled = true;
                btnSelectDirectory.Enabled = true;
                btnSelectFile.Enabled = true;
                linkLabel1.Enabled = true;
                Text = "Excel Comparer [Обработка завершена]";
            }
        }

        private string GetDataBaseFileName()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = false;
                ofd.Title = "Выберете файл базы";
                ofd.Filter = fileFilter;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string ext = Path.GetExtension(ofd.FileName);
                    if (!fileExtentions.Contains(ext))
                    {
                        throw new Exception("Данный тип файла не поддерживается");
                    }
                    return ofd.FileName;
                }
                throw new ComparerException("Пользователь не выбрал файл");
            }
        }

        private string GetDataBasesDirectory()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                if(fbd.ShowDialog() == DialogResult.OK)
                {
                    return fbd.SelectedPath;
                }
            }
            throw new ComparerException("Пользователь не выбрал каталог");
        }

        private bool IsValidSettings()
        {
            bool result = true;
            string filePath = lblDataBaseFile.Text;
            string directoryPath = lblDirectoryPath.Text;

            if (!File.Exists(filePath))
            {
                LogWriter.Write("Указанный файл не существует");
                result = false;
            }

            if (!Directory.Exists(directoryPath))
            {
                LogWriter.Write("Указанный каталог не существует");
                result = false;
            }

            string fileDirectory = Path.GetDirectoryName(filePath);
            if (fileDirectory.StartsWith(directoryPath))
            {
                LogWriter.Write("Файл не должен находится в указанной директории");
                result = false;
            }

            if (this.files.Length == 0)
            {
                LogWriter.Write("В директори нет файлов для анализа");
                result = false;
            }

            return result;
        }

        private void GetFileList()
        {
            string[] files = Directory.GetFiles(this.directory, "*.*", SearchOption.AllDirectories);
            this.files = (from f 
                          in files
                          where this.fileExtentions.Contains(Path.GetExtension(f)) && !Path.GetFileName(f).StartsWith("~$")
                          select f).ToArray();
        }

        #endregion
    }
}