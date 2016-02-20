using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelComparer
{
    public partial class MainForm : Form
    {
        #region FIELDS

        private readonly string fileFilter = "Excel(XLSX)|*.xlsx|Excel(XLSB)|*.xlsb|Excel(XLSM)|*.xlsm|Excel(XLS)|*.xls|Все файлы|*.*";
        private readonly List<string> fileExtentions = new List<string>() { ".xlsx", ".xlsb", ".xlsm", ".xls" };

        private string directory;
        private string[] files;
        private string file;

        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        #region CONTROL EVENTS

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (!IsValidPaths())
                {
                    MessageBox.Show("Есть ошибки в настройках программы. Смотри логи работы программы", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (GetFileList() <= 0)
                {
                    MessageBox.Show("Не удалось получить список файлов для анализа. Смотри логи работы программы", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                ExcelDbComparer comparer = new ExcelDbComparer(this.file, this.files);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
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

        #endregion

        #region METHODS

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

        private bool IsValidPaths()
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

            return result;
        }

        private int GetFileList()
        {
            try
            {
                string[] files = Directory.GetFiles(this.directory, "*.*", SearchOption.AllDirectories);
                this.files = (from f 
                              in files
                              where this.fileExtentions.Contains(Path.GetExtension(f)) && !Path.GetFileName(f).StartsWith("~$")
                              select f).ToArray();

                if (this.files.Length == 0)
                {
                    LogWriter.Write("В директории нет файлов для анализа");
                }
                LogWriter.Write("Найдено " + this.files.Length + " для анализа:");
                foreach (string f in files)
                {
                    LogWriter.Write(f);
                }
                return files.Length;
            }
            catch (Exception e)
            {
                LogWriter.Write("Не удалось получить список файлов" + Environment.NewLine + e.Message);
                return -1;
            }
        }

        #endregion

        private void btnTest_Click(object sender, EventArgs e)
        {
            ExcelDbComparer c = new ExcelDbComparer("test", new string[] { "test", "test", "test" });
            DataTable dt = c.ReadExcelFile(@"D:\My Coding\KWORK\ExcelComparer\Files\file1.xlsx");
            List<NameFromDataTable> names = c.ReadNamesFromData(dt);
        }
    }
}