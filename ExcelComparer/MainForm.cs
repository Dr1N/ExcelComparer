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

        private readonly string fileFilter = "Excel(XLSX)|*.xlsx|Excel(XLSM)|*.xlsm|Excel(XLS)|*.xls|Excel(CSV)|*.csv|Все файлы|*.*";
        private readonly List<string> fileFormats = new List<string>() { ".xlsx", ".xlsm", ".xls", ".csv" };

        #endregion

        public MainForm()
        {
            InitializeComponent();
        }

        #region CONTROL EVENTS

        private void btnStart_Click(object sender, EventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            try
            {
                string fileName = GetOpenableFileName();
                if (sender == btnFirst)
                {
                    lblFirstFile.Text = fileName;
                }
                else if(sender == btnSecond)
                {
                    lblSecondFile.Text = fileName;
                }
            }
            catch (ComparerException ceex)
            {
                Debug.WriteLine(ceex, "Выбор файла");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region METHODS
        
        private string GetOpenableFileName()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Multiselect = false;
                ofd.Title = "Выберете файл базы";
                ofd.Filter = fileFilter;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string ext = Path.GetExtension(ofd.FileName);
                    if (!fileFormats.Contains(ext))
                    {
                        throw new Exception("Данный тип файла не поддерживается");
                    }
                    return ofd.FileName;
                }
                throw new ComparerException("Пользователь не выбрал файл");
            }
        }

        #endregion
    }
}