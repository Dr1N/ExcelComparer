using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelComparer
{
    public partial class LogForm : Form
    {
        private static LogForm instance = null;
        public static LogForm Instance
        {
            get
            {
                return instance;
            }
        }

        static LogForm()
        {
            instance = new LogForm();
        }

        private LogForm()
        {
            InitializeComponent();
        }

        public void Write(string message)
        {
            if (this.tbLog.InvokeRequired)
            {
                Action<string> action = new Action<string>(Write);
                this.tbLog.Invoke(action, new object[] { message });
            }
            else
            {
                if (tbLog.Text.Length + message.Length > tbLog.MaxLength)
                {
                    int cr = tbLog.Text.IndexOf("rn");
                    if (cr > 0)
                    {
                        tbLog.Select(0, cr + 1);
                        tbLog.SelectedText = string.Empty;
                    }
                    else
                    {
                        tbLog.Select(0, message.Length);
                    }
                }
                tbLog.AppendText(message + Environment.NewLine);
                tbLog.SelectionStart = tbLog.Text.Length;
                tbLog.ScrollToCaret();
            }
        }
    }
}