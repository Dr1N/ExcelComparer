using System;
using System.ComponentModel;
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

        protected override void OnClosing(CancelEventArgs e)
        {
            Hide();
            e.Cancel = true;
            base.OnClosing(e);
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tbLog.Text = "";
        }

        public void Write(string message)
        {
            if (this.tbLog.InvokeRequired)
            {
                Action<string> action = new Action<string>(Write);
                this.tbLog.BeginInvoke(action, new object[] { message });
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