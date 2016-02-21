using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelComparer
{
    [Flags]
    public enum LOG_MODE
    {
        NONE = 0x0,
        DEBUG = 0x1,
        CONSOLE = 0x2,
        FILE = 0x4,
        WINDOW = 0x8
    }

    static class LogWriter
    {
        private static string logFile = "log.txt";
        private static LogForm logForm;

        public static LOG_MODE Mode { get; set; }

        static LogWriter()
        {
            Mode = LOG_MODE.DEBUG | LOG_MODE.WINDOW;
            logForm = LogForm.Instance;
        }
        
        public static void Write(string message)
        {
            bool isDebug = (Mode & LOG_MODE.DEBUG) == LOG_MODE.DEBUG;
            bool isConsole = (Mode & LOG_MODE.CONSOLE) == LOG_MODE.DEBUG;
            bool isFile = (Mode & LOG_MODE.FILE) == LOG_MODE.FILE;
            bool isWindow = (Mode & LOG_MODE.WINDOW) == LOG_MODE.WINDOW;

            DateTime dt = DateTime.Now;
            string dateTime = dt.ToShortDateString() + " " + dt.ToLongTimeString();

            if(isDebug)
            {
                System.Diagnostics.Debug.WriteLine(dateTime + "\t" + message);
            }
            if (isConsole)
            {
                Console.WriteLine(dateTime + "\t" + message);
            }
            if(isFile)
            {
                try
                {
                    File.AppendAllText(logFile, dateTime + "\t" + message);
                }
                catch(Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Не удалось записать в файл лога " + e.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }
            }
            if (isWindow)
            {
                logForm.Write(dateTime + "\t" + message);
            }  
        }
    }
}