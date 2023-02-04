using System;
using System.Windows.Forms;

namespace MonitorSwitcherGUI
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            // Parse command line
            string customSettingsDirectory = "";
            foreach (string iArg in args)
            {
                string[] argElements = iArg.Split(new char[] { ':' }, 2);

                switch (argElements[0].ToLower())
                {
                    case "-settings":
                        customSettingsDirectory = argElements[1];
                        break;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MonitorSwitcherGUI(customSettingsDirectory));
        }
    }
}