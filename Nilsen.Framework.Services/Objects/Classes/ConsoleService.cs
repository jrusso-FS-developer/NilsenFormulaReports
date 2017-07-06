using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Nilsen.Framework.Services.Objects.Classes
{
    public class ConsoleService
    {
        private TextBox ConsoleWindow;
        private Button BtnProcess;

        delegate void UpdateConsoleTextCallback(string sText, bool clear);
        delegate void DisableButtonCallback(bool bEnabled);

        public ConsoleService(TextBox consoleWindow, Button btnProcess)
        {
            ConsoleWindow = consoleWindow;
            BtnProcess = btnProcess;
        }

        public void ToggleProcessButton(bool bEnabled)
        {
            if (BtnProcess.InvokeRequired)
            {
                var d = new DisableButtonCallback(ToggleProcessButton);
                BtnProcess.Invoke(d, new object[] { bEnabled });
            }
            else
            {
                BtnProcess.Enabled = bEnabled;
            }
        }

        public void UpdateConsoleText(String sText, bool clear)
        {
            if (ConsoleWindow.InvokeRequired)
            {
                var d = new UpdateConsoleTextCallback(UpdateConsoleText);
                ConsoleWindow.Invoke(d, new object[] {sText, clear});
            }
            else
            {
                ConsoleWindow.Text = clear ? string.Empty : ConsoleWindow.Text;

                if (ConsoleWindow.Text.Equals(string.Empty))
                {
                    ConsoleWindow.Text = string.Format("{0} >> {1}", ConsoleWindow.Text, sText);
                }
                else
                {
                    ConsoleWindow.Text = string.Format("{0}\r\n\r\n >> {1}", ConsoleWindow.Text, sText);
                }
                ConsoleWindow.SelectionStart = ConsoleWindow.Text.Length;
                ConsoleWindow.ScrollToCaret();
                ConsoleWindow.Refresh();
            }
        }
    }
}
