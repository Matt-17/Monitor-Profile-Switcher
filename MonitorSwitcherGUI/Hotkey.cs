using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MonitorSwitcherGUI
{
    [StructLayout(LayoutKind.Sequential)]
    public class Hotkey
    {
        public bool Ctrl;
        public bool Alt;
        public bool Shift;
        public bool RemoveKey;
        public Keys Key;
        public string profileName;

        public HotkeyCtrl hotkeyCtrl;

        public Hotkey()
        {
            hotkeyCtrl = new HotkeyCtrl();
            RemoveKey = false;
        }

        public void RegisterHotkey(MonitorSwitcherGUI parent)
        {
            hotkeyCtrl.Alt = Alt;
            hotkeyCtrl.Shift = Shift;
            hotkeyCtrl.Control = Ctrl;
            hotkeyCtrl.KeyCode = Key;
            hotkeyCtrl.Pressed += new HandledEventHandler(parent.KeyHook_KeyUp);

            if (!hotkeyCtrl.GetCanRegister(parent))
            {
                // something went wrong, ignore for nw
            }
            else
            {
                hotkeyCtrl.Register(parent);
            }
        }

        public void UnregisterHotkey()
        {
            if (hotkeyCtrl.Registered)
            {
                hotkeyCtrl.Unregister();
            }
        }

        public void AssignFromKeyEventArgs(KeyEventArgs keyEvents)
        {
            Ctrl = keyEvents.Control;
            Alt = keyEvents.Alt;
            Shift = keyEvents.Shift;
            Key = keyEvents.KeyCode;
        }

        public override string ToString()
        {
            List<string> keys = new List<string>();

            if (Ctrl == true)
            {
                keys.Add("CTRL");
            }

            if (Alt == true)
            {
                keys.Add("ALT");
            }

            if (Shift == true)
            {
                keys.Add("SHIFT");
            }

            switch (Key)
            {
                case Keys.ControlKey:
                case Keys.Alt:
                case Keys.ShiftKey:
                case Keys.Menu:
                    break;
                default:
                    keys.Add(Key.ToString()
                        .Replace("Oem", string.Empty)
                    );
                    break;
            }

            return string.Join(" + ", keys);
        }
    }
}