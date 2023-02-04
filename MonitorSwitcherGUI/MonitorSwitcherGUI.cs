/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Xml.Serialization;

namespace MonitorSwitcherGUI
{
    public class MonitorSwitcherGUI : Form
    {
        private readonly NotifyIcon _trayIcon;
        private readonly ContextMenuStrip _trayMenu;
        private readonly string _settingsDirectory;
        private readonly string _settingsDirectoryProfiles;
        private readonly List<Hotkey> _hotkeys;

        // icons for tray menu
        private static readonly string[] Icons = { "MainIcon.ico", "DeleteProfile.ico", "Exit.ico", "Profile.ico", "SaveProfile.ico", "NewProfile.ico", "About.ico", "Hotkey.ico" };

        public MonitorSwitcherGUI(string customSettingsDirectory)
        {
            // Initialize settings directory
            _settingsDirectory = GetSettingsDirectory(customSettingsDirectory);
            _settingsDirectoryProfiles = GetSettingsProfielDirectotry(_settingsDirectory);

            if (!Directory.Exists(_settingsDirectory))
                Directory.CreateDirectory(_settingsDirectory);
            if (!Directory.Exists(_settingsDirectoryProfiles))
                Directory.CreateDirectory(_settingsDirectoryProfiles);

            // Initialize Hotkey list before loading settings
            _hotkeys = new List<Hotkey>();

            // Load all settings
            LoadSettings();

            // Refresh Hotkey Hooks
            KeyHooksRefresh();

            // Build up context menu
            _trayMenu = new ContextMenuStrip();
            _trayMenu.ImageList = new ImageList();

            // add icons to imagelist via foreach
            foreach (var icon in Icons)
            {
                _trayMenu.ImageList.Images.Add(new Icon(GetType(), $"Resources.{icon}"));
            }

            // add paypal png logo
            var myAssembly = Assembly.GetExecutingAssembly();
            var myStream = myAssembly.GetManifestResourceStream("MonitorSwitcherGUI.Resources.PayPal.png");
            _trayMenu.ImageList.Images.Add(Image.FromStream(myStream));

            // finally build tray menu
            BuildTrayMenu();

            // Create tray icon
            _trayIcon = new NotifyIcon();
            _trayIcon.Text = "Monitor Profile Switcher";
            _trayIcon.Icon = new Icon(GetType(), $"Resources.{Icons[0]}");
            _trayIcon.ContextMenuStrip = _trayMenu;
            _trayIcon.Visible = true;


            _trayIcon.MouseUp += OnTrayClick;
        }

        public static string GetSettingsDirectory(string customSettingsDirectory)
        {
            var dir = "";
            if (string.IsNullOrEmpty(customSettingsDirectory))
            {
                dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MonitorSwitcher");
            }
            else
            {
                dir = customSettingsDirectory;
            }
            return dir;
        }

        public static string GetSettingsProfielDirectotry(string settingsDirectory)
        {
            return Path.Combine(settingsDirectory, "Profiles");
        }

        private void KeyHooksRefresh()
        {
            var removeList = new List<Hotkey>();
            // check which hooks are still valid
            foreach (var hotkey in _hotkeys)
            {
                if (!File.Exists(ProfileFileFromName(hotkey.profileName)))
                {
                    hotkey.UnregisterHotkey();
                    removeList.Add(hotkey);
                }
            }
            if (removeList.Count > 0)
            {
                foreach (var hotkey in removeList)
                {
                    _hotkeys.Remove(hotkey);
                }
                removeList.Clear();
                SaveSettings();
            }

            // register the valid hooks
            foreach (var hotkey in _hotkeys)
            {
                hotkey.UnregisterHotkey();
                hotkey.RegisterHotkey(this);
            }
        }

        public void KeyHook_KeyUp(object sender, HandledEventArgs e)
        {
            var hotkeyCtrl = (sender as HotkeyCtrl);
            var hotkey = FindHotkey(hotkeyCtrl);
            LoadProfile(hotkey.profileName);
            e.Handled = true;
        }

        public void KeyHook_KeyDown(object sender, HandledEventArgs e)
        {
            e.Handled = true;
        }

        public void LoadSettings()
        {
            // Unregister and clear all existing hotkeys
            foreach (var hotkey in _hotkeys)
            {
                hotkey.UnregisterHotkey();
            }
            _hotkeys.Clear();

            // Loading the xml file
            if (!File.Exists(SettingsFileFromName("Hotkeys")))
                return;

            var readerHotkey = new XmlSerializer(typeof(Hotkey));

            try
            {
                var xml = XmlReader.Create(SettingsFileFromName("Hotkeys"));
                xml.Read();
                while (true)
                {
                    if ((xml.Name.CompareTo("Hotkey") == 0) && (xml.IsStartElement()))
                    {
                        var hotkey = (Hotkey)readerHotkey.Deserialize(xml);
                        _hotkeys.Add(hotkey);
                        continue;
                    }

                    if (!xml.Read())
                    {
                        break;
                    }
                }
                xml.Close();
            }
            catch (Exception)
            {
                // ignored
            }
        }

        public void SaveSettings()
        {
            var writerHotkey = new XmlSerializer(typeof(Hotkey));

            var xmlSettings = new XmlWriterSettings
            {
                CloseOutput = true
            };

            try
            {
                using (var fileStream = new FileStream(SettingsFileFromName("Hotkeys"), FileMode.Create))
                {
                    var xml = XmlWriter.Create(fileStream, xmlSettings);
                    xml.WriteStartDocument();
                    xml.WriteStartElement("hotkeys");
                    foreach (var hotkey in _hotkeys)
                    {
                        writerHotkey.Serialize(xml, hotkey);
                    }
                    xml.WriteEndElement();
                    xml.WriteEndDocument();
                    xml.Flush();
                    xml.Close();

                    fileStream.Close();
                }
            }
            catch
            {
            }
        }

        public Hotkey FindHotkey(HotkeyCtrl ctrl)
        {
            foreach (var hotkey in _hotkeys)
            {
                if (hotkey.hotkeyCtrl == ctrl)
                    return hotkey;
            }

            return null;
        }

        public Hotkey FindHotkey(string name)
        {
            foreach (var hotkey in _hotkeys)
            {
                if (hotkey.profileName.CompareTo(name) == 0)
                    return hotkey;
            }

            return null;
        }

        public void BuildTrayMenu()
        {
            ToolStripItem newMenuItem;

            _trayMenu.Items.Clear();

            _trayMenu.Items.Add("Load Profile").Enabled = false;
            _trayMenu.Items.Add("-");

            // Find all profile files
            var profiles = Directory.GetFiles(_settingsDirectoryProfiles, "*.xml");

            // Add to load menu
            foreach (var profile in profiles)
            {
                var itemCaption = Path.GetFileNameWithoutExtension(profile);
                newMenuItem = _trayMenu.Items.Add(itemCaption);
                newMenuItem.Click += OnMenuLoad;
                newMenuItem.ImageIndex = 3;
            }

            // Menu for saving items
            _trayMenu.Items.Add("-");
            var saveMenu = new ToolStripMenuItem("Save Profile");
            saveMenu.ImageIndex = 4;
            saveMenu.DropDown = new ToolStripDropDownMenu();
            saveMenu.DropDown.ImageList = _trayMenu.ImageList;
            _trayMenu.Items.Add(saveMenu);

            newMenuItem = saveMenu.DropDownItems.Add("New Profile...");
            newMenuItem.Click += OnMenuSaveAs;
            newMenuItem.ImageIndex = 5;
            saveMenu.DropDownItems.Add("-");

            // Menu for deleting items
            var deleteMenu = new ToolStripMenuItem("Delete Profile");
            deleteMenu.ImageIndex = 1;
            deleteMenu.DropDown = new ToolStripDropDownMenu();
            deleteMenu.DropDown.ImageList = _trayMenu.ImageList;
            _trayMenu.Items.Add(deleteMenu);

            // Menu for hotkeys
            var hotkeyMenu = new ToolStripMenuItem("Set Hotkeys");
            hotkeyMenu.ImageIndex = 7;
            hotkeyMenu.DropDown = new ToolStripDropDownMenu();
            hotkeyMenu.DropDown.ImageList = _trayMenu.ImageList;
            _trayMenu.Items.Add(hotkeyMenu);

            // Add to delete, save and hotkey menus
            foreach (var profile in profiles)
            {
                var itemCaption = Path.GetFileNameWithoutExtension(profile);
                newMenuItem = saveMenu.DropDownItems.Add(itemCaption);
                newMenuItem.Click += OnMenuSave;
                newMenuItem.ImageIndex = 3;

                newMenuItem = deleteMenu.DropDownItems.Add(itemCaption);
                newMenuItem.Click += OnMenuDelete;
                newMenuItem.ImageIndex = 3;

                var hotkeyString = "(No Hotkey)";
                // check if a hotkey is assigned
                var hotkey = FindHotkey(Path.GetFileNameWithoutExtension(profile));
                if (hotkey != null)
                {
                    hotkeyString = "(" + hotkey.ToString() + ")";
                }

                newMenuItem = hotkeyMenu.DropDownItems.Add(itemCaption + " " + hotkeyString);
                newMenuItem.Tag = itemCaption;
                newMenuItem.Click += OnHotkeySet;
                newMenuItem.ImageIndex = 3;
            }

            _trayMenu.Items.Add("-");
            newMenuItem = _trayMenu.Items.Add("Turn Off All Monitors");
            newMenuItem.Click += OnEnergySaving;
            newMenuItem.ImageIndex = 0;

            _trayMenu.Items.Add("-");
            newMenuItem = _trayMenu.Items.Add("About");
            newMenuItem.Click += OnMenuAbout;
            newMenuItem.ImageIndex = 6;

            newMenuItem = _trayMenu.Items.Add("Donate");
            newMenuItem.Click += OnMenuDonate;
            newMenuItem.ImageIndex = 8;

            newMenuItem = _trayMenu.Items.Add("Exit");
            newMenuItem.Click += OnMenuExit;
            newMenuItem.ImageIndex = 2;
        }

        public string ProfileFileFromName(string name)
        {
            var fileName = name + ".xml";
            var filePath = Path.Combine(_settingsDirectoryProfiles, fileName);

            return filePath;
        }

        public string SettingsFileFromName(string name)
        {
            var fileName = name + ".xml";
            var filePath = Path.Combine(_settingsDirectory, fileName);

            return filePath;
        }

        public void OnEnergySaving(object sender, EventArgs e)
        {
            System.Threading.Thread.Sleep(500); // wait for 500 milliseconds to give the user the chance to leave the mouse alone
            SendMessageAPI.PostMessage(new IntPtr(SendMessageAPI.HWND_BROADCAST), SendMessageAPI.WM_SYSCOMMAND, new IntPtr(SendMessageAPI.SC_MONITORPOWER), new IntPtr(SendMessageAPI.MONITOR_OFF));
        }

        public void OnMenuAbout(object sender, EventArgs e)
        {
            MessageBox.Show("Monitor Profile Switcher by Martin Krämer, Matthias Voigt \n(m.voigt@code-ix.de)\nVersion 0.8.0.0\nCopyright 2013-2023 \n\nhttps://github.com/Matt-17/Monitor-Profile-Switcher", "About Monitor Profile Switcher", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void OnMenuDonate(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.paypal.com/donate/?hosted_button_id=LCLKNELZ3VKFQ");
        }

        public void OnMenuSaveAs(object sender, EventArgs e)
        {
            var profileName = "New Profile";
            if (InputBox("Save as new profile", "Enter name of new profile", ref profileName) == DialogResult.OK)
            {
                var invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
                foreach (var invalidChar in invalidChars)
                {
                    profileName = profileName.Replace(invalidChar.ToString(), "");
                }

                if (profileName.Trim().Length > 0)
                {
                    if (!MonitorSwitcher.MonitorSwitcher.SaveDisplaySettings(ProfileFileFromName(profileName)))
                    {
                        _trayIcon.BalloonTipTitle = "Failed to save Multi Monitor profile";
                        _trayIcon.BalloonTipText = "MonitorSwitcher was unable to save the current profile to a new profile with name\"" + profileName + "\"";
                        _trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                        _trayIcon.ShowBalloonTip(5000);
                    }
                }
            }
        }

        public void OnHotkeySet(object sender, EventArgs e)
        {
            var profileName = (((ToolStripMenuItem)sender).Tag as string);
            var hotkey = FindHotkey(profileName);
            bool isNewHotkey = hotkey == null;
            if (HotkeySetting("Set Hotkey for Monitor Profile '" + profileName + "'", "Enter name of new profile", ref hotkey) != DialogResult.OK) 
                return;

            if (hotkey != null)
            {
                if (isNewHotkey)
                {
                    if (!hotkey.RemoveKey)
                    {
                        hotkey.profileName = profileName;
                        _hotkeys.Add(hotkey);
                    }
                }
                else
                {
                    if (hotkey.RemoveKey)
                    {
                        _hotkeys.Remove(hotkey);
                    }
                }
            }

            KeyHooksRefresh();
            SaveSettings();
        }

        public void LoadProfile(string name)
        {
            if (!MonitorSwitcher.MonitorSwitcher.LoadDisplaySettings(ProfileFileFromName(name)))
            {
                _trayIcon.BalloonTipTitle = "Failed to load Multi Monitor profile";
                _trayIcon.BalloonTipText = "MonitorSwitcher was unable to load the previously saved profile \"" + name + "\"";
                _trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                _trayIcon.ShowBalloonTip(5000);
            }
        }

        public void OnMenuLoad(object sender, EventArgs e)
        {
            LoadProfile(((ToolStripMenuItem)sender).Text);
        }

        public void OnMenuSave(object sender, EventArgs e)
        {
            if (!MonitorSwitcher.MonitorSwitcher.SaveDisplaySettings(ProfileFileFromName(((ToolStripMenuItem)sender).Text)))
            {
                _trayIcon.BalloonTipTitle = "Failed to save Multi Monitor profile";
                _trayIcon.BalloonTipText = "MonitorSwitcher was unable to save the current profile to name\"" + ((ToolStripMenuItem)sender).Text + "\"";
                _trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                _trayIcon.ShowBalloonTip(5000);
            }
        }

        public void OnMenuDelete(object sender, EventArgs e)
        {
            File.Delete(ProfileFileFromName(((ToolStripMenuItem)sender).Text));
        }

        public void OnTrayClick(object sender, MouseEventArgs e)
        {
            BuildTrayMenu();

            if (e.Button == MouseButtons.Left)
            {
                var mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
                mi.Invoke(_trayIcon, null);
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            KeyHooksRefresh();
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.

            base.OnLoad(e);
        }

        private void OnMenuExit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                _trayIcon.Dispose();
            }

            base.Dispose(isDisposing);
        }

        public static DialogResult HotkeySetting(string title, string promptText, ref Hotkey value)
        {
            var form = new Form();
            var label = new Label();
            var textBox = new TextBox();
            var buttonOk = new Button();
            var buttonCancel = new Button();
            var buttonClear = new Button();

            form.Text = title;
            label.Text = "Press hotkey combination or click 'Clear Hotkey' to remove the current hotkey";
            if (value != null)
                textBox.Text = value.ToString();
            textBox.Tag = value;

            buttonClear.Text = "Clear Hotkey";
            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 10, 372, 13);
            textBox.SetBounds(12, 36, 372 - 75 - 8, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);
            buttonClear.SetBounds(309, 36 - 1, 75, 23);

            buttonClear.Tag = textBox;
            buttonClear.Click += buttonClear_Click;
            textBox.KeyDown += textBox_KeyDown;
            textBox.KeyUp += textBox_KeyUp;

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonClear.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel, buttonClear });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            var dialogResult = form.ShowDialog();
            value = (textBox.Tag as Hotkey);
            return dialogResult;
        }

        static void textBox_KeyUp(object sender, KeyEventArgs e)
        {
            var textBox = (sender as TextBox);

            if (textBox.Tag != null)
            {
                var hotkey = (textBox.Tag as Hotkey);
                // check if any additional key was pressed, if not don't acceppt hotkey
                if ((hotkey.Key < Keys.D0) || ((!hotkey.Alt) && (!hotkey.Ctrl) && (!hotkey.Shift)))
                    textBox.Text = "";
            }
        }

        static void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            var textBox = (TextBox)sender;
            var hotkey = textBox.Tag as Hotkey ?? new Hotkey();
            hotkey.AssignFromKeyEventArgs(e);

            e.Handled = true;
            e.SuppressKeyPress = true; // don't add user input to text box, just use custom display

            textBox.Text = hotkey.ToString();
            textBox.Tag = hotkey; // store the current key combination in the textbox tag (for later use)
        }

        static void buttonClear_Click(object sender, EventArgs e)
        {
            var textBox = (sender as Button).Tag as TextBox;

            if (textBox.Tag != null)
            {
                var hotkey = (textBox.Tag as Hotkey);
                hotkey.RemoveKey = true;
            }
            textBox.Clear();
        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            var form = new Form();
            var label = new Label();
            var textBox = new TextBox();
            var buttonOk = new Button();
            var buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 10, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            var dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
    }
}
