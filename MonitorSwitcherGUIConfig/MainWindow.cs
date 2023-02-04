using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MonitorSwitcherGUI
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void toolStripContainer1_TopToolStripPanel_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateProfileList();
        }

        private void UpdateProfileList()
        {
            String settingsDirectory = MonitorSwitcherGUI.GetSettingsDirectory(null);
            string settingsDirectoryProfiles = MonitorSwitcherGUI.GetSettingsProfielDirectotry(settingsDirectory);

            // get profiles
            string[] profiles = Directory.GetFiles(settingsDirectoryProfiles, "*.xml");
            foreach (string profile in profiles)
            {
                string itemCaption = Path.GetFileNameWithoutExtension(profile);
                lbProfiles.Items.Add(itemCaption);
            }

            UpdateGUIStatus();
        }

        private void UpdateGUIStatus()
        {
            tsbAdd.Enabled = true;

            tsbDelete.Enabled = (lbProfiles.SelectedItem != null);
            tsbExport.Enabled = (lbProfiles.SelectedItem != null);
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            
        }

        private void lbProfiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateGUIStatus();
        }
    }
}
