/**
 * Copyright (C) 2021 M. V. Pereira - All Rights Reserved
 * 
 * This AddIn is available at: https://dejaview.lexem.cc/
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License. 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dejaview
{
    /// <summary>
    /// Dialog containing all of the configurable options for Deja View.
    /// </summary>
    public partial class OptionsDialog : Form
    {
        /// <summary>
        /// Used for checking to see if the update URL textbox changes
        /// </summary>
        private string updateURL = null;

        /// <summary>
        /// Used for bypassing the change events
        /// </summary>
        private bool bypassChange = false;

        /// <summary>
        /// Default constructor
        /// </summary>
        public OptionsDialog()
        {
            bypassChange = true;
            InitializeComponent();
            bypassChange = false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OptionsDialog_Load(object sender, EventArgs e)
        {
            bypassChange = true;

            chkEnable.Checked = DejaviewConfig.Instance.Enable;
            chkAutoUpdate.Checked = DejaviewConfig.Instance.CheckForUpdates;
            chkPrompt.Checked = DejaviewConfig.Instance.Prompt;

            chkLocation.Checked = DejaviewConfig.Instance.RememberWindowLocation;
            chkNavigationPanel.Checked = DejaviewConfig.Instance.RememberNavigationPanel;
            chkWindowType.Checked = DejaviewConfig.Instance.RememberWindowType;
            chkZoom.Checked = DejaviewConfig.Instance.RememberZoom;
            chkRulers.Checked = DejaviewConfig.Instance.RememberRulers;
            chkRibbon.Checked = DejaviewConfig.Instance.RememberRibbon;

            txtUpdateURL.Text = DejaviewConfig.Instance.UpdateURL;
            updateURL = txtUpdateURL.Text;

            setEnabled(chkEnable.Checked);

            btnViewTags.Enabled = Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled;

            Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            lblVersion.Text = "Version: " + lVersion.ToString();

            bypassChange = false;
        }

        private void OptionsDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (txtUpdateURL.Text != updateURL)
            {
                DialogResult r = MessageBox.Show(this, "The update URL has changed.\n\nDo you want to save these changes?", "Save Changes?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r == DialogResult.Yes)
                {
                    DejaviewConfig.Instance.UpdateURL = txtUpdateURL.Text;
                    DejaviewConfig.Instance.Save();
                }
            }
        }

        private void setEnabled(bool enabled)
        {
            chkPrompt.Enabled = enabled;
            chkAutoUpdate.Enabled = enabled;

            foreach (Control x in grpRemember.Controls)
                x.Enabled = enabled;
        }

        private void chkAutoUpdate_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.CheckForUpdates = chkAutoUpdate.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkEnable_CheckedChanged(object sender, EventArgs e)
        {
            setEnabled(chkEnable.Checked);

            if (bypassChange) return;

            DejaviewConfig.Instance.Enable = chkEnable.Checked;
            DejaviewConfig.Instance.Save();

            Microsoft.Office.Interop.Word.Document doc = Globals.DejaviewAddIn.Application.ActiveDocument;

            // If switching to enabled invoke DocumentOpen method to read Deja View tags
            if (chkEnable.Enabled && !Globals.DejaviewAddIn.IsLoaded(doc))
                Globals.DejaviewAddIn.DejaviewAddIn_DocumentOpen(doc);
        }

        private void chkPrompt_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.Prompt = chkPrompt.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkLocation_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberWindowLocation = chkLocation.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkNav_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberNavigationPanel = chkNavigationPanel.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkWindowType_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberWindowType = chkWindowType.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkZoom_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberZoom = chkZoom.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkRulers_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberRulers = chkRulers.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkRibbon_CheckedChanged(object sender, EventArgs e)
        {
            if (bypassChange) return;
            DejaviewConfig.Instance.RememberRibbon = chkRibbon.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void btnViewTags_Click(object sender, EventArgs e)
        {
            DejaviewSet d = Globals.DejaviewAddIn.GetDejaviewSet();
            BasicDialog bd = new BasicDialog("Deja View Tags", Globals.DejaviewAddIn.GetTags(d));
            bd.Show();
        }

        private void btnViewCurrent_Click(object sender, EventArgs e)
        {
            DejaviewSet d = Globals.DejaviewAddIn.GetDejaviewSetFromDisplay();
            BasicDialog bd = new BasicDialog("Current View", Globals.DejaviewAddIn.GetTags(d));
            bd.Show();
        }

        private void btnSetDefaults_Click(object sender, EventArgs e)
        {
            DialogResult r = MessageBox.Show(this, "This will restore all Deja View options to default.\n\nDo you want to continue?", "Overwrite Settings?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (r == DialogResult.Yes)
            {
                bypassChange = true;
                DejaviewConfig.Instance.SetDefaults();
                OptionsDialog_Load(this, null);
                bypassChange = false;
            }
        }

        private void OptionsDialog_DoubleClick(object sender, EventArgs e)
        {
            BasicDialog bd = new BasicDialog("Log", Globals.DejaviewAddIn.GetLogger().ToString());
            bd.Show();
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            BasicDialog bd = new BasicDialog("Log", Globals.DejaviewAddIn.GetLogger().ToString());
            bd.Show();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            // Get Logger
            Logger logger = Globals.DejaviewAddIn.GetLogger();
            logger.Add("Checking views...");

            DejaviewSet s = Globals.DejaviewAddIn.GetDejaviewSet();
            DejaviewSet d = Globals.DejaviewAddIn.GetDejaviewSetFromDisplay();
            if (d.Equals(s))
            {
                logger.Add("  Current view and saved view match!");
                MessageBox.Show(this, "The current document view is the same as the last saved document view.", "Same", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                logger.Add("  Applying saved view.");
                Globals.DejaviewAddIn.SetDocumentView(Globals.DejaviewAddIn.Application.ActiveDocument, s);
            }
        }

        private void btnSetDefaultView_Click(object sender, EventArgs e)
        {
            DialogResult r = MessageBox.Show(this, "This will set the default new document view to the current document view.\n\nDo you want to continue?", "Set Default View?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (r == DialogResult.Yes)
            {
                DejaviewSet s = Globals.DejaviewAddIn.GetDejaviewSet();
                DejaviewConfig.Instance.DefaultDejaviewSet = s;
                DejaviewConfig.Instance.Save();
            }
        }
    }
}
