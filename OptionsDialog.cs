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
    public partial class OptionsDialog : Form
    {
        public OptionsDialog()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OptionsDialog_Load(object sender, EventArgs e)
        {
            chkEnable.Checked = DejaviewConfig.Instance.Enable;
            chkAutoUpdate.Checked = DejaviewConfig.Instance.CheckForUpdates;
            chkPrompt.Checked = DejaviewConfig.Instance.Prompt;

            chkLocation.Checked = DejaviewConfig.Instance.RememberWindowLocation;
            chkNavigationPanel.Checked = DejaviewConfig.Instance.RememberNavigationPanel;
            chkWindowType.Checked = DejaviewConfig.Instance.RememberWindowType;
            chkZoom.Checked = DejaviewConfig.Instance.RememberZoom;
            chkRulers.Checked = DejaviewConfig.Instance.RememberRulers;
            chkRibbon.Checked = DejaviewConfig.Instance.RememberRibbon;

            setEnabled(chkEnable.Checked);

            btnViewTags.Enabled = Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled;

            Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            lblVersion.Text = "Version: " + lVersion.ToString();
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
            DejaviewConfig.Instance.CheckForUpdates = chkAutoUpdate.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkEnable_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.Enable = chkEnable.Checked;
            DejaviewConfig.Instance.Save();

            setEnabled(chkEnable.Checked);

            // If switching to enabled invoke DocumentOpen method to read Dejaview tags
            if (chkEnable.Enabled)
                Globals.DejaviewAddIn.DejaviewAddIn_DocumentOpen(Globals.DejaviewAddIn.Application.ActiveDocument);
        }

        private void chkPrompt_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.Prompt = chkPrompt.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkLocation_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberWindowLocation = chkLocation.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkNav_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberNavigationPanel = chkNavigationPanel.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkWindowType_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberWindowType = chkWindowType.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkZoom_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberZoom = chkZoom.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkRulers_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberRulers = chkRulers.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkRibbon_CheckedChanged(object sender, EventArgs e)
        {
            DejaviewConfig.Instance.RememberRibbon = chkRibbon.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void btnViewTags_Click(object sender, EventArgs e)
        {
            BasicDialog bd = new BasicDialog("Deja View Tags", Globals.DejaviewAddIn.GetSavedTags());
            bd.Show();
        }

        private void btnViewCurrent_Click(object sender, EventArgs e)
        {
            BasicDialog bd = new BasicDialog("Current View", Globals.DejaviewAddIn.GetCurrentParameters());
            bd.Show();
        }
    }
}
