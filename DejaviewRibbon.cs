/**
 * Copyright (C) 2021 M. V. Pereira - All Rights Reserved
 * 
 * This AddIn is available at: http://lexem.cc/dejaview/
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

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;

namespace Dejaview
{
    /// <summary>
    /// Ribbon class for MS Word application interface.
    /// </summary>
    public partial class DejaviewRibbon
    {
        private void DejaviewRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            chkEnable.Checked = DejaviewConfig.Instance.Enable;
            chkLocation.Checked = DejaviewConfig.Instance.Location;
            chkPrompt.Checked = DejaviewConfig.Instance.Prompt;

            chkLocation.Enabled = chkEnable.Checked;
            chkPrompt.Enabled = chkEnable.Checked;

            Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            chkEnable.SuperTip = "This option allows a quick and easy means to temporarily enable / disable Deja View.\n\nRevision: " + lVersion.ToString();
        }

        private void chkEnable_Click(object sender, RibbonControlEventArgs e)
        {
            DejaviewConfig.Instance.Enable = chkEnable.Checked;
            DejaviewConfig.Instance.Save();

            chkLocation.Enabled = chkEnable.Checked;
            chkPrompt.Enabled = chkEnable.Checked;

            // If switching to enabled invoke DocumentOpen method to read Dejaview tags
            if (chkEnable.Enabled)
                Globals.DejaviewAddIn.DejaviewAddIn_DocumentOpen(Globals.DejaviewAddIn.Application.ActiveDocument);
        }

        private void chkLocation_Click(object sender, RibbonControlEventArgs e)
        {
            DejaviewConfig.Instance.Location = chkLocation.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void chkPrompt_Click(object sender, RibbonControlEventArgs e)
        {
            DejaviewConfig.Instance.Prompt = chkPrompt.Checked;
            DejaviewConfig.Instance.Save();
        }

        private void btnRemove_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var xml = Globals.DejaviewAddIn.Application.ActiveDocument.CustomXMLParts["Dejaview"];
                if (xml != null) xml.Delete();
                btnRemove.Enabled = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewRibbon::btnClear_Click() => " + ex.StackTrace);
                Globals.DejaviewAddIn.DisplayStatus("Could not clear tags: " + ex.Message);
            }
        }
    }
}
