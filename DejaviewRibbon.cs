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

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Threading;
using System.Windows.Forms;

namespace Dejaview
{
    /// <summary>
    /// Ribbon class for MS Word application interface.
    /// </summary>
    public partial class DejaviewRibbon
    {
        private void DejaviewRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            System.Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            btnUpdate.SuperTip = "Check the Deja View website for updates to this Add-in.\n\nDeja View version: " + lVersion.ToString();
        }

        private void btnRemove_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.DejaviewAddIn.RemoveDejaviewFromDocument(Globals.DejaviewAddIn.Application.ActiveDocument);
        }

        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            Thread updateThread = new Thread(DejaviewAddIn.CheckForUpdate);
            updateThread.Start(false);
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            OptionsDialog optionsDialog = new OptionsDialog();
            optionsDialog.ShowDialog();
        }

        private void btnApplyDefault_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.DejaviewAddIn.Application.ActiveDocument;
            DejaviewSet s = DejaviewConfig.Instance.DefaultDejaviewSet;
            if (s != null)
            {
                DialogResult r = MessageBox.Show(null, "Do you want to apply the default view to the current document?", "Apply Default View?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (r == DialogResult.Yes)
                {
                    Globals.DejaviewAddIn.ShowDocumentView(doc, s);
                    Globals.DejaviewAddIn.Log("Applied default view to current document.", doc);
                }
            }
            else
            {
                Globals.DejaviewAddIn.Log("No default view is set.", doc);
            }
        }
    }
}
