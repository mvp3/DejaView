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

namespace Dejaview
{
    partial class DejaviewRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Ribbon object for docking in Microsoft Word.
        /// </summary>
        public DejaviewRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupDejaview = this.Factory.CreateRibbonGroup();
            this.chkEnable = this.Factory.CreateRibbonCheckBox();
            this.chkLocation = this.Factory.CreateRibbonCheckBox();
            this.chkPrompt = this.Factory.CreateRibbonCheckBox();
            this.btnRemove = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupDejaview.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupDejaview);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupDejaview
            // 
            this.groupDejaview.Items.Add(this.chkEnable);
            this.groupDejaview.Items.Add(this.chkLocation);
            this.groupDejaview.Items.Add(this.chkPrompt);
            this.groupDejaview.Items.Add(this.btnRemove);
            this.groupDejaview.Label = "Deja View";
            this.groupDejaview.Name = "groupDejaview";
            // 
            // chkEnable
            // 
            this.chkEnable.Checked = true;
            this.chkEnable.Label = "Enable";
            this.chkEnable.Name = "chkEnable";
            this.chkEnable.ScreenTip = "Enable Deja View";
            this.chkEnable.SuperTip = "This option allows a quick and easy means to temporarily enable / disable Deja Vi" +
    "ew. ";
            this.chkEnable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkEnable_Click);
            // 
            // chkLocation
            // 
            this.chkLocation.Checked = true;
            this.chkLocation.Label = "Location";
            this.chkLocation.Name = "chkLocation";
            this.chkLocation.ScreenTip = "Window Location";
            this.chkLocation.SuperTip = "Deja View will remember the document window\'s location. Default is checked. ";
            this.chkLocation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkLocation_Click);
            // 
            // chkPrompt
            // 
            this.chkPrompt.Label = "Prompt";
            this.chkPrompt.Name = "chkPrompt";
            this.chkPrompt.ScreenTip = "Ask Before Saving";
            this.chkPrompt.SuperTip = "If checked, Deja View will ask before saving view settings to this document. Defa" +
    "ult is unchecked.";
            this.chkPrompt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkPrompt_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRemove.Label = "Clear View Tags";
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.OfficeImageId = "TableOfContentsRemove";
            this.btnRemove.ScreenTip = "Clear Deja View Tags";
            this.btnRemove.ShowImage = true;
            this.btnRemove.SuperTip = "Remove all Deja View tags from this document.";
            this.btnRemove.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRemove_Click);
            // 
            // DejaviewRibbon
            // 
            this.Name = "DejaviewRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DejaviewRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupDejaview.ResumeLayout(false);
            this.groupDejaview.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDejaview;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnable;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkLocation;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkPrompt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemove;
    }

    partial class ThisRibbonCollection
    {
        internal DejaviewRibbon DejaviewRibbon
        {
            get { return this.GetRibbon<DejaviewRibbon>(); }
        }
    }
}
