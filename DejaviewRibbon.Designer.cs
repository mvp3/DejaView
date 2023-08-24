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
            this.btnRemove = this.Factory.CreateRibbonButton();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.btnApplyDefault = this.Factory.CreateRibbonButton();
            this.btnSettings = this.Factory.CreateRibbonButton();
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
            this.groupDejaview.Items.Add(this.btnRemove);
            this.groupDejaview.Items.Add(this.btnUpdate);
            this.groupDejaview.Items.Add(this.btnApplyDefault);
            this.groupDejaview.Items.Add(this.btnSettings);
            this.groupDejaview.Label = "Deja View";
            this.groupDejaview.Name = "groupDejaview";
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
            // btnUpdate
            // 
            this.btnUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdate.Label = "Check for Update";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.OfficeImageId = "DataSourceCatalogServerScript";
            this.btnUpdate.ScreenTip = "Check for Updates";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.SuperTip = "Check the Deja View website for updates to this Add-in.";
            this.btnUpdate.Visible = false;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // btnApplyDefault
            // 
            this.btnApplyDefault.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApplyDefault.Label = "Apply Default View";
            this.btnApplyDefault.Name = "btnApplyDefault";
            this.btnApplyDefault.OfficeImageId = "ZoomFitToWindow";
            this.btnApplyDefault.ScreenTip = "Open Options Dialog";
            this.btnApplyDefault.ShowImage = true;
            this.btnApplyDefault.SuperTip = "View Deja View options dialog.";
            this.btnApplyDefault.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyDefault_Click);
            // 
            // btnSettings
            // 
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Label = "Options";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.OfficeImageId = "OmsViewAccountSetting";
            this.btnSettings.ScreenTip = "Open Options Dialog";
            this.btnSettings.ShowImage = true;
            this.btnSettings.SuperTip = "View Deja View options dialog.";
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemove;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyDefault;
    }

    partial class ThisRibbonCollection
    {
        internal DejaviewRibbon DejaviewRibbon
        {
            get { return this.GetRibbon<DejaviewRibbon>(); }
        }
    }
}
