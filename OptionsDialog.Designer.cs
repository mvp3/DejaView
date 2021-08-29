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
    partial class OptionsDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.grpRemember = new System.Windows.Forms.GroupBox();
            this.chkWindowType = new System.Windows.Forms.CheckBox();
            this.chkRulers = new System.Windows.Forms.CheckBox();
            this.chkZoom = new System.Windows.Forms.CheckBox();
            this.chkRibbon = new System.Windows.Forms.CheckBox();
            this.chkNavigationPanel = new System.Windows.Forms.CheckBox();
            this.chkLocation = new System.Windows.Forms.CheckBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.grpSettings = new System.Windows.Forms.GroupBox();
            this.btnSetDefaults = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtUpdateURL = new System.Windows.Forms.TextBox();
            this.btnViewCurrent = new System.Windows.Forms.Button();
            this.btnViewTags = new System.Windows.Forms.Button();
            this.chkEnable = new System.Windows.Forms.CheckBox();
            this.chkPrompt = new System.Windows.Forms.CheckBox();
            this.chkAutoUpdate = new System.Windows.Forms.CheckBox();
            this.lblVersion = new System.Windows.Forms.Label();
            this.tip = new System.Windows.Forms.ToolTip(this.components);
            this.btnLogs = new System.Windows.Forms.Button();
            this.grpRemember.SuspendLayout();
            this.grpSettings.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpRemember
            // 
            this.grpRemember.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpRemember.Controls.Add(this.chkWindowType);
            this.grpRemember.Controls.Add(this.chkRulers);
            this.grpRemember.Controls.Add(this.chkZoom);
            this.grpRemember.Controls.Add(this.chkRibbon);
            this.grpRemember.Controls.Add(this.chkNavigationPanel);
            this.grpRemember.Controls.Add(this.chkLocation);
            this.grpRemember.Location = new System.Drawing.Point(12, 178);
            this.grpRemember.Name = "grpRemember";
            this.grpRemember.Size = new System.Drawing.Size(288, 100);
            this.grpRemember.TabIndex = 0;
            this.grpRemember.TabStop = false;
            this.grpRemember.Text = "Remember";
            // 
            // chkWindowType
            // 
            this.chkWindowType.AutoSize = true;
            this.chkWindowType.Location = new System.Drawing.Point(6, 71);
            this.chkWindowType.Name = "chkWindowType";
            this.chkWindowType.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.chkWindowType.Size = new System.Drawing.Size(98, 17);
            this.chkWindowType.TabIndex = 7;
            this.chkWindowType.Text = "Window Type";
            this.tip.SetToolTip(this.chkWindowType, "Save and restore the document window view type");
            this.chkWindowType.UseVisualStyleBackColor = true;
            this.chkWindowType.CheckedChanged += new System.EventHandler(this.chkWindowType_CheckedChanged);
            // 
            // chkRulers
            // 
            this.chkRulers.AutoSize = true;
            this.chkRulers.Location = new System.Drawing.Point(137, 48);
            this.chkRulers.Name = "chkRulers";
            this.chkRulers.Size = new System.Drawing.Size(56, 17);
            this.chkRulers.TabIndex = 6;
            this.chkRulers.Text = "Rulers";
            this.tip.SetToolTip(this.chkRulers, "Save and restore the setting to show rulers");
            this.chkRulers.UseVisualStyleBackColor = true;
            this.chkRulers.CheckedChanged += new System.EventHandler(this.chkRulers_CheckedChanged);
            // 
            // chkZoom
            // 
            this.chkZoom.AutoSize = true;
            this.chkZoom.Location = new System.Drawing.Point(137, 19);
            this.chkZoom.Name = "chkZoom";
            this.chkZoom.Padding = new System.Windows.Forms.Padding(0, 6, 0, 0);
            this.chkZoom.Size = new System.Drawing.Size(53, 23);
            this.chkZoom.TabIndex = 5;
            this.chkZoom.Text = "Zoom";
            this.tip.SetToolTip(this.chkZoom, "Save and restore the document zoom level");
            this.chkZoom.UseVisualStyleBackColor = true;
            this.chkZoom.CheckedChanged += new System.EventHandler(this.chkZoom_CheckedChanged);
            // 
            // chkRibbon
            // 
            this.chkRibbon.AutoSize = true;
            this.chkRibbon.Location = new System.Drawing.Point(137, 71);
            this.chkRibbon.Name = "chkRibbon";
            this.chkRibbon.Size = new System.Drawing.Size(60, 17);
            this.chkRibbon.TabIndex = 4;
            this.chkRibbon.Text = "Ribbon";
            this.tip.SetToolTip(this.chkRibbon, "Save and restore the settings for how the ribbon should be displayed");
            this.chkRibbon.UseVisualStyleBackColor = true;
            this.chkRibbon.CheckedChanged += new System.EventHandler(this.chkRibbon_CheckedChanged);
            // 
            // chkNavigationPanel
            // 
            this.chkNavigationPanel.AutoSize = true;
            this.chkNavigationPanel.Location = new System.Drawing.Point(6, 48);
            this.chkNavigationPanel.Name = "chkNavigationPanel";
            this.chkNavigationPanel.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.chkNavigationPanel.Size = new System.Drawing.Size(113, 17);
            this.chkNavigationPanel.TabIndex = 2;
            this.chkNavigationPanel.Text = "Navigation Panel";
            this.tip.SetToolTip(this.chkNavigationPanel, "Save and restore the visibility and width details for side navigation panel");
            this.chkNavigationPanel.UseVisualStyleBackColor = true;
            this.chkNavigationPanel.CheckedChanged += new System.EventHandler(this.chkNav_CheckedChanged);
            // 
            // chkLocation
            // 
            this.chkLocation.AutoSize = true;
            this.chkLocation.Location = new System.Drawing.Point(6, 19);
            this.chkLocation.Name = "chkLocation";
            this.chkLocation.Padding = new System.Windows.Forms.Padding(6, 6, 10, 0);
            this.chkLocation.Size = new System.Drawing.Size(125, 23);
            this.chkLocation.TabIndex = 1;
            this.chkLocation.Text = "Window Location";
            this.tip.SetToolTip(this.chkLocation, "Save and restore the location of the Microsoft Word document window");
            this.chkLocation.UseVisualStyleBackColor = true;
            this.chkLocation.CheckedChanged += new System.EventHandler(this.chkLocation_CheckedChanged);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(219, 289);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.tip.SetToolTip(this.btnClose, "Close this dialog. Options are automatically saved.");
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // grpSettings
            // 
            this.grpSettings.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpSettings.Controls.Add(this.btnSetDefaults);
            this.grpSettings.Controls.Add(this.groupBox1);
            this.grpSettings.Controls.Add(this.btnViewCurrent);
            this.grpSettings.Controls.Add(this.btnViewTags);
            this.grpSettings.Controls.Add(this.chkEnable);
            this.grpSettings.Controls.Add(this.chkPrompt);
            this.grpSettings.Controls.Add(this.chkAutoUpdate);
            this.grpSettings.Location = new System.Drawing.Point(12, 12);
            this.grpSettings.Name = "grpSettings";
            this.grpSettings.Size = new System.Drawing.Size(288, 160);
            this.grpSettings.TabIndex = 2;
            this.grpSettings.TabStop = false;
            this.grpSettings.Text = "Add-in Settings";
            // 
            // btnSetDefaults
            // 
            this.btnSetDefaults.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSetDefaults.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetDefaults.Location = new System.Drawing.Point(207, 77);
            this.btnSetDefaults.Name = "btnSetDefaults";
            this.btnSetDefaults.Size = new System.Drawing.Size(75, 23);
            this.btnSetDefaults.TabIndex = 6;
            this.btnSetDefaults.Text = "Defaults";
            this.tip.SetToolTip(this.btnSetDefaults, "Set all options to default");
            this.btnSetDefaults.UseVisualStyleBackColor = true;
            this.btnSetDefaults.Click += new System.EventHandler(this.btnSetDefaults_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.txtUpdateURL);
            this.groupBox1.Location = new System.Drawing.Point(6, 104);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(276, 50);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Update URL";
            // 
            // txtUpdateURL
            // 
            this.txtUpdateURL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtUpdateURL.Location = new System.Drawing.Point(6, 19);
            this.txtUpdateURL.Name = "txtUpdateURL";
            this.txtUpdateURL.Size = new System.Drawing.Size(264, 20);
            this.txtUpdateURL.TabIndex = 0;
            // 
            // btnViewCurrent
            // 
            this.btnViewCurrent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnViewCurrent.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnViewCurrent.Location = new System.Drawing.Point(207, 48);
            this.btnViewCurrent.Name = "btnViewCurrent";
            this.btnViewCurrent.Size = new System.Drawing.Size(75, 23);
            this.btnViewCurrent.TabIndex = 4;
            this.btnViewCurrent.Text = "Current";
            this.tip.SetToolTip(this.btnViewCurrent, "Show the current document view settings that will be saved in the Deja View tags");
            this.btnViewCurrent.UseVisualStyleBackColor = true;
            this.btnViewCurrent.Click += new System.EventHandler(this.btnViewCurrent_Click);
            // 
            // btnViewTags
            // 
            this.btnViewTags.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnViewTags.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnViewTags.Location = new System.Drawing.Point(207, 19);
            this.btnViewTags.Name = "btnViewTags";
            this.btnViewTags.Size = new System.Drawing.Size(75, 23);
            this.btnViewTags.TabIndex = 3;
            this.btnViewTags.Text = "View Tags";
            this.tip.SetToolTip(this.btnViewTags, "Show the Deja View tags that are saved in this document");
            this.btnViewTags.UseVisualStyleBackColor = true;
            this.btnViewTags.Click += new System.EventHandler(this.btnViewTags_Click);
            // 
            // chkEnable
            // 
            this.chkEnable.AutoSize = true;
            this.chkEnable.Location = new System.Drawing.Point(6, 19);
            this.chkEnable.Name = "chkEnable";
            this.chkEnable.Padding = new System.Windows.Forms.Padding(6, 6, 3, 2);
            this.chkEnable.Size = new System.Drawing.Size(119, 25);
            this.chkEnable.TabIndex = 2;
            this.chkEnable.Text = "Enable Deja View";
            this.tip.SetToolTip(this.chkEnable, "This option allows a quick and easy means to temporarily enable / disable Deja Vi" +
        "ew. ");
            this.chkEnable.UseVisualStyleBackColor = true;
            this.chkEnable.CheckedChanged += new System.EventHandler(this.chkEnable_CheckedChanged);
            // 
            // chkPrompt
            // 
            this.chkPrompt.AutoSize = true;
            this.chkPrompt.Location = new System.Drawing.Point(6, 50);
            this.chkPrompt.Name = "chkPrompt";
            this.chkPrompt.Padding = new System.Windows.Forms.Padding(6, 2, 3, 2);
            this.chkPrompt.Size = new System.Drawing.Size(199, 21);
            this.chkPrompt.TabIndex = 1;
            this.chkPrompt.Text = "Prompt before saving view settings";
            this.tip.SetToolTip(this.chkPrompt, "If checked, Deja View will ask before saving view settings to this document.");
            this.chkPrompt.UseVisualStyleBackColor = true;
            this.chkPrompt.CheckedChanged += new System.EventHandler(this.chkPrompt_CheckedChanged);
            // 
            // chkAutoUpdate
            // 
            this.chkAutoUpdate.AutoSize = true;
            this.chkAutoUpdate.Location = new System.Drawing.Point(6, 77);
            this.chkAutoUpdate.Name = "chkAutoUpdate";
            this.chkAutoUpdate.Padding = new System.Windows.Forms.Padding(6, 2, 3, 2);
            this.chkAutoUpdate.Size = new System.Drawing.Size(186, 21);
            this.chkAutoUpdate.TabIndex = 0;
            this.chkAutoUpdate.Text = "Automatically check for updates";
            this.tip.SetToolTip(this.chkAutoUpdate, "Automatically check for updates on startup. Will not check more than once per day" +
        ".");
            this.chkAutoUpdate.UseVisualStyleBackColor = true;
            this.chkAutoUpdate.CheckedChanged += new System.EventHandler(this.chkAutoUpdate_CheckedChanged);
            // 
            // lblVersion
            // 
            this.lblVersion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(15, 294);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(47, 13);
            this.lblVersion.TabIndex = 3;
            this.lblVersion.Text = "(version)";
            // 
            // btnLogs
            // 
            this.btnLogs.Image = global::Dejaview.Properties.Resources.Log_16x;
            this.btnLogs.Location = new System.Drawing.Point(189, 289);
            this.btnLogs.Name = "btnLogs";
            this.btnLogs.Size = new System.Drawing.Size(24, 23);
            this.btnLogs.TabIndex = 4;
            this.tip.SetToolTip(this.btnLogs, "View logs");
            this.btnLogs.UseVisualStyleBackColor = true;
            this.btnLogs.Click += new System.EventHandler(this.btnLogs_Click);
            // 
            // OptionsDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(312, 324);
            this.ControlBox = false;
            this.Controls.Add(this.btnLogs);
            this.Controls.Add(this.lblVersion);
            this.Controls.Add(this.grpSettings);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.grpRemember);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.HelpButton = true;
            this.Name = "OptionsDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Deja View Options";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OptionsDialog_FormClosing);
            this.Load += new System.EventHandler(this.OptionsDialog_Load);
            this.DoubleClick += new System.EventHandler(this.OptionsDialog_DoubleClick);
            this.grpRemember.ResumeLayout(false);
            this.grpRemember.PerformLayout();
            this.grpSettings.ResumeLayout(false);
            this.grpSettings.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grpRemember;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckBox chkLocation;
        private System.Windows.Forms.GroupBox grpSettings;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.CheckBox chkAutoUpdate;
        private System.Windows.Forms.ToolTip tip;
        private System.Windows.Forms.CheckBox chkPrompt;
        private System.Windows.Forms.CheckBox chkEnable;
        private System.Windows.Forms.Button btnViewTags;
        private System.Windows.Forms.CheckBox chkRibbon;
        private System.Windows.Forms.CheckBox chkNavigationPanel;
        private System.Windows.Forms.CheckBox chkZoom;
        private System.Windows.Forms.CheckBox chkWindowType;
        private System.Windows.Forms.CheckBox chkRulers;
        private System.Windows.Forms.Button btnViewCurrent;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtUpdateURL;
        private System.Windows.Forms.Button btnSetDefaults;
        private System.Windows.Forms.Button btnLogs;
    }
}