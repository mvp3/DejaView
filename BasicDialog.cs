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
using System.Windows.Forms;

namespace Dejaview
{
    /// <summary>
    /// A basic form for displaying text in a scrolling pane.
    /// </summary>
    public partial class BasicDialog : Form
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="caption">String to display as the caption of the dialog</param>
        /// <param name="text">String to be displayed as the main text</param>
        public BasicDialog(string caption, string text)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.Manual;
            Top = Cursor.Position.Y;
            Left = Cursor.Position.X;
            Text = caption;
            txt.Text = text;
        }

        /// <summary>
        /// Sets the string to be displayed as the main text in the body of the dialog.
        /// This method is used as an alternative to the default constructor.
        /// </summary>
        /// <param name="text">String to be displayed as the main text</param>
        public void setText(string text)
        {
            txt.Text = text;
        }

        private void BasicDialog_Load(object sender, EventArgs e)
        {
            txt.SelectionStart = 0;
            txt.SelectionLength = 0;
            this.Focus();
        }
    }
}
