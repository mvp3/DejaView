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

using System;
using System.Diagnostics;
using System.Drawing;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Dejaview
{
    /// <summary>
    /// Main Add-in class.
    /// </summary>
    public partial class DejaviewAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (!DejaviewConfig.Instance.Enable) return;
                
                // Fire the DocumentOpen event for the first time
                if (this.Application.Documents.Count >= 1)
                    DejaviewAddIn_DocumentOpen(this.Application.ActiveDocument);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error opening DejaviewAddIn::DejaviewAddIn_Startup() => " + ex.Message);
            }

            // Check for updates on a separate thread to ensure
            // a speedy startup for the add-in.
            if (DejaviewConfig.Instance.CheckForUpdates)
            {
                Thread updateThread = new Thread(CheckForUpdate);
                updateThread.Start();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Method called when a document is opened.
        /// </summary>
        /// <param name="doc">Word document that opened</param>
        internal void DejaviewAddIn_DocumentOpen(Word.Document doc)
        {
            if (!DejaviewConfig.Instance.Enable) return;
            try
            {
                DejaviewSet ds = GetDejaviewFromDocument(doc);

                doc.ActiveWindow.WindowState = (Word.WdWindowState)ds.WindowState;
                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateMinimize)
                    doc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;

                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateNormal)
                    SetShowable(doc.Application, ds);

                doc.ActiveWindow.View.Type = (Word.WdViewType)ds.WindowViewType;

                doc.ActiveWindow.View.Zoom.Percentage = ds.WindowZoom;
                doc.ActiveWindow.DisplayRulers = ds.DisplayRulers;

                Office.CommandBar nav = Application.CommandBars["Navigation"];
                nav.Visible = ds.ShowNavigationPanel;
                nav.Width = ds.NavigationPanelWidth;

                Office.CommandBar ribbon = Application.CommandBars["Ribbon"];
                if (ds.RibbonHeight != 0 && ds.RibbonHeight != ribbon.Height)
                {
                    // Use 100 as a toggle threshold instead of simply checking
                    // the saved hieght against the present display height.
                    // If the document is opened on another computer with
                    // a different screen resolution a simple toggle will
                    // misbehave.
                    if ((ds.RibbonHeight > 100 && ribbon.Height < 100) ||
                        (ds.RibbonHeight < 100 && ribbon.Height > 100))
                        doc.ActiveWindow.ToggleRibbon();
                }

                DisplayStatus("Document view restored.");

                SetButtonTip(ds);
            }
            catch (NullReferenceException)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
            }
            catch (Exception ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                DisplayStatus("Could not restore document view. " + ex.Message);
            }
        }

        private void DejaviewAddIn_DocumentBeforeSave(Word.Document doc, ref bool saveAsUI, ref bool cancel)
        {
            if (!DejaviewConfig.Instance.Enable) return;
            try
            {
                bool save = true;
                if (Globals.Ribbons.DejaviewRibbon.chkPrompt.Checked)
                {
                    DialogResult r = MessageBox.Show(null, "Do you want to save this document's view settings?", "Save View?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    save = (r == DialogResult.Yes);
                }
                if (!save) return;

                DejaviewSet ds = new DejaviewSet();

                ds.WindowState = (int)Application.WindowState;
                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateNormal)
                {
                    ds.WindowLeft = (doc.ActiveWindow.Left < 0) ? 0 : doc.ActiveWindow.Left;
                    ds.WindowTop = (doc.ActiveWindow.Top < 0) ? 0 : doc.ActiveWindow.Top;
                    ds.WindowWidth = doc.ActiveWindow.Width;
                    ds.WindowHeight = doc.ActiveWindow.Height;
                }
                ds.WindowViewType = (int)doc.ActiveWindow.View.Type;
                if (ds.WindowViewType == (int)Word.WdViewType.wdNormalView) ds.DraftView = true;

                ds.WindowZoom = doc.ActiveWindow.View.Zoom.Percentage;
                ds.DisplayRulers = doc.ActiveWindow.DisplayRulers;

                Office.CommandBar nav = doc.CommandBars["Navigation"];
                ds.ShowNavigationPanel = nav.Visible;
                ds.NavigationPanelWidth = nav.Width;

                Office.CommandBar ribbon = doc.CommandBars["Ribbon"];
                ds.RibbonHeight = ribbon.Height;

                if (ds.WindowViewType == (int)Word.WdViewType.wdNormalView) ds.DraftView = true;

                SetDejaviewToDocument(doc, ds);

                SetButtonTip(ds);
            }
            catch (Exception ex)
            {
                DisplayStatus("Could not save document view. " + ex.Message);
            }
        }

        /// <summary>
        /// Used to save custom Dejaview data to the Word document. 
        /// This is used for preserving viewing settings per document.
        /// </summary>
        private void SetDejaviewToDocument(Word.Document doc, DejaviewSet ds)
        {
            // First check to see if a Dejaview CustomXMLParts already exists.
            // If so, delete it. This is the way that Microsoft updates
            // the Custom XML parts.
            try
            {
                var xml = doc.CustomXMLParts["Dejaview"];
                if (xml != null) xml.Delete();
                Debug.WriteLine("Old Dejaview tags removed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("SetDejaviewToDocument() => " + ex.StackTrace);
                DisplayStatus("Could not remove previously document view parameters.");
            }

            string xmlString =
                "<lexidata xmlns=\"Dejaview\">" +
                    "<navigation>" +
                        "<width>" + ds.NavigationPanelWidth + "</width>" +
                        "<show>" + ds.ShowNavigationPanel + "</show>" +
                    "</navigation>" +
                    "<application>" +
                        "<left>" + ds.WindowLeft + "</left>" +
                        "<top>" + ds.WindowTop + "</top>" +
                        "<width>" + ds.WindowWidth + "</width>" +
                        "<height>" + ds.WindowHeight + "</height>" +
                        "<windowstate>" + (int)ds.WindowState + "</windowstate>" +
                        "<view>" + (int)ds.WindowViewType + "</view>" +
                        "<draft>" + ds.DraftView + "</draft>" +
                        "<rulers>" + ds.DisplayRulers + "</rulers>" +
                        "<zoom>" + ds.WindowZoom + "</zoom>" +
                        "<ribbonheight>" + ds.RibbonHeight + "</ribbonheight>" +
                    "</application>" +
                "</lexidata>";
            doc.CustomXMLParts.Add(xmlString, missing);
            Debug.WriteLine("New Dejaview tags saved");

            DisplayStatus("Document view saved.");
        }

        /// <summary>
        /// Retrieve a DejaviewSet from the Word document. 
        /// This data is stored in Custom XML parts.
        /// </summary>
        /// <param name="doc">Microsoft Word Document</param>
        /// <returns>A DejaviewSet from the Word document</returns>
        public DejaviewSet GetDejaviewFromDocument(Word.Document doc)
        {
            var xml = doc.CustomXMLParts["Dejaview"];
            if (xml == null) return null;
            DejaviewSet ds = new DejaviewSet();
            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(xml.XML);
            var n = xdoc.SelectSingleNode("//*[local-name()='lexidata']");
            if (n == null) return null;

            var nodes = n.ChildNodes;
            foreach (XmlNode x in nodes)
            {
                if (x.Name == "navigation")
                {
                    var c = x.ChildNodes;
                    foreach (XmlNode _c in c)
                    {
                        if (_c.Name == "width" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.NavigationPanelWidth = int.Parse(_c.InnerText);
                        else if (_c.Name == "show" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.ShowNavigationPanel = bool.Parse(_c.InnerText);
                    }
                }
                else if (x.Name == "application")
                {
                    var c = x.ChildNodes;
                    foreach (XmlNode _c in c)
                    {
                        if (_c.Name == "left" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowLeft = int.Parse(_c.InnerText);
                        else if (_c.Name == "top" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowTop = int.Parse(_c.InnerText);
                        else if (_c.Name == "width" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowWidth = int.Parse(_c.InnerText);
                        else if (_c.Name == "height" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowHeight = int.Parse(_c.InnerText);
                        else if (_c.Name == "windowstate" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowState = int.Parse(_c.InnerText);
                        else if (_c.Name == "view" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowViewType = int.Parse(_c.InnerText);
                        else if (_c.Name == "draft" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.DraftView = bool.Parse(_c.InnerText);
                        else if (_c.Name == "rulers" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.DisplayRulers = bool.Parse(_c.InnerText);
                        else if (_c.Name == "zoom" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.WindowZoom = int.Parse(_c.InnerText);
                        else if (_c.Name == "ribbonheight" && !string.IsNullOrEmpty(_c.InnerText))
                            ds.RibbonHeight = int.Parse(_c.InnerText);
                    }
                }
            }
            return ds;
        }

        /// <summary>
        /// Test to see if this Point is showable in the current user's display arrangement.
        /// </summary>
        /// <param name="p">Point to test</param>
        /// <returns></returns>
        private static bool IsShowable(Point p)
        {
            if (p.IsEmpty) return false;
            foreach (Screen scrn in Screen.AllScreens)
            {
                //MessageBox.Show("Screen: " + scrn.Bounds + " Form: " + p);
                if (scrn.Bounds.Contains(p)) return true;
            }
            return false;
        }

        /// <summary>
        /// Show a given Form ensuring that it is showable. If no Point is provided,
        /// the Form will be centered on the primary screen.
        /// </summary>
        /// <param name="app">Word application</param>
        /// <param name="p">Point where Form should show</param>
        private static void SetShowable(Word.Application app, DejaviewSet ds)
        {
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;

            if (ds.WindowWidth < 100) ds.WindowWidth = (int)(wa.Width - Math.Round(wa.Width * 0.3));
            if (ds.WindowHeight < 100) ds.WindowHeight = (int)(wa.Height - Math.Round(wa.Height * 0.2));

            app.Width = ds.WindowWidth;
            app.Height = ds.WindowHeight;

            Point p = new Point(ds.WindowLeft, ds.WindowTop);
            // If showable, simply set coordinates
            if (!p.IsEmpty && IsShowable(p))
            {
                app.Left = p.X;
                app.Top = p.Y;
            }
            // Otherwise center on primary screen
            else
            {
                app.Left = (wa.Width - app.Width) / 2;
                app.Top = (wa.Height - app.Height) / 2;
            }
        }

        /// <summary>
        /// Helper method for displaying status to MS Word's status bar.
        /// </summary>
        /// <param name="txt">Text to display in Word's status bar</param>
        internal void DisplayStatus(string txt)
        {
            this.Application.StatusBar = txt;
        }

        /// <summary>
        /// Set tooltip for the Remove Button with details from a DejaviewSet
        /// </summary>
        /// <param name="ds"></param>
        internal void SetButtonTip(DejaviewSet ds)
        {
            // Set tooltip for the Remove Button with details from this DejaviewSet
            string tip = "Remove all Deja View tags from this document.\n\nWindow State: ";
            switch ((Word.WdWindowState)ds.WindowState)
            {
                case Word.WdWindowState.wdWindowStateNormal:
                    tip += "Normal";
                    break;
                case Word.WdWindowState.wdWindowStateMinimize:
                    tip += "Minimized";
                    break;
                case Word.WdWindowState.wdWindowStateMaximize:
                    tip += "Maximized";
                    break;
                default:
                    tip += "Unknown: " + ds.WindowState;
                    break;
            }
            if (ds.WindowState == 0)
            {
                tip += "\n   Window Left: " + ds.WindowLeft;
                tip += "\n   Window Top: " + ds.WindowTop;
                tip += "\n   Window Width: " + ds.WindowWidth;
                tip += "\n   Window Height: " + ds.WindowHeight;
            }
            tip += "\nDocument View: ";
            switch (ds.WindowViewType)
            {
                case 1:
                    tip += "Normal";
                    break;
                case 2:
                    tip += "Outline";
                    break;
                case 3:
                    tip += "Print Preview";
                    break;
                case 4:
                    tip += "Print";
                    break;
                case 6:
                    tip += "Web";
                    break;
                case 7:
                    tip += "Reading";
                    break;
                default:
                    tip += "Unknown: " + ds.WindowViewType;
                    break;
            }
            tip += "\n   Zoom: " + ds.WindowZoom;
            tip += "\nView Rulers: " + ds.DisplayRulers;
            tip += "\nView Navigation: " + ds.ShowNavigationPanel;
            tip += "\n   Width: " + ds.NavigationPanelWidth;
            tip += "\nRibbon Height: " + ds.RibbonHeight;

            Microsoft.Office.Tools.Ribbon.RibbonButton rb = Globals.Ribbons.DejaviewRibbon.btnRemove;
            rb.Enabled = true;
            rb.SuperTip = tip;
        }

        /// <summary>
        /// Thread compatible method invoked to check for updates.
        /// </summary>
        public static void CheckForUpdate()
        {
            using (var client = new WebClient())
            {
                Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

                string url = DejaviewConfig.Instance.UpdateURL;
                client.Headers.Add("User-Agent", "Deja View Update Client " + lVersion.ToString());
                string str = client.DownloadString(url + "check.php").Trim();
                int dash = str.IndexOf('-') + 1;
                str = str.Substring(dash, str.Length - dash - 3) + "0";

                Version rVersion = Version.Parse(str);

                int r = lVersion.CompareTo(rVersion);
                if (r < 0)
                {
                    DialogResult dr = MessageBox.Show(null, "An update is available for Deja View.\n\nDo you want to update it?", "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        Process.Start(url);
                    }
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
            this.Application.DocumentOpen += DejaviewAddIn_DocumentOpen;
            this.Application.DocumentBeforeSave += DejaviewAddIn_DocumentBeforeSave;
        }

        #endregion
    }
}
