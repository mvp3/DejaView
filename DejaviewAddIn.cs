﻿/**
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
using System.Diagnostics;
using System.Drawing;
using System.Security.Cryptography;
using System.Net;
using System.Threading;
using System.Windows.Forms;
using System.Xml;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;

namespace Dejaview
{
    /// <summary>
    /// Main Add-in class.
    /// </summary>
    public partial class DejaviewAddIn
    {
        /// <summary>
        /// Private member that retains the dataset of view settings from 
        /// documents that have already been viewed. This is especially 
        /// needful to retain view settings from views on other 
        /// screens.
        /// </summary>
        private DejaviewSet djvSet = null;

        /// <summary>
        /// Private member that indicates if Deja View tags were loaded from 
        /// the current active document.
        /// </summary>
        private bool loaded = false;

        /// <summary>
        /// Called when this Add-in is initialized.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <seealso cref="InternalStartup"/>
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

        /// <summary>
        /// Called when this Add-in is closing.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <seealso cref="InternalStartup"/>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// Method called when a document is opened.
        /// </summary>
        /// <param name="doc">Active Word document that opened</param>
        internal void DejaviewAddIn_DocumentOpen(Word.Document doc)
        {
            if (!DejaviewConfig.Instance.Enable) return;
            try
            {
                djvSet = GetDejaviewFromDocument(doc);

                doc.ActiveWindow.WindowState = (Word.WdWindowState)djvSet.WindowState;
                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateMinimize)
                    doc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;

                if (DejaviewConfig.Instance.RememberWindowLocation)
                {
                    if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateNormal)
                        SetShowable(doc.Application, djvSet);
                }

                if (DejaviewConfig.Instance.RememberNavigationPanel)
                {
                    Office.CommandBar nav = Application.CommandBars["Navigation"];
                    nav.Visible = djvSet.ShowNavigationPanel;
                    nav.Width = djvSet.NavigationPanelWidth;
                }

                if (DejaviewConfig.Instance.RememberWindowType)
                    doc.ActiveWindow.View.Type = (Word.WdViewType)djvSet.WindowViewType;
                if (DejaviewConfig.Instance.RememberZoom)
                    doc.ActiveWindow.View.Zoom.Percentage = djvSet.WindowZoom;
                if (DejaviewConfig.Instance.RememberRulers)
                    doc.ActiveWindow.DisplayRulers = djvSet.DisplayRulers;

                if (DejaviewConfig.Instance.RememberRibbon)
                {
                    Office.CommandBar ribbon = Application.CommandBars["Ribbon"];
                    if (djvSet.RibbonHeight != 0 && djvSet.RibbonHeight != ribbon.Height)
                    {
                        // Use 100 as a toggle threshold instead of simply checking
                        // the saved hieght against the present display height.
                        // If the document is opened on another computer with
                        // a different screen resolution a simple toggle will
                        // misbehave.
                        if ((djvSet.RibbonHeight > 100 && ribbon.Height < 100) ||
                            (djvSet.RibbonHeight < 100 && ribbon.Height > 100))
                            doc.ActiveWindow.ToggleRibbon();
                    }
                }
                loaded = true;

                DisplayStatus("Document view restored.");

                SetButtonTip();
            }
            catch (NullReferenceException)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
            }
            catch (IndexOutOfRangeException)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
            }
            catch (Exception ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                DisplayStatus("Could not restore document view. " + ex.Message);
            }
        }

        /// <summary>
        /// This methods is called immediately before Microsoft Word saves the active document.
        /// Here Dejaview gathers the current view parameters and sends them to 
        /// <code>SetDejaviewToDocument</code> be saved in the document.
        /// </summary>
        /// <param name="doc">Active Word document to be saved</param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        /// <seealso cref="SetDejaviewToDocument"/>
        private void DejaviewAddIn_DocumentBeforeSave(Word.Document doc, ref bool saveAsUI, ref bool cancel)
        {
            if (!DejaviewConfig.Instance.Enable) return;
            try
            {
                bool save = true;
                if (DejaviewConfig.Instance.Prompt)
                {
                    DialogResult r = MessageBox.Show(null, "Do you want to save this document's view settings?", "Save View?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    save = (r == DialogResult.Yes);
                }
                if (!save) return;

                if (djvSet == null) djvSet = new DejaviewSet();

                djvSet.WindowState = (int)Application.WindowState;
                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateNormal)
                {
                    djvSet.WindowWidth = doc.ActiveWindow.Width;
                    djvSet.WindowHeight = doc.ActiveWindow.Height;
                    djvSet.WindowLeft = doc.ActiveWindow.Left;
                    djvSet.WindowTop = doc.ActiveWindow.Top;

                    DejaviewSet.WindowLocation wloc = new DejaviewSet.WindowLocation();
                    wloc.DisplayArrangementUID = GetDisplayArrangementUID();
                    wloc.ScreenUID = GetActiveScreenUID();
                    wloc.WindowTop = djvSet.WindowTop;
                    wloc.WindowLeft = djvSet.WindowLeft;
                    wloc.LastViewed = DateTime.Now;

                    List<DejaviewSet.WindowLocation> locList = new List<DejaviewSet.WindowLocation>();
                    if (djvSet.Locations != null)
                    {
                        bool exists = false;
                        locList.AddRange(djvSet.Locations);
                        foreach (DejaviewSet.WindowLocation _wloc in djvSet.Locations)
                        {
                            if (_wloc.ScreenUID == wloc.ScreenUID)
                            {
                                exists = true;
                                _wloc.WindowTop = wloc.WindowTop;
                                _wloc.WindowLeft = wloc.WindowLeft;
                                _wloc.LastViewed = DateTime.Now;
                            }
                        }
                        if (!exists) locList.Add(wloc);
                    }
                    else
                    {
                        locList.Add(wloc);
                    }
                    djvSet.Locations = locList.ToArray();
                }
                djvSet.WindowViewType = (int)doc.ActiveWindow.View.Type;
                if (djvSet.WindowViewType == (int)Word.WdViewType.wdNormalView) djvSet.DraftView = true;

                djvSet.WindowZoom = doc.ActiveWindow.View.Zoom.Percentage;
                djvSet.DisplayRulers = doc.ActiveWindow.DisplayRulers;

                Office.CommandBar nav = doc.CommandBars["Navigation"];
                djvSet.ShowNavigationPanel = nav.Visible;
                djvSet.NavigationPanelWidth = nav.Width;

                Office.CommandBar ribbon = doc.CommandBars["Ribbon"];
                djvSet.RibbonHeight = ribbon.Height;

                if (djvSet.WindowViewType == (int)Word.WdViewType.wdNormalView) djvSet.DraftView = true;

                SetDejaviewToDocument(doc);

                SetButtonTip();
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
        /// <param name="doc">Microsoft Word Document</param>
        private void SetDejaviewToDocument(Word.Document doc)
        {
            // First check to see if a Dejaview CustomXMLParts already exists.
            // If so, delete it. This is the way that Microsoft updates
            // the Custom XML parts.
            try
            {
                var cp = doc.CustomXMLParts["Dejaview"];
                if (cp != null) cp.Delete();
                Debug.WriteLine("Old Dejaview tags removed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("SetDejaviewToDocument() => " + ex.StackTrace);
                DisplayStatus("Could not remove previously document view parameters.");
            }

            StringBuilder xml = new StringBuilder("<lexidata xmlns=\"Dejaview\">", 1024);
            xml.Append("<navigation>");
            xml.Append("<width>");
            xml.Append(djvSet.NavigationPanelWidth);
            xml.Append("</width>");
            xml.Append("<show>");
            xml.Append(djvSet.ShowNavigationPanel);
            xml.Append("</show>");
            xml.Append("</navigation>");
            xml.Append("<application>");
            xml.Append("<width>");
            xml.Append(djvSet.WindowWidth);
            xml.Append("</width>");
            xml.Append("<height>");
            xml.Append(djvSet.WindowHeight);
            xml.Append("</height>");
            xml.Append("<windowstate>");
            xml.Append((int)djvSet.WindowState);
            xml.Append("</windowstate>");
            xml.Append("<view>");
            xml.Append((int)djvSet.WindowViewType);
            xml.Append("</view>");
            xml.Append("<draft>");
            xml.Append(djvSet.DraftView);
            xml.Append("</draft>");
            xml.Append("<rulers>");
            xml.Append(djvSet.DisplayRulers);
            xml.Append("</rulers>");
            xml.Append("<zoom>");
            xml.Append(djvSet.WindowZoom);
            xml.Append("</zoom>");
            xml.Append("<ribbonheight>");
            xml.Append(djvSet.RibbonHeight);
            xml.Append("</ribbonheight>");

            if (djvSet.Locations != null)
            {
                xml.Append("<locations>");
                foreach (DejaviewSet.WindowLocation loc in djvSet.Locations)
                {
                    xml.Append("<location>");
                    xml.Append("<dauid>");
                    xml.Append(loc.DisplayArrangementUID);
                    xml.Append("</dauid>");
                    xml.Append("<uid>");
                    xml.Append(loc.ScreenUID);
                    xml.Append("</uid>");
                    xml.Append("<top>");
                    xml.Append(loc.WindowTop);
                    xml.Append("</top>");
                    xml.Append("<left>");
                    xml.Append(loc.WindowLeft);
                    xml.Append("</left>");
                    xml.Append("<ts>");
                    xml.Append(loc.LastViewed);
                    xml.Append("</ts>");
                    xml.Append("</location>");
                }
                xml.Append("</locations>");
            }

            xml.Append("</application>");
            xml.Append("</lexidata>");

            Debug.WriteLine("**********************************");
            Debug.WriteLine(xml.ToString());
            Debug.WriteLine("**********************************");

            doc.CustomXMLParts.Add(xml.ToString(), missing);
            Debug.WriteLine("New Dejaview tags saved");

            DisplayStatus("Document view saved.");
        }

        /// <summary>
        /// Status of whether or not Deja View tags were loaded from the current active document.
        /// </summary>
        /// <returns>Return <code>true</code> if Deja View tags were loaded from this document, 
        /// otherwise <code>false</code></returns>
        internal bool IsLoaded()
        {
            return loaded;
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

            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(xml.XML);

            var root = xdoc.SelectSingleNode("//*[local-name()='lexidata']");
            if (root == null) return null;

            Debug.WriteLine("**********************************");
            Debug.WriteLine(root.OuterXml);
            Debug.WriteLine("**********************************");

            djvSet = new DejaviewSet();

            var categories = root.ChildNodes;
            foreach (XmlNode x in categories)
            {
                if (x.Name == "navigation")
                {
                    var nodes = x.ChildNodes;
                    foreach (XmlNode n in nodes)
                    {
                        if (n.Name == "width" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.NavigationPanelWidth = int.Parse(n.InnerText);
                        else if (n.Name == "show" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.ShowNavigationPanel = bool.Parse(n.InnerText);
                    }
                }
                else if (x.Name == "application")
                {
                    var nodes = x.ChildNodes;
                    foreach (XmlNode n in nodes)
                    {
                        if (n.Name == "locations" && !string.IsNullOrEmpty(n.InnerText))
                        {
                            var locs = n.ChildNodes;
                            List<DejaviewSet.WindowLocation> locations = new List<DejaviewSet.WindowLocation>();
                            foreach (XmlNode loc in locs)
                            {
                                if (loc.Name == "location" && !string.IsNullOrEmpty(n.InnerText))
                                {
                                    DejaviewSet.WindowLocation wl = new DejaviewSet.WindowLocation();
                                    var locationNodes = loc.ChildNodes;
                                    foreach (XmlNode nl in locationNodes)
                                    {
                                        if (nl.Name == "dauid" && !string.IsNullOrEmpty(nl.InnerText))
                                            wl.DisplayArrangementUID = nl.InnerText;
                                        else if (nl.Name == "uid" && !string.IsNullOrEmpty(nl.InnerText))
                                            wl.ScreenUID = nl.InnerText;
                                        else if (nl.Name == "top" && !string.IsNullOrEmpty(nl.InnerText))
                                            wl.WindowTop = int.Parse(nl.InnerText);
                                        else if (nl.Name == "left" && !string.IsNullOrEmpty(nl.InnerText))
                                            wl.WindowLeft = int.Parse(nl.InnerText);
                                        else if (nl.Name == "ts" && !string.IsNullOrEmpty(nl.InnerText))
                                            wl.LastViewed = DateTime.Parse(nl.InnerText);
                                    }
                                    locations.Add(wl);
                                }
                            }
                            // Sort window locations by time in descending order
                            locations.Sort();
                            locations.Reverse();
                            djvSet.Locations = locations.ToArray();
                        }
                        else if (n.Name == "width" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowWidth = int.Parse(n.InnerText);
                        else if (n.Name == "height" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowHeight = int.Parse(n.InnerText);
                        else if (n.Name == "windowstate" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowState = int.Parse(n.InnerText);
                        else if (n.Name == "view" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowViewType = int.Parse(n.InnerText);
                        else if (n.Name == "draft" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.DraftView = bool.Parse(n.InnerText);
                        else if (n.Name == "rulers" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.DisplayRulers = bool.Parse(n.InnerText);
                        else if (n.Name == "zoom" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowZoom = int.Parse(n.InnerText);
                        else if (n.Name == "ribbonheight" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.RibbonHeight = int.Parse(n.InnerText);
                    }
                }
            }
            return djvSet;
        }

        /// <summary>
        /// Remove the DejaviewSet from this Word document. This data is stored in Custom XML parts.
        /// This method will also reset the DejaviewSet private member (<code>djvSet</code>).
        /// </summary>
        /// <param name="doc">Microsoft Word Document</param>
        public void RemoveDejaviewFromDocument(Word.Document doc)
        {
            try
            {
                djvSet = null;
                var xml = doc.CustomXMLParts["Dejaview"];
                if (xml != null) xml.Delete();
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewRibbon::btnClear_Click() => " + ex.StackTrace);
                DisplayStatus("Could not clear tags: " + ex.Message);
            }
        }

        /// <summary>
        /// Show a given Form ensuring that it is showable. If no Point is provided,
        /// the Form will be centered on the primary screen.
        /// </summary>
        /// <param name="app">Word application</param>
        /// <param name="ds">DejaviewSet for the active document</param>
        private static void SetShowable(Word.Application app, DejaviewSet ds)
        {
            // Setup key variables
            Screen screen = Screen.FromPoint(new Point(app.Left, app.Top));
            DejaviewSet.WindowLocation windowLocation = null;

            // Make sure a valid DejaviewSet is stored in this document
            if (ds.Locations != null)
            {
                bool matched = false;
                DateTime latest = DateTime.MinValue;

                string dauid = GetDisplayArrangementUID();

                // First try to match a Display Arrangement
                foreach (DejaviewSet.WindowLocation wl in ds.Locations)
                {
                    if (dauid == wl.DisplayArrangementUID)
                    {
                        latest = wl.LastViewed;
                        windowLocation = wl;
                        foreach (Screen scrn in Screen.AllScreens)
                        {
                            if (wl.ScreenUID == GetScreenUID(scrn))
                            {
                                screen = scrn;
                                break;
                            }
                        }
                        matched = true;
                        break;
                    }
                }

                // If a Display Arrangement is not matched, then try to set the document to
                // display on a screen that has the same ScreenUID
                if (!matched)
                {
                    foreach (Screen scrn in Screen.AllScreens)
                    {
                        Debug.WriteLine(" [Screen: " + GetScreenUID(scrn) + "]");
                        foreach (DejaviewSet.WindowLocation wl in ds.Locations)
                        {
                            Debug.WriteLine("   [WL: " + wl.ScreenUID + ", " + wl.LastViewed + "]");
                            if (wl.ScreenUID == GetScreenUID(scrn))
                            {
                                Debug.WriteLine("     *SELECTED*");
                                latest = wl.LastViewed;
                                screen = scrn;
                                windowLocation = wl;
                                matched = true;
                                break;
                            }
                        }
                        if (matched) break;
                    }
                }
            }

            Rectangle workingArea = screen.WorkingArea;

            // If for some reason the height or width of the window is less 
            // than 100 pixels, then set it to a proportionate size.
            if (ds.WindowWidth < 100) ds.WindowWidth = (int)(workingArea.Width - Math.Round(workingArea.Width * 0.3));
            if (ds.WindowHeight < 100) ds.WindowHeight = (int)(workingArea.Height - Math.Round(workingArea.Height * 0.2));
            app.Width = ds.WindowWidth;
            app.Height = ds.WindowHeight;

            // If 'Window Location' option is not selected then do not set window location
            if (!DejaviewConfig.Instance.RememberWindowLocation) return;

            Debug.WriteLine("     [Left: " + ds.WindowLeft + "]");
            Debug.WriteLine("     [ Top: " + ds.WindowTop + "]");

            // If remembered, simply set coordinates
            if (windowLocation != null)
            {
                Debug.WriteLine("     >> setting location");
                app.Left = windowLocation.WindowLeft;
                app.Top = windowLocation.WindowTop;
            }
            else
            {
                Debug.WriteLine("     >> constructing new location");
                // If Word window is not viewable, then center on current screen.
                if (!workingArea.Contains(new Point(app.Left, app.Top)))
                {
                    // Adjust if application window is larger than working area of the current screen
                    if (app.Width > workingArea.Width) app.Width = (int)(workingArea.Width - Math.Round(workingArea.Width * 0.3));
                    if (app.Height > workingArea.Height) app.Height = (int)(workingArea.Height - Math.Round(workingArea.Height * 0.2));

                    // Baisc adjustments for DPI differences
                    float dpiAdjust = 1;
                    using (Graphics graphics = Graphics.FromHwnd(Process.GetCurrentProcess().MainWindowHandle))
                    {
                        float res = 72;
                        float dpi = graphics.DpiX;
                        dpiAdjust = res / dpi;
                    }
                    app.Left = (int)((workingArea.Width / 2) * dpiAdjust - (app.Width / 2));
                    app.Top = (int)((workingArea.Height / 2) * dpiAdjust - (app.Height / 2));
                }
                // Otherwise let Word display as normal.
            }
        }

        /// <summary>
        /// Returns a unique identifier (in hash format) uniquely identifying
        /// the current Display Arrangement of the computer.
        /// </summary>
        /// <returns>The current Display Arrangement of the computer</returns>
        internal static string GetDisplayArrangementUID()
        {
            StringBuilder str = new StringBuilder();
            foreach (Screen scr in Screen.AllScreens)
            {
                str.Append(GetScreenUID(scr));
            }
            return GetHashCode(str.ToString());
        }

        /// <summary>
        /// Returns a SHA1 hashcode for the string provided.
        /// </summary>
        /// <param name="input">Input string for the hashcode generator</param>
        /// <returns>SHA1 hashcode for <code>input</code></returns>
        internal static string GetHashCode(string input)
        {
            using (var algo = new SHA1Managed())
            {
                algo.ComputeHash(Encoding.UTF8.GetBytes(input));

                // Get has value in array of bytes
                var result = algo.Hash;

                // Return as hexadecimal string
                return string.Join(
                    string.Empty,
                    result.Select(x => x.ToString("x2")));
            }
        }

        /// <summary>
        /// Get the Unique Identifier of the Screen displaying the active document.
        /// </summary>
        /// <returns>Screen UID</returns>
        /// <seealso cref="GetScreenUID(Screen)"/>
        internal string GetActiveScreenUID()
        {
            int x = this.Application.Left;
            int y = this.Application.Top;
            return GetScreenUID(Screen.FromPoint(new Point(x, y)));
        }

        /// <summary>
        /// Get the Unique Identifier of the Screen specified. This ID is comprised
        /// of the system assigned device name concatenated with the string representation
        /// of the screen's working area.
        /// </summary>
        /// <param name="scr">Screen to identify</param>
        /// <returns>Screen UID</returns>
        internal static string GetScreenUID(Screen scr)
        {
            return scr.DeviceName + scr.WorkingArea.ToString();
        }

        /// <summary>
        /// Returns a human readable name for the screen based on the screen's UID.
        /// 
        /// For example, a UID is in the form: '\\.\DISPLAY1{X=-1920,Y=0,Width=1920,Height=1050}'
        /// The name for this screen would be: 'Display 1 (1920 x 1050) [Left]'
        /// </summary>
        /// <param name="uid">Screen UID</param>
        /// <returns>A human readable name</returns>
        internal static string GetScreenNameFromUID(string uid)
        {
            try
            {
                int i = uid.IndexOf('{');
                string n = uid.Substring(i - 1, 1);
                string b = uid.Substring(i);

                var m = Regex.Match(b, @"{X=(\d+|-\d+),\s*Y=(\d+|-\d+),\s*Width=(\d+),Height=(\d+)}");

                StringBuilder str = new StringBuilder("Display ");
                str.Append(n);
                str.Append(" (");
                str.Append(m.Groups[3].Value);
                str.Append(" x ");
                str.Append(m.Groups[4].Value);
                str.Append(")");
                if (int.Parse(m.Groups[1].Value) < 0) str.Append(" [Left]");

                return str.ToString();
            }
            catch (Exception ex)
            {
                return uid;
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
        internal void SetButtonTip()
        {
            Microsoft.Office.Tools.Ribbon.RibbonButton rb = Globals.Ribbons.DejaviewRibbon.btnRemove;
            rb.Enabled = true;
            rb.SuperTip = "Remove all Deja View tags from this document.";
        }

        /// <summary>
        /// Returns a string with Dejaview parameter details from the current application window 
        /// </summary>
        /// <returns>A string with Dejaview parameter details from the current application window</returns>
        internal string GetCurrentParameters()
        {
            // Build a string with details from the current application window
            StringBuilder str = new StringBuilder("Window State: \t", 1024);
            Word.Document doc = this.Application.ActiveDocument;
            Word.WdWindowState ws = (Word.WdWindowState)this.Application.ActiveWindow.WindowState;
            switch (ws)
            {
                case Word.WdWindowState.wdWindowStateNormal:
                    str.AppendLine("Normal");
                    break;
                case Word.WdWindowState.wdWindowStateMinimize:
                    str.AppendLine("Minimized");
                    break;
                case Word.WdWindowState.wdWindowStateMaximize:
                    str.AppendLine("Maximized");
                    break;
                default:
                    str.Append("Unknown: ");
                    str.Append(ws);
                    str.AppendLine();
                    break;
            }
            if (ws == 0)
            {
                str.Append("  Window Width: \t");
                str.Append(this.Application.ActiveWindow.Width);
                str.AppendLine();
                str.Append("  Window Height: \t");
                str.Append(this.Application.ActiveWindow.Height);
                str.AppendLine();
            }
            str.AppendLine();

            str.Append("Document View: \t");
            switch ((int)this.Application.ActiveWindow.View.Type)
            {
                case 1:
                    str.Append("Normal");
                    break;
                case 2:
                    str.Append("Outline");
                    break;
                case 3:
                    str.Append("Print Preview");
                    break;
                case 4:
                    str.Append("Print");
                    break;
                case 6:
                    str.Append("Web");
                    break;
                case 7:
                    str.Append("Reading");
                    break;
                default:
                    str.Append("Unknown: ");
                    str.Append((int)this.Application.ActiveWindow.View.Type);
                    break;
            }
            str.AppendLine();

            Office.CommandBar nav = doc.CommandBars["Navigation"];
            djvSet.ShowNavigationPanel = nav.Visible;
            djvSet.NavigationPanelWidth = nav.Width;

            Office.CommandBar ribbon = doc.CommandBars["Ribbon"];
            djvSet.RibbonHeight = ribbon.Height;

            str.Append("   Zoom: \t\t");
            str.Append(this.Application.ActiveWindow.View.Zoom.Percentage);
            str.Append("%");
            str.AppendLine();
            str.Append("View Rulers: \t");
            str.Append(this.Application.ActiveWindow.DisplayRulers);
            str.AppendLine();
            str.Append("View Navigation: \t");
            str.Append(nav.Visible);
            str.AppendLine();
            str.Append("   Width: \t\t");
            str.Append(nav.Width);
            str.AppendLine();
            str.Append("Ribbon Height: \t");
            str.Append(ribbon.Height);
            str.AppendLine();
            str.AppendLine();

            DejaviewSet.WindowLocation wloc = new DejaviewSet.WindowLocation();
            wloc.DisplayArrangementUID = GetDisplayArrangementUID();
            wloc.ScreenUID = GetActiveScreenUID();
            wloc.WindowTop = this.Application.ActiveWindow.Top;
            wloc.WindowLeft = this.Application.ActiveWindow.Left;
            wloc.LastViewed = DateTime.Now;

            str.AppendLine("Window Location: ");
            str.Append("   Display Arrangement: ");
            str.Append(wloc.DisplayArrangementUID);
            str.AppendLine();
            str.Append("   Screen: ");
            str.Append(GetScreenNameFromUID(wloc.ScreenUID));
            str.AppendLine();
            str.Append("   \tTop: \t");
            str.Append(wloc.WindowTop);
            str.AppendLine();
            str.Append("   \tLeft: \t");
            str.Append(wloc.WindowLeft);
            str.AppendLine();
            str.Append("   \tLast viewed: ");
            str.Append(wloc.LastViewed);
            str.AppendLine();

            return str.ToString();
        }

        /// <summary>
        /// Returns a string with Dejaview parameter details from a DejaviewSet
        /// </summary>
        /// <returns>A string with Dejaview parameter details from a DejaviewSet</returns>
        internal string GetSavedTags()
        {
            // Build a string with details from a DejaviewSet
            StringBuilder str = new StringBuilder("Window State: \t", 1024);
            switch ((Word.WdWindowState)djvSet.WindowState)
            {
                case Word.WdWindowState.wdWindowStateNormal:
                    str.Append("Normal");
                    break;
                case Word.WdWindowState.wdWindowStateMinimize:
                    str.Append("Minimized");
                    break;
                case Word.WdWindowState.wdWindowStateMaximize:
                    str.Append("Maximized");
                    break;
                default:
                    str.Append("Unknown: ");
                    str.Append(djvSet.WindowState);
                    break;
            }
            str.AppendLine();
            if (djvSet.WindowState == 0)
            {
                str.Append("   Window Width: \t");
                str.Append(djvSet.WindowWidth);
                str.AppendLine();
                str.Append("   Window Height: \t");
                str.Append(djvSet.WindowHeight);
                str.AppendLine();
            }
            str.AppendLine();

            str.Append("Document View: \t");
            switch (djvSet.WindowViewType)
            {
                case 1:
                    str.Append("Normal");
                    break;
                case 2:
                    str.Append("Outline");
                    break;
                case 3:
                    str.Append("Print Preview");
                    break;
                case 4:
                    str.Append("Print");
                    break;
                case 6:
                    str.Append("Web");
                    break;
                case 7:
                    str.Append("Reading");
                    break;
                default:
                    str.Append("Unknown: ");
                    str.Append(djvSet.WindowViewType);
                    break;
            }
            str.AppendLine();
            str.Append("   Zoom: \t\t");
            str.Append(djvSet.WindowZoom);
            str.Append("%");
            str.AppendLine();
            str.Append("View Rulers: \t");
            str.Append(djvSet.DisplayRulers);
            str.AppendLine();
            str.Append("View Navigation: \t");
            str.Append(djvSet.ShowNavigationPanel);
            str.AppendLine();
            str.Append("   Width: \t\t");
            str.Append(djvSet.NavigationPanelWidth);
            str.AppendLine();
            str.Append("Ribbon Height: \t");
            str.Append(djvSet.RibbonHeight);
            str.AppendLine();
            str.AppendLine();

            string sid = GetActiveScreenUID();
            str.AppendLine("Window Location: ");
            foreach (DejaviewSet.WindowLocation wl in djvSet.Locations)
            {
                str.Append((sid == wl.ScreenUID) ? "*" : " ");
                str.Append("   Display Arrangement: ");
                str.Append(wl.DisplayArrangementUID);
                str.AppendLine();
                str.Append("   Screen: ");
                str.Append(GetScreenNameFromUID(wl.ScreenUID));
                str.AppendLine();
                str.Append("   \tTop: \t");
                str.Append(wl.WindowTop);
                str.AppendLine();
                str.Append("   \tLeft: \t");
                str.Append(wl.WindowLeft);
                str.AppendLine();
                str.Append("   \tLast viewed: ");
                str.Append(wl.LastViewed);
                str.AppendLine();
            }

            return str.ToString();
        }

        /// <summary>
        /// Thread compatible method invoked to check for updates.
        /// </summary>
        public static void CheckForUpdate(object silent = null)
        {
            try
            {
                using (var client = new WebClient())
                {
                    Version lVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                    Debug.WriteLine("CheckForUpdate() lVersion => " + lVersion);

                    string url = DejaviewConfig.Instance.UpdateURL;
                    client.Headers.Add("User-Agent", "Deja View Update Client " + lVersion.ToString());
                    string str = client.DownloadString(url).Trim();
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
                    else if (silent != null && silent is bool && !(bool)silent)
                    {
                        MessageBox.Show(null, "You are using the latest version of Deja View.", "Up to Date", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("CheckForUpdate() => " + ex.Message);
                if (silent != null && silent is bool && !(bool)silent)
                {
                    MessageBox.Show(null, "An error occurred while checking for an update:\n\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
