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
using System.Collections;

namespace Dejaview
{
    /// <summary>
    /// Main Add-in class.
    /// </summary>
    public partial class DejaviewAddIn
    {
        /// <summary>
        /// Collection that retains the dataset of view settings from 
        /// documents that have already been viewed. This is especially 
        /// needful to retain view settings from views on other 
        /// screens.
        /// </summary>
        private Hashtable djvSets = new Hashtable();

        /// <summary>
        /// Collection of unique Logger instances, each keyed to its own 
        /// ActiveDocument. This is necessary since Microsoft Word
        /// uses a shared instance of this Add-in across all open documents.
        /// </summary>
        private Hashtable loggers = new Hashtable();

        /// <summary>
        /// Collection that retains what documents have had their view set 
        /// by DV. Each element in this Collection is a bool flag keyed to its own
        /// ActiveDocument, indicating if that document's view settings have been 
        /// applied. This is necessary since Microsoft Word does not reliably 
        /// fire the DocumentOpen event when a document is opened.
        /// </summary>
        private Hashtable views = new Hashtable();

        /// <summary>
        /// Collection that retains what documents have been saved. This is used
        /// for saving the very last view before closing the document.
        /// </summary>
        private ArrayList saves = new ArrayList();

        /// <summary>
        /// Used by the InitializeCustom() method to prevent registering duplicate 
        /// event handlers.
        /// </summary>
        private bool initialized = false;

        /// <summary>
        /// Called when this Add-in is initialized.
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event arguments</param>
        /// <seealso cref="InternalStartup"/>
        private void DejaviewAddIn_Startup(object sender, EventArgs e)
        {
            //new InfoForm("DejaView started " + DateTime.Now, Application.ActiveDocument.Name);
            try
            {
                if (!DejaviewConfig.Instance.Enable) return;

                // Fire the DocumentOpen event for the first time
                DejaviewAddIn_DocumentOpen(Application.ActiveDocument);
                this.Log("DejaView loaded (" + Application.ActiveDocument.Name + "): " + DateTime.Now);
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
        /// <param name="sender">Sender object</param>
        /// <param name="e">Event arguments</param>
        /// <seealso cref="InternalStartup"/>
        private void DejaviewAddIn_Shutdown(object sender, EventArgs e)
        {
            // Nothing to do for now.
        }

        /// <summary>
        /// A trigger method to ensure that DejaView initializes a document's 
        /// settings upon opening each document.
        /// 
        /// Without this trigger, DejaView properties are not always applied 
        /// when a document opens because the VSTO plugin is already initialized.
        /// </summary>
        private void InitializeCustom()
        {
            if (!initialized)
            {
                Word.ApplicationEvents4_Event wdEvents4 = (Word.ApplicationEvents4_Event)this.Application;
                wdEvents4.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(DejaviewAddIn_DocumentOpen);
                wdEvents4.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(DejaviewAddIn_DocumentBeforeSave);
                wdEvents4.NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(DejaviewAddIn_NewDocument);
                wdEvents4.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(DejaviewAddIn_DocumentChange);
                wdEvents4.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(DejaviewAddIn_DocumentBeforeClose);
            }
            initialized = true;
        }

        /// <summary>
        /// Method called when a new document is created.
        /// </summary>
        /// <param name="doc">Active Word document newly created.</param>
        internal void DejaviewAddIn_NewDocument(Word.Document doc)
        {
            if (!DejaviewConfig.Instance.Enable) return;

            // Create a unique instance of Logger for this document.
            loggers.Add(doc.DocID, new Logger());

            // See if a default DejaviewSet is set.
            DejaviewSet s = DejaviewConfig.Instance.DefaultDejaviewSet;
            if (s != null)
            {
                SetDocumentView(doc, s);
                Log("Set new document view to the default view.");
            }
            else
            {
                Log("No default view is set for new documents.");
            }
        }

        /// <summary>
        /// Method for intercepting a close and saving the document's view settings just before
        /// the document is closed. If the document has not been saved during its current view,
        /// then Deja View will not initiate a save. If, however, it has been saved since it was opened,
        /// Deja View will save it before closing. If this logic is not in place, then every document
        /// opened for view will force a save and thereby alter the modified timestamp unecessarily.
        /// </summary>
        /// <param name="doc">Active Word document being closed</param>
        /// <param name="bCancel">Flag to interrupt the close</param>
        internal void DejaviewAddIn_DocumentBeforeClose(Word.Document doc, ref bool bCancel)
        {
            if (!DejaviewConfig.Instance.Enable) return;
            int id = doc.DocID;
            if (saves.Contains(id))
            {
                // Save before closing.
                doc.Save();
                // Clear this doc's title from the list.
                saves.Remove(id);
            }
        }

        /// <summary>
        /// Method called when a change occurs in the document.
        /// This method is necessary since opening a document that has 
        /// already been opened and closed within the same instance of MS Word
        /// does not always fire ApplicationEvents4_DocumentOpenEventHandler.
        /// </summary>
        internal void DejaviewAddIn_DocumentChange()
        {
            int key = Application.ActiveDocument.DocID;
            // If the Active document has not had its view set by DJ
            if (views.ContainsKey(key) && !(bool)views[key])
            {
                // Get Logger
                Logger logger = GetLogger();
                logger.Add("Checking views...");

                DejaviewSet s = GetDejaviewSet();
                DejaviewSet d = GetDejaviewSetFromDisplay();
                if (d.Equals(s))
                {
                    logger.Add("  Current view and saved view match, no changes.");
                }
                else
                {
                    logger.Add("  Applying saved view.");
                    SetDocumentView(Application.ActiveDocument, s);
                }
            }
        }

        /// <summary>
        /// Method called when a document is opened.
        /// </summary>
        /// <param name="doc">Active Word document that opened.</param>
        internal void DejaviewAddIn_DocumentOpen(Word.Document doc)
        {
            if (!DejaviewConfig.Instance.Enable) return;

            // Create a unique instance of Logger for this document.
            loggers.Add(doc.DocID, new Logger());

            // Create first log event as the title of the ActiveDocument window.
            Log(doc.ActiveWindow.Caption);

            try
            {
                DejaviewSet djvSet = GetDejaviewSetFromDocument(doc);
                SetDocumentView(doc, djvSet);
            }
            catch (DejaViewException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log(ex.Message);
            }
            catch (NullReferenceException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log("Error: " + ex.Message);
            }
            catch (IndexOutOfRangeException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log("Index error: " + ex.Message);
            }
            catch (Exception ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                DisplayStatus("Could not restore document view. " + ex.Message);
                new InfoForm("Could not restore document view: " + ex.Message, doc.ActiveWindow.Caption);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// This methods is called immediately before Microsoft Word saves the active document.
        /// Here Deja View gathers the current view parameters and sends them to 
        /// <code>SetDejaviewToDocument</code> be saved in the document.
        /// </summary>
        /// <param name="doc">Active Word document to be saved</param>
        /// <param name="saveAsUI"></param>
        /// <param name="cancel"></param>
        /// <seealso cref="SaveDejaviewToDocument"/>
        private void DejaviewAddIn_DocumentBeforeSave(Word.Document doc, ref bool saveAsUI, ref bool cancel)
        {
            if (!DejaviewConfig.Instance.Enable) return;

            DejaviewSet djvSetDisplay = GetDejaviewSetFromDisplay();
            DejaviewSet djvSet = GetDejaviewSet();

            // If the window has not changed then abort the save.
            if (djvSetDisplay.Equals(djvSet)) return;

            try
            {
                bool save = true;
                if (DejaviewConfig.Instance.Prompt)
                {
                    DialogResult r = MessageBox.Show(null, "Do you want to save this document's view settings?", "Save View?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    save = (r == DialogResult.Yes);
                }
                if (!save) return;

                UpdateDejaviewSet(djvSetDisplay);

                djvSet = GetDejaviewSet();

                SaveDejaviewToDocument(doc, djvSet);

                SetButtonTip();

                if (!saves.Contains(doc.DocID)) saves.Add(doc.DocID);
            }
            catch (Exception ex)
            {
                DisplayStatus("Could not save document view. " + ex.Message);
            }
        }

        /// <summary>
        /// Used to save custom Deja View data to the Word document. 
        /// This is used for preserving viewing settings per document.
        /// </summary>
        /// <param name="doc">Microsoft Word document</param>
        /// <param name="djvSet">DejaviewSet object to save to document</param>
        private void SaveDejaviewToDocument(Word.Document doc, DejaviewSet djvSet)
        {
            // Abort save if the djvSet object is empty
            if (djvSet == null || (djvSet.WindowWidth == 0 && djvSet.WindowHeight == 0)) return;

            // First check to see if a Deja View CustomXMLParts already exists.
            // If so, delete it. This is the way that Microsoft updates
            // the Custom XML parts.
            try
            {
                var cp = doc.CustomXMLParts["Dejaview"];
                try
                {
                    if (cp != null) cp.Delete();
                    Debug.WriteLine("Old Deja View tags removed");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("SetDejaviewToDocument() => " + ex.StackTrace);
                    DisplayStatus("Could not remove previously document view parameters.");
                }
            }
            catch (Exception)
            {
                Debug.WriteLine("No Deja View tags found; new document?");
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
            xml.Append("<left>");
            xml.Append(djvSet.WindowLeft);
            xml.Append("</left>");
            xml.Append("<top>");
            xml.Append(djvSet.WindowTop);
            xml.Append("</top>");
            xml.Append("<windowstate>");
            xml.Append(djvSet.WindowState);
            xml.Append("</windowstate>");
            xml.Append("<view>");
            xml.Append(djvSet.WindowViewType);
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
            Debug.WriteLine("New Deja View tags saved");

            DisplayStatus("Document view saved.");
        }

        /// <summary>
        /// Status of whether or not Deja View tags were loaded for the specified document.
        /// </summary>
        /// <param name="doc">Microsoft Word document</param>
        /// <returns>Return <code>true</code> if Deja View tags were loaded from <code>doc</code>, 
        /// otherwise <code>false</code></returns>
        internal bool IsLoaded(Word.Document doc)
        {
            if (doc == null) return false;
            return (views.ContainsKey(doc.DocID) && (bool)views[doc.DocID]);
        }

        /// <summary>
        /// Retrieve a DejaviewSet from the Word document. 
        /// This data is stored in Custom XML parts.
        /// </summary>
        /// <param name="doc">Microsoft Word document</param>
        /// <returns>A DejaviewSet from the Word document</returns>
        public DejaviewSet GetDejaviewSetFromDocument(Word.Document doc)
        {
            Office.CustomXMLPart xml = null;
            
            try
            {
                xml = doc.CustomXMLParts["Dejaview"];
            }
            catch (Exception)
            {
                throw new DejaViewNoTagsException();
            }

            if (xml == null) return null;

            XmlDocument xdoc = new XmlDocument();
            xdoc.LoadXml(xml.XML);

            var root = xdoc.SelectSingleNode("//*[local-name()='lexidata']");
            if (root == null) throw new DejaViewInvalidTagException();

            Debug.WriteLine("**********************************");
            Debug.WriteLine(root.OuterXml);
            Debug.WriteLine("**********************************");

            DejaviewSet djvSet = GetDejaviewSet();

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
                        else if (n.Name == "left" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowLeft = int.Parse(n.InnerText);
                        else if (n.Name == "top" && !string.IsNullOrEmpty(n.InnerText))
                            djvSet.WindowTop = int.Parse(n.InnerText);
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

            SetDejaviewSet(djvSet);

            return djvSet;
        }

        /// <summary>
        /// Remove the DejaviewSet from this Word document. This data is stored in Custom XML parts.
        /// This method will also reset the DejaviewSet collection (<code>djvSets</code>).
        /// </summary>
        /// <seealso cref="djvSets"/>
        /// <param name="doc">Microsoft Word document</param>
        public void RemoveDejaviewFromDocument(Word.Document doc)
        {
            try
            {
                djvSets.Remove(doc.DocID);
                var xml = doc.CustomXMLParts["Dejaview"];
                if (xml != null) xml.Delete();
                doc.Save();
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                DisplayStatus("Deja View tags removed from this document.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewRibbon::btnClear_Click() => " + ex.StackTrace);
                DisplayStatus("Could not clear tags: " + ex.Message);
            }
        }

        /// <summary>
        /// Get the unique Logger for the ActiveDocument.
        /// </summary>
        /// <returns>Logger instance</returns>
        internal Logger GetLogger()
        {
            return (Logger)loggers[Globals.DejaviewAddIn.Application.ActiveDocument.DocID];
        }

        /// <summary>
        /// Convenience method for logging an event associated with the ActiveDocument.
        /// </summary>
        /// <example>Globals.DejaviewAddIn.Log("Window restored.");</example>
        /// <param name="description">String description of the event.</param>
        internal void Log(string description)
        {
            try
            {
                GetLogger().Add(description);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewAddIn::Log(string) => " + ex.StackTrace);
            }
        }

        /// <summary>
        /// Convenience method for logging an Deja View exception associated with the ActiveDocument.
        /// </summary>
        /// <example>Globals.DejaviewAddIn.Log(new Exception("Failed."));</example>
        /// <param name="ex">Exception representing the event</param>
        internal void Log(Exception ex)
        {
            try
            {
                GetLogger().Add(ex);
            }
            catch (Exception e)
            {
                Debug.WriteLine("DejaviewAddIn::Log(exception) => " + e.StackTrace);
            }
        }

        /// <summary>
        /// This method retrieves the DejaviewSet associated with the ActiveDocument.
        /// </summary>
        /// <returns>DejaviewSet object associated with the ActiveDocument</returns>
        internal DejaviewSet GetDejaviewSet()
        {
            DejaviewSet djvSet = (DejaviewSet)djvSets[Globals.DejaviewAddIn.Application.ActiveDocument.DocID];
            return djvSet ?? new DejaviewSet();
        }

        /// <summary>
        /// This method links the provided DejaviewSet object to the ActiveDocument and then stores it
        /// in a Collection for easy retrieval. If a DejaviewSet object is already associated with the ActiveDocument
        /// that object is simply updated.
        /// </summary>
        /// <param name="djvSet">DejaviewSet object to link to the ActiveDocument</param>
        internal void SetDejaviewSet(DejaviewSet djvSet)
        {
            int id = Globals.DejaviewAddIn.Application.ActiveDocument.DocID;
            if (djvSets.Contains(id)) djvSets.Remove(id);
            djvSets.Add(id, djvSet);
        }

        /// <summary>
        /// This method creates a new DejaviewSet object and then assigns values to it based
        /// on the current ActiveWindow display parameters.
        /// </summary>
        /// <returns>DejaviewSet object containing current view parameters</returns>
        internal DejaviewSet GetDejaviewSetFromDisplay()
        {
            DejaviewSet djvSet = new DejaviewSet();
            djvSet.WindowState = (int)Application.WindowState;
            if (djvSet.WindowState != (int)Word.WdWindowState.wdWindowStateMinimize)
            {
                djvSet.WindowTop = Application.Top;
                djvSet.WindowLeft = Application.Left;
                djvSet.WindowWidth = Application.Width;
                djvSet.WindowHeight = Application.Height;
            }

            djvSet.WindowViewType = (int)Application.ActiveWindow.View.Type;
            if (djvSet.WindowViewType == (int)Word.WdViewType.wdNormalView) djvSet.DraftView = true;

            djvSet.WindowZoom = Application.ActiveWindow.View.Zoom.Percentage;
            djvSet.DisplayRulers = Application.ActiveWindow.DisplayRulers;

            djvSet.ShowNavigationPanel = Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneNav].Visible;
            djvSet.NavigationPanelWidth = Application.CommandBars["Navigation"].Width;

            Office.CommandBar ribbon = Application.CommandBars["Ribbon"];
            djvSet.RibbonHeight = ribbon.Height;

            return djvSet;
        }

        /// <summary>
        /// Updates the DejaviewSet object associated with the ActiveDocument with the
        /// values of the provided DejaviewSet object.
        /// </summary>
        /// <param name="newSet">DejaviewSet object containing new values</param>
        internal void UpdateDejaviewSet(DejaviewSet newSet)
        {
            DejaviewSet djvSet = GetDejaviewSet();
            djvSet.WindowState = newSet.WindowState;
            djvSet.WindowHeight = newSet.WindowHeight;
            djvSet.WindowLeft = newSet.WindowLeft;
            djvSet.WindowTop = newSet.WindowTop;
            djvSet.WindowWidth = newSet.WindowWidth;
            djvSet.WindowViewType = newSet.WindowViewType;
            djvSet.WindowZoom = newSet.WindowZoom;
            djvSet.DisplayRulers = newSet.DisplayRulers;
            djvSet.ShowNavigationPanel = newSet.ShowNavigationPanel;
            djvSet.NavigationPanelWidth = newSet.NavigationPanelWidth;
            djvSet.RibbonHeight = newSet.RibbonHeight;

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

            // Remove all WindowLocation objects that have the same 
            // DisplayArrangementUID and ScreenID as a WindowLocation in the new set.
            if (newSet.Locations != null)
            {
                foreach (DejaviewSet.WindowLocation ol in djvSet.Locations)
                {
                    foreach (DejaviewSet.WindowLocation nl in newSet.Locations)
                    {
                        try
                        {
                            if (ol.SameScreenAs(nl)) locList.Remove(ol);
                        }
                        catch { }
                    }
                }
                locList.AddRange(newSet.Locations);
            }
            djvSet.Locations = locList.ToArray();
            SetDejaviewSet(djvSet);
        }

        /// <summary>
        /// Set the view of a given document.
        /// </summary>
        /// <param name="doc">Document to be displayed</param>
        /// <param name="djvSet">DejaviewSet representing the view parameters to be set</param>
        internal void SetDocumentView(Word.Document doc, DejaviewSet djvSet)
        {
            try
            {
                if ((Word.WdWindowState)djvSet.WindowState == Word.WdWindowState.wdWindowStateMinimize)
                    doc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateNormal;
                else
                    doc.ActiveWindow.WindowState = (Word.WdWindowState)djvSet.WindowState;

                if (DejaviewConfig.Instance.RememberWindowType)
                {
                    try
                    {
                        doc.ActiveWindow.View.Type = (Word.WdViewType)djvSet.WindowViewType;
                        Log("Window type restored (" + (Word.WdViewType)djvSet.WindowViewType + ").");
                    }
                    catch (Exception)
                    {
                        Log("Window type could not be restored (type=" + djvSet.WindowViewType + ").");
                        doc.ActiveWindow.View.Type = Word.WdViewType.wdNormalView;
                        Log("Window type set to Normal.");
                    }
                }

                if (DejaviewConfig.Instance.RememberZoom)
                {
                    try
                    {
                        doc.ActiveWindow.View.Zoom.Percentage = djvSet.WindowZoom;
                        Log("Window zoom restored (" + djvSet.WindowZoom + ").");
                    }
                    catch (Exception)
                    {
                        Log("Window zoom could not be restored (zoom=" + djvSet.WindowZoom + ").");
                        doc.ActiveWindow.View.Zoom.Percentage = 100;
                        Log("Window zoom set to 100%.");
                    }
                }

                if (DejaviewConfig.Instance.RememberRulers)
                {
                    try
                    {
                        doc.ActiveWindow.DisplayRulers = djvSet.DisplayRulers;
                        Log("Window rulers restored (" + djvSet.DisplayRulers + ").");
                    }
                    catch (Exception)
                    {
                        Log("Window rulers could not be restored (rulers=" + djvSet.DisplayRulers + ").");
                    }
                }

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
                        {
                            doc.ActiveWindow.ToggleRibbon();
                            Log("Window ribbon toggled (height: " + djvSet.RibbonHeight + ").");
                        }
                    }
                }

                if (DejaviewConfig.Instance.RememberNavigationPanel)
                {
                    if (djvSet.ShowNavigationPanel)
                    {
                        Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneNav].Visible = true;
                        Application.CommandBars["Navigation"].Width = djvSet.NavigationPanelWidth;
                        Log("Navigation panel restored (" + djvSet.ShowNavigationPanel + ", " + djvSet.NavigationPanelWidth + ").");
                    }
                    else
                    {
                        Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneNav].Visible = false;
                    }
                }

                // Attempt to restore window
                if (doc.ActiveWindow.WindowState == Word.WdWindowState.wdWindowStateNormal)
                    SetShowable(doc.Application, djvSet);

                SetDejaviewSet(djvSet);

                DisplayStatus("Document view restored.");

                SetButtonTip();

                views.Add(doc.DocID, true);
            }
            catch (DejaViewException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log(ex.Message);
            }
            catch (NullReferenceException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log("Error: " + ex.Message);
            }
            catch (IndexOutOfRangeException ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                Log("Index error: " + ex.Message);
            }
            catch (Exception ex)
            {
                Globals.Ribbons.DejaviewRibbon.btnRemove.Enabled = false;
                DisplayStatus("Could not restore document view. " + ex.Message);
                Debug.WriteLine(ex.StackTrace);
            }
        }

        /// <summary>
        /// Show a given Form ensuring that it is showable. If no Point is provided,
        /// the Form will be centered on the primary screen.
        /// </summary>
        /// <param name="app">Word application</param>
        /// <param name="ds">DejaviewSet for the ActiveDocument</param>
        private static void SetShowable(Word.Application app, DejaviewSet ds)
        {
            // Get Logger
            Logger logger = Globals.DejaviewAddIn.GetLogger();

            // Setup key variables
            Screen screen = Screen.FromPoint(new Point(app.Left, app.Top));
            DejaviewSet.WindowLocation windowLocation = null;

            // Make sure a valid DejaviewSet is stored in this document
            if (ds.Locations != null)
            {
                bool matched = false;
                DateTime latest = DateTime.MinValue;

                string dauid = GetDisplayArrangementUID();

                Debug.WriteLine(" Checking for a Display Arrangement...");
                Debug.WriteLine("   dauid: " + dauid);

                // First try to match a Display Arrangement
                foreach (DejaviewSet.WindowLocation wl in ds.Locations)
                {
                    Debug.WriteLine("    - " + wl.DisplayArrangementUID);
                    if (dauid == wl.DisplayArrangementUID)
                    {
                        Debug.WriteLine("      * matched *");
                        latest = wl.LastViewed;
                        windowLocation = wl;
                        Debug.WriteLine("      enumerating screens:");
                        Debug.WriteLine("        > " + wl.ScreenUID);
                        foreach (Screen scrn in Screen.AllScreens)
                        {
                            Debug.WriteLine("        - " + scrn);
                            if (wl.ScreenUID == GetScreenUID(scrn))
                            {
                                Debug.WriteLine("        * match *");
                                screen = scrn;
                                break;
                            }
                        }
                        matched = true;
                        Debug.WriteLine(" Display Arrangement MATCHED!");
                        logger.Add("Window location found in Display Arrangement.");
                        break;
                    }
                }

                // If a Display Arrangement is not matched, then try to set the document to
                // display on a screen that has the same ScreenUID
                if (!matched)
                {
                    Debug.WriteLine(" Display Arrangement not matched.");
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
                                logger.Add("Window location found for matching Screen.");
                                break;
                            }
                        }
                        if (matched) break;
                    }
                }
                if (!matched) logger.Add("Window location not found.");
            }

            // Get working area from the designated screen
            Rectangle workingArea = screen.WorkingArea;

            // If for some reason the height or width of the window is less 
            // than 100 pixels, then set it to a proportionate size.
            if (ds.WindowWidth < 100) ds.WindowWidth = (int)(workingArea.Width - Math.Round(workingArea.Width * 0.3));
            if (ds.WindowHeight < 100) ds.WindowHeight = (int)(workingArea.Height - Math.Round(workingArea.Height * 0.2));

            app.Width = ds.WindowWidth;
            logger.Add("Window width restored (" + ds.WindowWidth + ")");

            app.Height = ds.WindowHeight;
            logger.Add("Window height restored (" + ds.WindowHeight + ")");

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
                logger.Add("Window location restored (" + app.Left + ", " + app.Top + ").");
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

                    logger.Add("Window location reset (was not viewable).");
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
            int x = Application.Left;
            int y = Application.Top;
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
                Globals.DejaviewAddIn.GetLogger().Add(ex);
                return uid;
            }
        }

        /// <summary>
        /// Helper method for displaying status to MS Word's status bar.
        /// </summary>
        /// <param name="txt">Text to display in Word's status bar</param>
        internal void DisplayStatus(string txt)
        {
            Application.StatusBar = txt;
            Log(txt);
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
        /// Returns a formatted human readable string with Deja View tag details.
        /// </summary>
        /// <param name="djvSet">The DejaviewSet object to read</param>
        /// <returns>Human readable string with tag details</returns>
        internal string GetTags(DejaviewSet djvSet)
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
            if (djvSet.WindowState == 0)
            {
                str.Append("   (width: ");
                str.Append(djvSet.WindowWidth);
                str.Append(", height: ");
                str.Append(djvSet.WindowHeight);
                str.Append(")");
                str.AppendLine();
            }

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
            str.Append("   (zoom: ");
            str.Append(djvSet.WindowZoom);
            str.Append("%)");
            str.AppendLine();
            str.Append("View Rulers: \t");
            str.Append(djvSet.DisplayRulers);
            str.AppendLine();
            str.Append("View Navigation: \t");
            str.Append(djvSet.ShowNavigationPanel);
            str.Append("   (width: ");
            str.Append(djvSet.NavigationPanelWidth);
            str.Append(")");
            str.AppendLine();
            str.Append("Ribbon Height: \t");
            str.Append(djvSet.RibbonHeight);
            str.AppendLine();
            str.AppendLine();

            if (djvSet.Locations != null)
            {
                string sid = GetActiveScreenUID();
                str.AppendLine("Window Locations: ");
                foreach (DejaviewSet.WindowLocation wl in djvSet.Locations)
                {
                    str.Append((sid == wl.ScreenUID) ? " * " : "   ");
                    str.Append("Screen: ");
                    str.Append(GetScreenNameFromUID(wl.ScreenUID));
                    str.Append("  (top: ");
                    str.Append(wl.WindowTop);
                    str.Append(", left: ");
                    str.Append(wl.WindowLeft);
                    str.Append(")");
                    str.AppendLine();
                    str.Append("      Display Arrangement: ");
                    str.Append(wl.DisplayArrangementUID);
                    str.AppendLine();
                    str.Append("      Last viewed: ");
                    str.Append(wl.LastViewed);
                    str.AppendLine();
                    str.AppendLine();
                }
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
                Globals.DejaviewAddIn.Log("Error while checking for an update: " + ex);
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
            this.Startup += new EventHandler(DejaviewAddIn_Startup);
            this.Shutdown += new EventHandler(DejaviewAddIn_Shutdown);
            // Moved to InitializeCustom()
            //this.Application.DocumentOpen += DejaviewAddIn_DocumentOpen;
            //this.Application.DocumentBeforeSave += DejaviewAddIn_DocumentBeforeSave;
        }

        #endregion
    }
}
