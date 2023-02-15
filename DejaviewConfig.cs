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
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Dejaview
{
    /// <summary>
    /// Custom configuration class used for persistent memory of Add-in settings.
    /// </summary>
    internal class DejaviewConfig
    {
        private static readonly string configFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Dejaview.config");

        private static DejaviewConfig _instance = null;

        /// <summary>
        /// Flag for enabling Deja View.
        /// </summary>
        public bool Enable { get; set; }
        /// <summary>
        /// Flag for prompting before saving.
        /// </summary>
        public bool Prompt { get; set; }
        /// <summary>
        /// Flag for checking for updates.
        /// </summary>
        public bool CheckForUpdates { get; set; }
        /// <summary>
        /// Fully formed URL for to check for updates.
        /// Example: https://dejaview.lexem.cc/latest/
        /// </summary>
        public string UpdateURL { get; set; }
        /// <summary>
        /// Remember the Word application window location
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberWindowLocation { get; set; }
        /// <summary>
        /// Remember the state and width of the Navigation Panel
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberNavigationPanel { get; set; }
        /// <summary>
        /// Remember the Word application window view type 
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberWindowType { get; set; }
        /// <summary>
        /// Remember the Word application window zoom 
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberZoom { get; set; }
        /// <summary>
        /// Remember the Word application rulers
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberRulers { get; set; }
        /// <summary>
        /// Remember the Word application ribbon 
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool RememberRibbon { get; set; }

        /// <summary>
        /// DejaviewSet object representing a user-defined default view 
        /// for documents. Optional.
        /// </summary>
        public DejaviewSet DefaultDejaviewSet { get; set; }

        /// <summary>
        /// Standard method for getting an active instance object of 
        /// this class.
        /// </summary>
        public static DejaviewConfig Instance
        {
            get { return _instance ?? new DejaviewConfig(); }
        }

        /// <summary>
        /// Instances of this class should be created using DejaviewConfig.Instance.
        /// </summary>
        /// <see cref="Instance"/>
        private DejaviewConfig()
        {
            _instance = this;

            try
            {
                if (!File.Exists(configFile))
                    SetDefaults();
                else
                    LoadFromConfig();
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewConfig::DejaviewConfig() => " + ex.StackTrace);
                Globals.DejaviewAddIn.DisplayStatus(ex.Message);
                Globals.DejaviewAddIn.Log(ex);
            }
        }

        /// <summary>
        /// Reset all in memory configuration settings to default. 
        /// Save() must be called to write these to the file.
        /// </summary>
        /// <see cref="Save"/>
        public void SetDefaults()
        {
            Enable = true;
            RememberWindowLocation = true;
            Prompt = false;
            CheckForUpdates = true;
            UpdateURL = "https://dejaview.lexem.cc/autoupdate";

            RememberWindowLocation = true;
            RememberNavigationPanel = true;
            RememberWindowType = true;
            RememberZoom = true;
            RememberRulers = true;
            RememberRibbon = true;

            Debug.WriteLine("DejaviewConfig::SetDefaults() -> done");
        }

        /// <summary>
        /// A helper method that is automatically called by the constructor.
        /// </summary>
        private void LoadFromConfig()
        {
            using (FileStream fs = File.OpenRead(configFile))
            {
                try
                {
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(fs);

                    var djv = xdoc.SelectSingleNode("//*[local-name()='Dejaview']");
                    if (djv == null) return;

                    var nodes = djv.ChildNodes;
                    foreach (XmlNode n in nodes)
                    {
                        if (n.Name == "Enable" && !string.IsNullOrEmpty(n.InnerText))
                            Enable = bool.Parse(n.InnerText);
                        else if (n.Name == "Prompt" && !string.IsNullOrEmpty(n.InnerText))
                            Prompt = bool.Parse(n.InnerText);
                        else if (n.Name == "CheckForUpdates" && !string.IsNullOrEmpty(n.InnerText))
                            CheckForUpdates = bool.Parse(n.InnerText);
                        else if (n.Name == "UpdateURL" && !string.IsNullOrEmpty(n.InnerText))
                            UpdateURL = n.InnerText;
                        else if (n.Name == "RememberWindowLocation" && !string.IsNullOrEmpty(n.InnerText))
                            RememberWindowLocation = bool.Parse(n.InnerText);
                        else if (n.Name == "RememberNavigationPanel" && !string.IsNullOrEmpty(n.InnerText))
                            RememberNavigationPanel = bool.Parse(n.InnerText);
                        else if (n.Name == "RememberWindowType" && !string.IsNullOrEmpty(n.InnerText))
                            RememberWindowType = bool.Parse(n.InnerText);
                        else if (n.Name == "RememberZoom" && !string.IsNullOrEmpty(n.InnerText))
                            RememberZoom = bool.Parse(n.InnerText);
                        else if (n.Name == "RememberRulers" && !string.IsNullOrEmpty(n.InnerText))
                            RememberRulers = bool.Parse(n.InnerText);
                        else if (n.Name == "RememberRibbon" && !string.IsNullOrEmpty(n.InnerText))
                            RememberRibbon = bool.Parse(n.InnerText);
                    }

                    djv = xdoc.SelectSingleNode("//*[local-name()='DefaultDejaviewSet']");

                    if (djv != null)
                    {
                        DefaultDejaviewSet = new DejaviewSet();
                        nodes = djv.ChildNodes;
                        foreach (XmlNode n in nodes)
                        {
                            if (n.Name == "NavigationPanelWidth" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.NavigationPanelWidth = int.Parse(n.InnerText);
                            else if (n.Name == "ShowNavigationPanel" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.ShowNavigationPanel = bool.Parse(n.InnerText);
                            else if (n.Name == "WindowWidth" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowWidth = int.Parse(n.InnerText);
                            else if (n.Name == "WindowHeight" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowHeight = int.Parse(n.InnerText);
                            else if (n.Name == "WindowLeft" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowLeft = int.Parse(n.InnerText);
                            else if (n.Name == "WindowTop" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowTop = int.Parse(n.InnerText);
                            else if (n.Name == "WindowState" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowState = int.Parse(n.InnerText);
                            else if (n.Name == "WindowViewType" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowViewType = int.Parse(n.InnerText);
                            else if (n.Name == "DraftView" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.DraftView = bool.Parse(n.InnerText);
                            else if (n.Name == "DisplayRulers" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.DisplayRulers = bool.Parse(n.InnerText);
                            else if (n.Name == "WindowZoom" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.WindowZoom = int.Parse(n.InnerText);
                            else if (n.Name == "RibbonHeight" && !string.IsNullOrEmpty(n.InnerText))
                                DefaultDejaviewSet.RibbonHeight = int.Parse(n.InnerText);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("DejaviewConfig::LoadFromConfig() => " + ex.StackTrace);
                }
            }
            Debug.WriteLine("DejaviewConfig::LoadFromConfig() -> success");
            Debug.WriteLine("  Enable    => " + Enable);
            Debug.WriteLine("  Prompt    => " + Prompt);
            Debug.WriteLine("  Check     => " + CheckForUpdates);
            Debug.WriteLine("  UpdateURL => " + UpdateURL);
        }

        /// <summary>
        /// Method to save current view settings to document.
        /// </summary>
        internal void Save()
        {
            try
            {
                using (FileStream fs = new FileStream(configFile, FileMode.Create))
                {
                    XmlWriterSettings settings = new XmlWriterSettings();
                    settings.ConformanceLevel = ConformanceLevel.Document;
                    settings.Indent = true;

                    using (XmlWriter w = XmlWriter.Create(fs, settings))
                    {
                        w.WriteStartDocument();

                        w.WriteStartElement("Dejaview");
                        w.WriteElementString("Enable", Enable.ToString());
                        w.WriteElementString("Prompt", Prompt.ToString());
                        w.WriteElementString("CheckForUpdates", CheckForUpdates.ToString());
                        w.WriteElementString("UpdateURL", UpdateURL);
                        w.WriteElementString("RememberWindowLocation", RememberWindowLocation.ToString());
                        w.WriteElementString("RememberNavigationPanel", RememberNavigationPanel.ToString());
                        w.WriteElementString("RememberWindowType", RememberWindowType.ToString());
                        w.WriteElementString("RememberZoom", RememberZoom.ToString());
                        w.WriteElementString("RememberRulers", RememberRulers.ToString());
                        w.WriteElementString("RememberRibbon", RememberRibbon.ToString());
                        w.WriteEndElement();

                        if (DefaultDejaviewSet != null)
                        {
                            w.WriteStartElement("DefaultDejaviewSet");
                            w.WriteAttributeString("NavigationPanelWidth", DefaultDejaviewSet.NavigationPanelWidth.ToString());
                            w.WriteAttributeString("ShowNavigationPanel", DefaultDejaviewSet.ShowNavigationPanel.ToString());
                            w.WriteAttributeString("WindowWidth", DefaultDejaviewSet.WindowWidth.ToString());
                            w.WriteAttributeString("WindowHeight", DefaultDejaviewSet.WindowHeight.ToString());
                            w.WriteAttributeString("WindowLeft", DefaultDejaviewSet.WindowLeft.ToString());
                            w.WriteAttributeString("WindowTop", DefaultDejaviewSet.WindowTop.ToString());
                            w.WriteAttributeString("WindowState", DefaultDejaviewSet.WindowState.ToString());
                            w.WriteAttributeString("WindowViewType", DefaultDejaviewSet.WindowViewType.ToString());
                            w.WriteAttributeString("DraftView", DefaultDejaviewSet.DraftView.ToString());
                            w.WriteAttributeString("DisplayRulers", DefaultDejaviewSet.DisplayRulers.ToString());
                            w.WriteAttributeString("WindowZoom", DefaultDejaviewSet.WindowZoom.ToString());
                            w.WriteAttributeString("RibbonHeight", DefaultDejaviewSet.RibbonHeight.ToString());
                            w.WriteFullEndElement();
                        }

                        w.WriteEndDocument();
                        w.Flush();
                    }
                    fs.Close();
                }
                Debug.WriteLine("DejaviewConfig::Save() -> success");
                Debug.WriteLine("  Enable    => " + Enable);
                Debug.WriteLine("  Prompt    => " + Prompt);
                Debug.WriteLine("  Check     => " + CheckForUpdates);
                Debug.WriteLine("  UpdateURL => " + UpdateURL);

                Globals.DejaviewAddIn.Log("Successfully saved configuration.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewConfig::Save() => " + ex.StackTrace);
                Globals.DejaviewAddIn.Log(ex);
            }
        }
    }
}
