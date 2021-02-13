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
        /// Flag for enabling Dejaview.
        /// </summary>
        public bool Enable { get; set; }
        /// <summary>
        /// Flag for including the Word application window's location
        /// in the DejaviewSet of parameters.
        /// </summary>
        public bool Location { get; set; }
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
        /// Example: https://lexem.cc/dejaview/latest/
        /// </summary>
        public string UpdateURL { get; set; }

        /*
        public bool ShowNavigationPanel { get; set; }
        public int NavigationPanelWidth { get; set; }
        public int WindowWidth { get; set; }
        public int WindowHeight { get; set; }
        public int WindowZoom { get; set; }
        public int WindowViewType { get; set; }
        public bool DraftView { get; set; }
        */

        /// <summary>
        /// Standard method for getting an active instance object of 
        /// this class.
        /// </summary>
        public static DejaviewConfig Instance
        {
            get { return _instance ?? new DejaviewConfig(); }
        }

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
            }
        }

        private void SetDefaults()
        {
            Enable = true;
            Location = true;
            Prompt = false;
            CheckForUpdates = true;
            UpdateURL = "http://lexem.cc/dejaview/latest/";

            /*
            ShowNavigationPanel = true;
            NavigationPanelWidth = 250;
            WindowWidth = 800;
            WindowZoom = 100;
            WindowViewType = (int)Microsoft.Office.Interop.Word.WdViewType.wdNormalView;
            DraftView = true;
            */

            Debug.WriteLine("DejaviewConfig::SetDefaults() -> done");
        }

        private void LoadFromConfig()
        {
            using (FileStream fs = File.OpenRead(configFile))
            {
                try
                {
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(fs);

                    var djv = xdoc.SelectSingleNode("//*[local-name()='Dejaview']");
                    if (djv == null) 
                        djv = xdoc.SelectSingleNode("//*[local-name()='Perview']");
                    if (djv == null) return;

                    var nodes = djv.ChildNodes;
                    foreach (XmlNode n in nodes)
                    {
                        if (n.Name == "Enable" && !string.IsNullOrEmpty(n.InnerText))
                            Enable = bool.Parse(n.InnerText);
                        else if (n.Name == "Location" && !string.IsNullOrEmpty(n.InnerText))
                            Location = bool.Parse(n.InnerText);
                        else if (n.Name == "Prompt" && !string.IsNullOrEmpty(n.InnerText))
                            Prompt = bool.Parse(n.InnerText);
                        else if (n.Name == "CheckForUpdates" && !string.IsNullOrEmpty(n.InnerText))
                            CheckForUpdates = bool.Parse(n.InnerText);
                        else if (n.Name == "UpdateURL" && !string.IsNullOrEmpty(n.InnerText))
                            UpdateURL = n.InnerText;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("DejaviewConfig::LoadFromConfig() => " + ex.StackTrace);
                }
            }
            Debug.WriteLine("DejaviewConfig::LoadFromConfig() -> success");
            Debug.WriteLine("  Enable    => " + Enable);
            Debug.WriteLine("  Location  => " + Location);
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
                        w.WriteElementString("Location", Location.ToString());
                        w.WriteElementString("Prompt", Prompt.ToString());
                        w.WriteElementString("CheckForUpdates", CheckForUpdates.ToString());
                        w.WriteElementString("UpdateURL", UpdateURL);

                        /*
                        w.WriteStartElement("Defaults");
                        w.WriteAttributeString("ShowNavigationPanel", ShowNavigationPanel.ToString());
                        w.WriteAttributeString("NavigationPanelWidth", NavigationPanelWidth.ToString());
                        w.WriteAttributeString("WindowWidth", WindowWidth.ToString());
                        w.WriteAttributeString("WindowHeight", WindowHeight.ToString());
                        w.WriteAttributeString("WindowTop", WindowTop.ToString());
                        w.WriteAttributeString("WindowLeft", WindowLeft.ToString());
                        w.WriteAttributeString("WindowZoom", WindowZoom.ToString());
                        w.WriteAttributeString("WindowViewType", WindowViewType.ToString());
                        w.WriteAttributeString("DraftView", DraftView.ToString());
                        w.WriteFullEndElement();
                        */

                        w.WriteEndElement();
                        w.WriteEndDocument();
                        w.Flush();
                    }
                    fs.Close();
                }
                Debug.WriteLine("DejaviewConfig::Save() -> success");
                Debug.WriteLine("  Enable    => " + Enable);
                Debug.WriteLine("  Location  => " + Location);
                Debug.WriteLine("  Prompt    => " + Prompt);
                Debug.WriteLine("  Check     => " + CheckForUpdates);
                Debug.WriteLine("  UpdateURL => " + UpdateURL);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DejaviewConfig::Save() => " + ex.StackTrace);
            }
        }
    }
}
