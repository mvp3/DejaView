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
using System.Collections.Generic;

namespace Dejaview
{
    /// <summary>
    /// Dataset representing Deja View parameters.
    /// </summary>
    public class DejaviewSet
    {
        /// <summary>
        /// List of DejaviewSet.Location elements.
        /// </summary>
        /// <seealso cref="WindowLocation"/>
        public WindowLocation[] Locations { get; set; }
        /// <summary>
        /// WindowState of Word application window.
        /// </summary>
        public int WindowState { get; set; }
        /// <summary>
        /// Left point of Word application window location.
        /// </summary>
        public int WindowLeft { get; set; }
        /// <summary>
        /// Top point of Word application window location.
        /// </summary>
        public int WindowTop { get; set; }
        /// <summary>
        /// Width of Word application window.
        /// </summary>
        public int WindowWidth { get; set; }
        /// <summary>
        /// Height of Word application window.
        /// </summary>
        public int WindowHeight { get; set; }
        /// <summary>
        /// Value for this document's "View Type".
        /// </summary>
        public int WindowViewType { get; set; }
        /// <summary>
        /// Value for this document's zoom level.
        /// </summary>
        public int WindowZoom { get; set; }
        /// <summary>
        /// Height of the application ribbon.
        /// </summary>
        public int RibbonHeight { get; set; }
        /// <summary>
        /// Width of navigation panel if displayed.
        /// </summary>
        public int NavigationPanelWidth { get; set; }
        /// <summary>
        /// Flag for showing the navigation panel.
        /// </summary>
        public bool ShowNavigationPanel { get; set; }
        /// <summary>
        /// Flag for "Draft View" view type.
        /// </summary>
        public bool DraftView { get; set; }
        /// <summary>
        /// Value flag for showing rulers.
        /// </summary>
        public bool DisplayRulers { get; set; }

        /// <summary>
        /// Method for comparing two DejaviewSet objects. Only the following
        /// attributes will be compared: top, left, height, width, state, 
        /// view type, zoom level, rulers, navigation panel, and ribbon.
        /// </summary>
        /// <param name="djvSet"></param>
        /// <returns></returns>
        public bool Equals(DejaviewSet djvSet)
        {
            if (WindowHeight != djvSet.WindowHeight ||
                WindowLeft != djvSet.WindowLeft ||
                WindowState != djvSet.WindowState ||
                WindowTop != djvSet.WindowTop ||
                WindowViewType != djvSet.WindowViewType ||
                WindowWidth != djvSet.WindowWidth ||
                WindowZoom != djvSet.WindowZoom ||
                DisplayRulers != djvSet.DisplayRulers ||
                NavigationPanelWidth != djvSet.NavigationPanelWidth ||
                RibbonHeight != djvSet.RibbonHeight) return false;

            return true;
        }

        /// <summary>
        /// Inner class representing the window location of the document. 
        /// This added dimension allows for retainig a document's viewing 
        /// location per computer display screen. 
        /// 
        /// The same document can retain its window location on computer C1 
        /// with a display resolution of R1 and on computer C2 with a display 
        /// resolution of R2. Or, computer C1 with R1 and R2 configurations.
        /// </summary>
        public class WindowLocation : IComparable<WindowLocation>
        {
            /// <summary>
            /// Unique Identifier of the Display Arrangement for the current 
            /// device. A Display Arrangement represents: (a) the number of screens, 
            /// (b) the size of each screen, and (c) the order of each screen in 
            /// relation to one another. This allows Deja View to index a document's
            /// view settings per Display Arrangement.
            /// </summary>
            public string DisplayArrangementUID { get; set; }
            /// <summary>
            /// Unique Identifier of the Screen to which this Location belongs.
            /// </summary>
            public string ScreenUID { get; set; }
            /// <summary>
            /// Left point of Word application window location.
            /// </summary>
            public int WindowLeft { get; set; }
            /// <summary>
            /// Top point of Word application window location.
            /// </summary>
            public int WindowTop { get; set; }
            /// <summary>
            /// Timestamp of the last time the window was viewed on this screen.
            /// </summary>
            public DateTime LastViewed { get; set; }

            /// <summary>
            /// Method used to see if two WindowLocation objects refer to the same screen.
            /// </summary>
            /// <param name="other">The WindowLocation object to compare.</param>
            /// <returns>True if the DisplayArrangementUID and ScreenUID match.</returns>
            public bool SameScreenAs(WindowLocation other)
            {
                if (string.IsNullOrEmpty(DisplayArrangementUID) ||
                    string.IsNullOrEmpty(other.DisplayArrangementUID) ||
                    !DisplayArrangementUID.Equals(other.DisplayArrangementUID))
                    return false;

                if (string.IsNullOrEmpty(ScreenUID) ||
                    string.IsNullOrEmpty(other.ScreenUID))
                    return false;

                return ScreenUID.Equals(other.ScreenUID);
            }

            /// <summary>
            /// Method used for comparing and sorting instances of this object.
            /// When sorting, the LastViewed property is compared.
            /// </summary>
            /// <param name="other">The DejaviewSet.WindowLocation object to compare</param>
            /// <returns>Compare value</returns>
            /// <see cref="Comparer{T}"/>
            public int CompareTo(WindowLocation other)
            {
                return LastViewed.CompareTo(other.LastViewed);
            }
        }
    }
}
