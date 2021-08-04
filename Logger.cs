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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dejaview
{
    /// <summary>
    /// This is a class that stores logged events in memory.
    /// These logged events may be retrieved at runtime or
    /// saved to a file. 
    /// </summary>
    public class Logger
    {
        private static Logger _instance = null;

        private List<LoggerEvent> events = new List<LoggerEvent>();

        /// <summary>
        /// Standard method for getting an active instance object of 
        /// this class.
        /// </summary>
        public static Logger Instance
        {
            get { return _instance ?? new Logger(); }
        }

        /// <summary>
        /// Instances of this class should be created using Logger.Instance.
        /// </summary>
        /// <see cref="Instance"/>
        private Logger()
        {
            _instance = this;
        }

        /// <summary>
        /// Logs an event. Takes a <code>LoggerEvent</code>.
        /// </summary>
        /// <param name="e">Event</param>
        public void Add(LoggerEvent e)
        {
            events.Add(e);
        }

        /// <summary>
        /// Logs an event. Automatically creates a <code>LoggerEvent</code> given
        /// a string description.
        /// </summary>
        /// <param name="e">Event information</param>
        public void Add(string description)
        {
            events.Add(new LoggerEvent(description));
        }

        /// <summary>
        /// Logs an event. Automatically creates a <code>LoggerEvent</code> given
        /// a <code>Exception</code>.
        /// </summary>
        /// <param name="e">Event information</param>
        public void Add(Exception ex)
        {
            events.Add(new LoggerEvent(ex.Message.ToString()));
        }

        /// <summary>
        /// Returns the number of events that are currently logged.
        /// </summary>
        /// <returns>Total number of events logged.</returns>
        public int GetEventCount()
        {
            return events.Count;
        }

        /// <summary>
        /// Returns the last logged event.
        /// </summary>
        /// <returns>Last logged event.</returns>
        public string GetLastEvent()
        {
            LoggerEvent e = events.LastOrDefault();
            if (e != null) return e.Description;
            else return null;
        }

        /// <summary>
        /// Returns an array of all logged event.
        /// </summary>
        /// <returns>All logged events</returns>
        public LoggerEvent[] GetEvents()
        {
            return events.ToArray();
        }

        public override string ToString()
        {
            StringBuilder str = new StringBuilder();
            LoggerEvent[] events = GetEvents();
            foreach (LoggerEvent e in events)
            {
                str.Append("[");
                str.Append(e.Timestamp);
                str.Append("] - ");
                str.AppendLine(e.Description);
            }
            return str.ToString();
        }
    }
    
    /// <summary>
    /// A class that represents an event that can be logged.
    /// </summary>
    public class LoggerEvent
    {
        /// <summary>
        /// String that describes this event.
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Timestamp that identifies when this event was raised.
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// Default constructor that takes the description and automatically
        /// sets the timestamp to the moment when the constructor is called.
        /// </summary>
        /// <param name="description">String that describes this event.</param>
        public LoggerEvent(string description)
        {
            this.Description = description;
            this.Timestamp = DateTime.Now;
        }
    }
}
