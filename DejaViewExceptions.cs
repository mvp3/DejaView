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
    /// Foundation class for all Deja View exceptions. 
    /// Specific exceptions should use the appropriate subclass.
    /// </summary>
    /// <seealso cref="DejaViewNoTagsException"/>
    /// <seealso cref="DejaViewInvalidTagException"/>
    [Serializable]
    public class DejaViewException : Exception
    {
        /// <summary>
        /// Use to create a Deja View exception without a message.
        /// </summary>
        public DejaViewException() : base() { }
        /// <summary>
        /// Use to create a Deja View exception with a specified message.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        public DejaViewException(string message) : base(message) { }
        /// <summary>
        /// Use to create a Deja View exception with a specified message and
        /// an inner exception that gave rise to this exception.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        /// <param name="innerException">The exception that gave rise to this exception.</param>
        public DejaViewException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// Exception indicating that the no Deja View tags could be found in the document.
    /// </summary>
    [Serializable]
    public class DejaViewNoTagsException : DejaViewException
    {
        private static readonly string DefaultMessage = "No Deja View tags found in document.";

        /// <summary>
        /// Use to create a DejaViewNoTagsException with the default message.
        /// </summary>
        public DejaViewNoTagsException() : base(DefaultMessage) { }
        /// <summary>
        /// Use to create a DejaViewNoTagsException with a specified message.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        public DejaViewNoTagsException(string message) : base(message) { }
        /// <summary>
        /// Use to create a DejaViewNoTagsException with a specified message and
        /// an inner exception that gave rise to this exception.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        /// <param name="innerException">The exception that gave rise to this exception.</param>
        public DejaViewNoTagsException(string message, Exception innerException) : base(message, innerException) { }
    }

    /// <summary>
    /// Exception indicating that the Deja View tags were found in the document but were invalid.
    /// If an error occurs while trying to read the Deja View tags, this exception is thrown.
    /// It can be used to indicated a deprecated tag format of a potentially incompatible verseion.
    /// By all means, we will try to keep all Deja View tag versions compatible, but in the event that
    /// this is not possible, this exception is a safety net.
    /// </summary>
    [Serializable]
    public class DejaViewInvalidTagException : DejaViewException
    {
        private static readonly string DefaultMessage = "Invalid Deja View tags found in document (deprecated?).";

        /// <summary>
        /// Use to create a DejaViewInvalidTagException with the default message.
        /// </summary>
        public DejaViewInvalidTagException() : base(DefaultMessage) { }
        /// <summary>
        /// Use to create a DejaViewInvalidTagException with a specified message.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        public DejaViewInvalidTagException(string message) : base(message) { }
        /// <summary>
        /// Use to create a DejaViewInvalidTagException with a specified message and
        /// an inner exception that gave rise to this exception.
        /// </summary>
        /// <param name="message">Specified text message to display.</param>
        /// <param name="innerException">The exception that gave rise to this exception.</param>
        public DejaViewInvalidTagException(string message, Exception innerException) : base(message, innerException) { }
    }
}
