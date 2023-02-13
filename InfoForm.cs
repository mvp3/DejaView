using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dejaview
{
    /// <summary>
    /// This is a basic form designed to display messages.
    /// </summary>
    public partial class InfoForm : Form
    {
        /// <summary>
        /// Constructor that automatically displays the message provided.
        /// </summary>
        /// <param name="info">Information to display</param>
        /// <param name="caption">Caption (title) for message window</param>
        public InfoForm(string info, string caption = null)
        {
            InitializeComponent();

            this.label1.Text = info;
            if (caption != null) this.Text = caption;
            this.Show();
        }
    }
}
