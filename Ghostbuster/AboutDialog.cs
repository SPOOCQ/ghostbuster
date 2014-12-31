#region License

/*
Copyright (c) 2009, G.W. van der Vegt
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided
that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this list of conditions and the
  following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice, this list of conditions and
  the following disclaimer in the documentation and/or other materials provided with the distribution.

* Neither the name of G.W. van der Vegt nor the names of its contributors may be
  used to endorse or promote products derived from this software without specific prior written
  permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF
THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
#endregion License

namespace Ghostbuster
{
    using System;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Reflection;
    using System.Windows.Forms;

    /// <summary>
    /// About Dialog.
    /// </summary>
    public partial class AboutDialog : Form
    {
        const String picasadownloader_at_codeplex = "http://ghostbuster.codeplex.com";
        const String picasadownloader_mailto = "mailto:wim@vander-vegt.nl?SUBJECT=Suggestion for Ghostbuster";
        const String paypal_picasadownloader = "https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=ZZLQ2FAAN757U";


        /// <summary>
        /// Constructor.
        /// </summary>
        public AboutDialog()
        {
            InitializeComponent();

            textBox1.Text = String.Format(textBox1.Text, Assembly.GetExecutingAssembly().GetName().Version);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(paypal_picasadownloader);
            }
            catch (Win32Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(picasadownloader_at_codeplex);
            }
            catch (Win32Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(picasadownloader_mailto);
            }
            catch (Win32Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
    }
}
