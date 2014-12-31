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
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Deployment.Application;
    using System.Windows.Forms;
    using Swiss;
    using System.Diagnostics;

    /// <summary>
    /// See http://msdn.microsoft.com/en-us/library/ms404263.aspx
    /// </summary>
    public class ClickOnceUpdater
    {
        internal static void InstallUpdateSyncWithInfo(String Url)
        {
            String Caption = String.Format("{0} - {1}", IniFile.AppTitle, "Updater");

            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    using (WaitCursor wc = new WaitCursor())
                    {
                        info = ad.CheckForDetailedUpdate();
                    }
                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message, Caption);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message, Caption);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message, Caption);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;

                    if (!info.IsUpdateRequired)
                    {
                        DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", Caption, MessageBoxButtons.OKCancel);
                        if (!(DialogResult.OK == dr))
                        {
                            doUpdate = false;
                        }
                    }
                    else
                    {
                        // Display a message that the app MUST reboot. Display the minimum required version.
                        MessageBox.Show("This application has detected a mandatory update from your current " +
                            "version to version " + info.MinimumRequiredVersion.ToString() +
                            ". The application will now install the update and restart.",
                            Caption, MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();
                            MessageBox.Show("The application has been upgraded, and will now restart.", Caption);
                            Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                            MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde, Caption);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("There is currently no update available.", Caption);
                }
            }
            else
            {
                if (MessageBox.Show(
                     String.Format("This automated check is only available for Click-Once Installers\r\n\r\nVisit the {0} website?", IniFile.AppTitle), Caption, MessageBoxButtons.YesNo,
                             MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    VisitWebsite(Url);
                }
            }
        }

        internal static void VisitWebsite(String Url)
        {
            Process.Start(Url);
        }
    }
}
