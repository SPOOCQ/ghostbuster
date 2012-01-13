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

#region Header

//Application Icon from: http://www.iconspedia.com/icon/scream-473-.html
//By: Jojo Mendoza
//License: Creative Commons Attribution-Noncommercial-No Derivative Works 3.0 License.
//----------   ---   -------------------------------------------------------------------------------
//Purpose:           GhostBuster
//By:                G.W. van der Vegt (wvd_vegt@knoware.nl)
//Url:               http://ghostbuster.codeplex.com
//Depends:           IniFile
//License:           New BSD License
//----------   ---   -------------------------------------------------------------------------------
//dd-mm-yyyy - who - description
//----------   ---   -------------------------------------------------------------------------------
//27-11-2009 - veg - Created.
//02-12-2009 - veg - Removed Listview group/column sorter as devices and groups are now inserts
//                   alphabetical.
//                 - Changed reference from IniCollection to IniFile.
//                 - Added Installer project
//                 - Uploaded to CodePlex.
//08-01-2010   veg - Added WaitCursor in Enumerate().
//                 - Removed Path from IniFileName so it will default to '%AppData%\GhostBuster'
//                   with the latest IniFile component.
//                 - Added UAC Shield to RemoveBtn (http://www.codeproject.com/KB/vista-security/UAC_Shield_for_Elevation.aspx).
//                 - Added code to restart with Elevated security.
//                 - Added UAC Tooltip.
//                 - Renamed some Methods and Components.
//                 - Ensure visibility after refresh (not removal).
//                 - Correctly Enable and Disable Context MenuItems.
//                 - Added Registry Access Rights to SetupDi.EnumServices().
//                   This solves the Security Violations on WHS/W2K3 Server.
//                 - Added try/catch in SetupDi.EnumServices() to prevent
//                   screwed up service registry entries like SBCore to crash the program.
// 16-03-2011  veg - Made the listview multiselect.
//                 - Removed non functional checkboxes.
//                 - Separated Coloring Code.
//                 - Added Statusbar.
// 21-04-2011  veg - Added Partial WildCard Support (pattern isnt'stored yet and does not get removed).
//                 - Added Match Type Column.
//                 - Use Match Type Column for coloring too.
//                 - Add load/save to wildcards.
//----------   ---   -------------------------------------------------------------------------------
//TODO             - SetupDiLoadClassIcon()
//                 - SetupDiLoadDeviceIcon()
//                 - More Device Info:
//
//                 - [HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Class\{4D36E978-E325-11CE-BFC1-08002BE10318}\0004]
//                   "DevLoader"="*ntkern"
//                   "NTMPDriver"="ser2pl64.sys"
//                   "EnumPropPages32"="MsPorts.dll,SerialPortPropPageProvider"
//                   "InfPath"="oem22.inf"
//                   "InfSection"="ComPort"
//                   "InfSectionExt"=".NTAMD64"
//                   "ProviderName"="Prolific"
//                   "DriverDateData"=hex:00,80,58,f5,76,c1,ca,01
//                   "DriverDate"="3-12-2010"
//                   "DriverVersion"="3.3.11.152"
//                   "MatchingDeviceId"="usb\\vid_067b&pid_2303"
//                   "DriverDesc"="Prolific USB-to-Serial Comm Port"
//----------   ---   -------------------------------------------------------------------------------

#endregion Header

namespace Ghostbuster
{
    using System;
    using System.Collections;
    using System.Collections.Specialized;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Security.Principal;
    using System.Windows.Forms;

    using GhostBuster;

    using Swiss;

    using HDEVINFO = System.IntPtr;
    using System.IO;
    using System.Text;

    public partial class Mainform : Form
    {
        #region Fields

        internal const int BCM_FIRST = 0x1600; //Normal button
        internal const int BCM_SETSHIELD = (BCM_FIRST + 0x000C); //Elevated button

        /// <summary>
        /// The ToolTip used for displaying Context Menu Info.
        /// </summary>
        internal ToolTip InfoToolTip = new ToolTip();

        /// <summary>
        /// The ToolTip used for displaying UAC Info.
        /// </summary>
        internal ToolTip UACToolTip = new ToolTip();

        private Buster buster;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// The Constructor.
        /// </summary>
        public Mainform()
        {
            InitializeComponent();

            buster = new Buster();

            DoubleBufferListView.SetExStyles(listView1);
        }

        #endregion Constructors

        #region Enumerations

        public enum LVC
        {
            DeviceCol = 0,
            StatusCol,
            MatchTypeCol,
            DescriptionCol
        }

        #endregion Enumerations

        #region Methods

        [DllImport("user32")]
        public static extern UInt32 SendMessage(IntPtr hWnd, UInt32 msg, UInt32 wParam, UInt32 lParam);

        public void ReColorDevices(Boolean updatemax)
        {
            Int32 cnt = 0;
            Int32 watched = 0;
            Int32 ghosted = 0;
            Int32 removed = 0;

            foreach (ListViewItem lvi in listView1.Items)
            {
                String grp = lvi.Group.ToString().Trim();

                lvi.SubItems[(int)LVC.MatchTypeCol].Text = "";

                if (Buster.Classes.Contains(grp))
                {
                    lvi.SubItems[(int)LVC.MatchTypeCol].Text += "[Class]";
                }

                if (Buster.Devices.Contains(lvi.Text.Trim()))
                {
                    lvi.SubItems[(int)LVC.MatchTypeCol].Text += "[Device]";
                }

                foreach (Wildcard w in Buster.Wildcards)
                {
                    if (w.IsMatch(lvi.Text.Trim()))
                    {
                        lvi.SubItems[(int)LVC.MatchTypeCol].Text += "[" + w.Pattern + "]";
                    }
                }

                if (!String.IsNullOrEmpty(lvi.SubItems[2].Text))
                {
                    if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("Ghosted"))
                    {
                        lvi.BackColor = Color.LightSalmon;

                        ghosted++;

                        if (updatemax)
                        {
                            toolStripProgressBar1.Maximum = ghosted;
                        }
                    }
                    else if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("REMOVED"))
                    {
                        lvi.BackColor = Color.Orchid;
                        removed++;
                    }
                    else
                    {
                        lvi.BackColor = Color.PaleGreen;
                        watched++;
                    }
                }
                else
                {
                    lvi.BackColor = SystemColors.Window;
                }

                cnt++;
            }

            listView1.Update();

            toolStripStatusLabel1.Text = String.Format("{0} Device(s)", cnt - removed);
            toolStripStatusLabel2.Text = String.Format("{0} Filtered", watched + ghosted);
            toolStripStatusLabel3.Text = String.Format("{0} to be removed", ghosted);
        }

        internal static void AddShieldToButton(Button b)
        {
            b.FlatStyle = FlatStyle.System;
            SendMessage(b.Handle, BCM_SETSHIELD, 0, 0xFFFFFFFF);
        }

        /// <summary>
        /// Add a DeviceClass to the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddClassMnu_Click(object sender, EventArgs e)
        {
            for (Int32 i = 0; i < listView1.SelectedItems.Count; i++)
            {
                Buster.AddClassKey(listView1.SelectedItems[0].Group.ToString());
            }

            ReColorDevices(true);
        }

        /// <summary>
        /// Add a Device to the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddDeviceMnu_Click(object sender, EventArgs e)
        {
            for (Int32 i = 0; i < listView1.SelectedItems.Count; i++)
            {
                Buster.AddDeviceKey(listView1.SelectedItems[i].Text);
            }

            ReColorDevices(true);
        }

        private void AddWildCardMnu_Click(object sender, EventArgs e)
        {
            String Device = String.Empty;

            if (listView1.SelectedItems.Count != 0)
            {
                using (IniFile ini = new IniFile(Buster.IniFileName))
                {
                    Device = listView1.SelectedItems[0].Text;

                    using (InputDialog dlg = new InputDialog("Wildcard", "Enter WildCard", Device))
                    {
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            Buster.AddWildCard(dlg.Input);

                            ReColorDevices(false);
                        }
                    }
                }
            }
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                using (IniFile ini = new IniFile(Buster.IniFileName))
                {
                    AddDeviceMnu.Enabled = true;
                    RemoveDeviceMnu.Enabled = false;

                    String Device = listView1.SelectedItems[0].Text;

                    foreach (DictionaryEntry de in ini.ReadSection(Buster.DeviceKey))
                    {
                        if (de.Value.ToString() == Device)
                        {
                            AddDeviceMnu.Enabled = false;
                            RemoveDeviceMnu.Enabled = true;

                            break;
                        }
                    }

                    AddClassMnu.Enabled = true;
                    RemoveClassMnu.Enabled = false;

                    foreach (DictionaryEntry de in ini.ReadSection(Buster.ClassKey))
                    {
                        String Class = listView1.SelectedItems[0].Group.ToString();

                        if (de.Value.ToString() == Class)
                        {
                            AddClassMnu.Enabled = false;
                            RemoveClassMnu.Enabled = true;

                            break;
                        }
                    }
                }

                removeToolStripMenuItem.DropDownItems.Clear();

                foreach (Wildcard wildcard in Buster.Wildcards)
                {
                    removeToolStripMenuItem.DropDownItems.Add(wildcard.Pattern.Replace("&", "&&"), null, RemoveWildcardToolStripMenuItem_Click).Tag = wildcard.Pattern;
                }

                removeToolStripMenuItem.Enabled = (removeToolStripMenuItem.DropDownItems.Count != 0);
            }
        }

        /// <summary>
        /// Enumerate all devices and optionally uninstall ghosted ones.
        /// </summary>
        /// <param name="RemoveGhosts">true if ghosted devices should be uninstalled</param>
        private void Enumerate(Boolean RemoveGhosts)
        {
            using (new WaitCursor())
            {
                //! Add a Overlay here?
                Buster.Enumerate();

                Enabled = false;

                StringBuilder sb = new StringBuilder();

                toolStripProgressBar1.Value = 0;

                ReColorDevices(true);

                Int32 fndx = -1;

                try
                {
                    listView1.BeginUpdate();

                    listView1.Items.Clear();
                    listView1.Groups.Clear();

                    using (IniFile ini = new IniFile(Buster.IniFileName))
                    {
                        foreach (HwEntry he in Buster.HwEntries)
                        {

                            //Use Insert instead of Add...
                            ListViewItem lvi = listView1.Items.Add(he.Description);
                            for (int j = 1; j < listView1.Columns.Count; j++)
                            {
                                lvi.SubItems.Add("");
                            }

                            foreach (ListViewGroup lvg in listView1.Groups)
                            {
                                if (lvg.Name == he.DeviceClass)
                                {
                                    lvg.Items.Add(lvi);
                                    break;
                                }
                            }

                            if (lvi.Group == null)
                            {
                                //Use Insert instead of Add...
                                foreach (ListViewGroup lvg in listView1.Groups)
                                {
                                    if (String.Compare(lvg.Name, he.DeviceClass, true) >= 0)
                                    {
                                        Int32 ndx = listView1.Groups.IndexOf(lvg);
                                        listView1.Groups.Insert(ndx, new ListViewGroup(he.DeviceClass, he.DeviceClass));
                                        listView1.Groups[ndx].Items.Add(lvi);

                                        break;
                                    }
                                }
                            }

                            if (lvi.Group == null)
                            {
                                Int32 ndx = listView1.Groups.Add(new ListViewGroup(he.DeviceClass, he.DeviceClass));
                                listView1.Groups[ndx].Items.Add(lvi);
                            }

                            lvi.SubItems[(int)LVC.StatusCol].Text = he.DeviceStatus;
                            lvi.SubItems[(int)LVC.DescriptionCol].Text = he.FriendlyName;

                            if (RemoveGhosts && he.RemoveDevice_IF_Ghosted_AND_Marked())
                            {
                                lvi.SubItems[(int)LVC.StatusCol].Text = "REMOVED";

                                fndx = lvi.Index;

                                if (toolStripProgressBar1.Value < toolStripProgressBar1.Maximum)
                                {
                                    toolStripProgressBar1.Increment(1);
                                }

                                statusStrip1.Invalidate();
                                Application.DoEvents();
                            }

                            Application.DoEvents();
                        }
                    }

                }
                finally
                {
                    Enabled = true;
                }

                listView1.Columns[(int)LVC.DeviceCol].Width = -1;
                listView1.Columns[(int)LVC.StatusCol].Width = -1;
                listView1.Columns[(int)LVC.DescriptionCol].Width = -1;

                listView1.EndUpdate();

                if (fndx != -1)
                {
                    listView1.EnsureVisible(fndx);
                }

                ReColorDevices(false);

                if (RemoveGhosts && toolStripProgressBar1.Value != 0)
                {
                    toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                }
            }
        }

        /// <summary>
        /// Enumerate All Devices and Set the UAC shield on the button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            WindowsPrincipal pricipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            bool hasAdministrativeRight = pricipal.IsInRole(WindowsBuiltInRole.Administrator);

            if (!Buster.IsAdmin())
            {
                AddShieldToButton(RemoveBtn);

                this.UACToolTip.ToolTipTitle = "UAC";
                this.UACToolTip.SetToolTip(RemoveBtn,
                    "\r\nFor Vista and Windows 7 Users:\r\n\r\n" +
                    "Ghostbuster requires admin rights for device removal.\r\n" +
                    "If you click this button GhostBuster will restart and ask for these rights.");
            }

            Enumerate(false);

            this.InfoToolTip.ToolTipTitle = "Help on Usage";
            this.InfoToolTip.SetToolTip(listView1,
                "\r\nUse the Right Click Context Menu to:\r\n\r\n" +
            "1) Add devices or classes to the removal list (if ghosted)\r\n" +
            "2) Removed devices or classes of the removal list.");
        }

        /// <summary>
        /// Enumerate Devices.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RefreshBtn_Click(object sender, EventArgs e)
        {
            Enumerate(false);
        }

        /// <summary>
        /// Enumerate Devices and Delete Ghosts that are marked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveBtn_Click(object sender, EventArgs e)
        {
            if (Buster.IsAdmin())
            {
                Enumerate(true);
            }
            else
            {
                //Must run with limited privileges in order to see the UAC window
                RestartElevated(Application.ExecutablePath);
            }
        }

        /// <summary>
        /// Remove a DeviceClass from the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveClassMnu_Click(object sender, EventArgs e)
        {
            for (Int32 i = 0; i < listView1.SelectedItems.Count; i++)
            {
                Buster.RemoveClassKey(listView1.SelectedItems[i].Group.ToString());
            }

            ReColorDevices(true);
        }

        /// <summary>
        /// Remove a Device from the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveDeviceMnu_Click(object sender, EventArgs e)
        {
            for (Int32 i = 0; i < listView1.SelectedItems.Count; i++)
            {
                Buster.RemoveDeviceKey(listView1.SelectedItems[i].Text);
            }

            ReColorDevices(true);
        }

        private void RemoveWildcardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String pattern = (String)((ToolStripMenuItem)sender).Tag;

            Buster.RemoveWildCard(pattern);

            ReColorDevices(false);
        }

        private void RestartElevated(String fileName)
        {
            ProcessStartInfo processInfo = new ProcessStartInfo();

            processInfo.Verb = "runas";
            processInfo.FileName = fileName;

            try
            {
                Process.Start(processInfo);

                Application.Exit();
            }
            catch (Win32Exception)
            {
                //Do nothing. Probably the user canceled the UAC window
            }
        }

        #endregion Methods
    }
}