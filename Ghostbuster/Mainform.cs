#region Header

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
//Application Icon from: http://www.iconspedia.com/icon/scream-473-.html
//By: Jojo Mendoza
//License: Creative Commons Attribution-Noncommercial-No Derivative Works 3.0 License.
//----------   ---   -------------------------------------------------------------------------------
//Purpose:           GhostBuster
//By:                G.W. van der Vegt (wvd_vegt@knoware.nl)
//Url:               http://ghostbuster.codeplex.com
//Depends:           IniFile
//                   SystemRestore
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
// 16-05-2012      - Added slightly modified System Restore Code to create a checkpoint.
//                 - Switched to simpler WMI based solution (CoInitializeEx and CoInitializeSecurity) failed.
// 17-05-2012      - Moved WmiRestorePoint into this class.
//                 - Removed test code from program.cs
// 18-05-2012      - Improved counting.
//                 - Changed color of ghosted but unfiltered devices.
//                 - Changed HwEntries into an ObservableCollection.
// 19-05-2012      - Added Properties Form.
//                 - Added Properties MenuItem.
//                 - Uploaded to CodePlex.
//----------   ---   -------------------------------------------------------------------------------
// 29-12-2012  veg - Improved System Restore Detection (search for rstrui.exe).
//                 - Flagged null guid device as system so it cannot be removed anymore.
//                 - Added <No device class specified> to devices with an empty deviceclass.
//                 - Set copyright to 2012.
//                 - Version set to v1.0.2.0.
//                 - Uploaded to CodePlex.
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
//
//TODO             - Respond to DeviceChanges?
//                 - Hide all non filtered items.
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
    using System.IO;
    using System.Management;
    using System.Runtime.InteropServices;
    using System.Security.Principal;
    using System.Text;
    using System.Threading;
    using System.Windows.Forms;

    using GhostBuster;

    using Swiss;

    using HDEVINFO = System.IntPtr;
    using System.Collections.Generic;
    using System.Reflection;

    public partial class Mainform : Form
    {
        #region Fields

        internal const int BCM_FIRST = 0x1600; //Normal button
        internal const int BCM_SETSHIELD = (BCM_FIRST + 0x000C); //Elevated button
        internal const int S_OK = 0;

        /// <summary>
        /// Name of the Wmi ReturnValue Property.
        /// </summary>
        public const String ReturnValue = "ReturnValue";

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

        /// <summary>
        /// SystemRestore EventType
        /// </summary>
        public enum EventType
        {
            /// <summary>
            /// Start of operation.
            /// </summary>
            BeginSystemChange = 100,

            /// <summary>
            /// End of operation.
            /// </summary>
            EndSystemChange = 101,

            /// <summary>
            // Windows XP only - used to prevent the restore points intertwined.
            /// </summary>
            BeginNestedSystemChange = 102,

            /// <summary>
            // Windows XP only - used to prevent the restore points intertwined.
            /// </summary>
            EndNestedSystemChange = 103
        }

        public enum LVC
        {
            DeviceCol = 0,
            ManuCol,
            StatusCol,
            MatchTypeCol,
            DescriptionCol
        }

        /// <summary>
        /// Type of restorations.
        /// </summary> 
        public enum RestoreType
        {
            /// <summary>
            /// Installing a new application.
            /// </summary>
            ApplicationInstall = 0,

            /// <summary>
            /// An application has been uninstalled.
            /// </summary>
            ApplicationUninstall = 1,

            /// <summary>
            /// System Restore.
            /// </summary>
            Restore = 6,

            /// <summary>
            /// Checkpoint.
            /// </summary>
            Checkpoint = 7,

            /// <summary>
            /// Device driver has been installed.
            /// </summary>
            DeviceDriverInstall = 10,

            /// <summary>
            /// Program used for 1st time.
            /// </summary>
            FirstRun = 11,

            /// <summary>
            /// An application has had features added or removed.
            /// </summary>
            ModifySettings = 12,

            /// <summary>
            /// An application needs to delete the restore point it created.
            /// </summary>
            CancelledOperation = 13,

            /// <summary>
            /// Restoring a backup.
            /// </summary>
            BackupRecovery = 14
        }

        #endregion Enumerations

        #region Methods

        [DllImport("user32")]
        public static extern UInt32 SendMessage(IntPtr hWnd, UInt32 msg, UInt32 wParam, UInt32 lParam);

        /// <summary>
        /// http://msdn.microsoft.com/en-us/library/windows/desktop/aa378847(v=vs.85).aspx
        /// </summary>
        /// <param name="description"></param>
        /// <param name="rt">RestoreType</param>
        /// <param name="et">EventType</param>
        /// <returns></returns>
        public static ManagementBaseObject WmiRestorePoint(String description, RestoreType rt = RestoreType.Checkpoint, EventType et = EventType.BeginSystemChange)
        {
            ManagementScope scope = new ManagementScope(@"\\.\root\DEFAULT");
            ManagementPath path = new ManagementPath("SystemRestore");
            ObjectGetOptions options = new ObjectGetOptions();
            ManagementClass process = new ManagementClass(scope, path, options);

            // Obtain in-parameters for the method
            ManagementBaseObject inParams =
                process.GetMethodParameters("CreateRestorePoint");

            // Add the input parameters.
            inParams["Description"] = description;
            inParams["EventType"] = (int)et;
            inParams["RestorePointType"] = (int)rt;

            // Execute the method and obtain the return values.
            ManagementBaseObject outParams =
                       process.InvokeMethod("CreateRestorePoint", inParams, null);

            return outParams;
        }

        public void ReColorDevices(Boolean updatemax)
        {
            Int32 cnt = 0;
            Int32 watched = 0;
            Int32 ghosted = 0;
            Int32 removed = 0;
            Int32 remove = 0;

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

                if (!String.IsNullOrEmpty(lvi.SubItems[(int)LVC.StatusCol].Text))
                {
                    //Clount Devices to be Removed
                    if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("Ghosted") &&
                       !String.IsNullOrEmpty(lvi.SubItems[(int)LVC.MatchTypeCol].Text))
                    {
                        lvi.BackColor = Color.LightSalmon;
                        remove++;
                    }
                    //Count Ghosted Devices that are not removed.
                    else if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("Ghosted") &&
                    String.IsNullOrEmpty(lvi.SubItems[(int)LVC.MatchTypeCol].Text))
                    {
                        lvi.BackColor = Color.Bisque;

                        ghosted++;

                        if (updatemax)
                        {
                            toolStripProgressBar1.Maximum = ghosted;
                        }
                    }
                    //Count Removed Devices
                    else if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("REMOVED"))
                    {
                        lvi.BackColor = Color.Orchid;
                        removed++;
                    }
                    //Count Filtered Devices
                    else if (!String.IsNullOrEmpty(lvi.SubItems[(int)LVC.MatchTypeCol].Text))
                    {
                        lvi.BackColor = Color.PaleGreen;

                        watched++;
                    }
                    //Present Devices.
                    else if (lvi.SubItems[(int)LVC.StatusCol].Text.Equals("Ok") ||
                        lvi.SubItems[(int)LVC.StatusCol].Text.Equals("Service"))
                    {
                        lvi.BackColor = SystemColors.Window;
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
            toolStripStatusLabel2.Text = String.Format("{0} Filtered", watched /*+ ghosted + remove*/);
            toolStripStatusLabel3.Text = String.Format("{0} Ghosted", ghosted + remove);
            toolStripStatusLabel4.Text = String.Format("{0} to be removed", remove);
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

        private void chkSysRestore_CheckedChanged(object sender, EventArgs e)
        {
            using (IniFile ini = new IniFile(Buster.IniFileName))
            {
                ini.WriteBool("Setup", "CreateCheckPoint", chkSysRestore.Checked);
                if (ini.Dirty)
                {
                    ini.UpdateFile();
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
        private void Enumerate(Boolean RemoveGhosts, Boolean HideUnfiltered = false)
        {
            using (new WaitCursor())
            {
                toolStripStatusLabel1.Text = String.Empty;
                toolStripStatusLabel2.Text = String.Empty;
                toolStripStatusLabel3.Text = String.Empty;
                toolStripStatusLabel4.Text = String.Empty;

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

                            if (he.Properties.ContainsKey(SPDRP.MFG.ToString()))
                            {
                                lvi.SubItems[(int)LVC.ManuCol].Text = he.Properties[SPDRP.MFG.ToString()];
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
                                        listView1.Groups.Insert(ndx, new ListViewGroup(he.DeviceClass, String.IsNullOrEmpty(he.DeviceClass) ? "<No device class specified>" : he.DeviceClass));
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

                            lvi.Tag = he;

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

                //Remove all devices with just Ok/Service and No Match
                if (HideUnfiltered)
                {
                    for (Int32 i = listView1.Items.Count - 1; i >= 0; i--)
                    {
                        if ((listView1.Items[i].SubItems[(int)LVC.StatusCol].Text.Equals("Ok") ||
                            listView1.Items[i].SubItems[(int)LVC.StatusCol].Text.Equals("Service")) &&
                            String.IsNullOrEmpty(listView1.Items[i].SubItems[(int)LVC.MatchTypeCol].Text))
                        {
                            listView1.Items.RemoveAt(i);
                        }
                    }
                }

                ReColorDevices(false);

                if (RemoveGhosts && toolStripProgressBar1.Value != 0)
                {
                    toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                }

                //foreach (ListViewGroup lvg in listView1.Groups)
                //{
                //    if (String.IsNullOrEmpty(lvg.Header))
                //    {
                //        lvg.Header = "<No Class Specified>";
                //    }
                //}
            }
        }

        /// <summary>
        /// Verifies that the OS can do system restores.
        /// </summary>
        /// <returns>True if OS is either ME,XP,Vista,7</returns>
        public static bool SysRestoreAvailable()
        {
            int majorVersion = Environment.OSVersion.Version.Major;
            int minorVersion = Environment.OSVersion.Version.Minor;

            StringBuilder sbPath = new StringBuilder(260);

            if (File.Exists(Path.Combine(Environment.SystemDirectory, "rstrui.exe")))
            {
                Debug.Print("rstrui.exe detected: '{0}'", Path.Combine(Environment.SystemDirectory, "rstrui.exe"));
                return true;
            }

            if (SearchPath(null, "rstrui.exe", null, 260, sbPath, null) != 0)
            {
                Debug.Print("rstrui.exe detected: '{0}'", sbPath.ToString());
                return true;
            }

            // See if DLL exists
            //if (SearchPath(null, "srclient.dll", null, 260, sbPath, null) != 0)
            //    return true;

            //// Windows ME
            //if (majorVersion == 4 && minorVersion == 90)
            //    return true;

            //// Windows XP
            //if (majorVersion == 5 && minorVersion == 1)
            //    return true;

            //// Windows Vista
            //if (majorVersion == 6 && minorVersion == 0)
            //    return true;

            //// Windows 7
            //if (majorVersion == 6 && minorVersion == 1)
            //    return true;

            // All others : Win 95, 98, 2000, Server
            return false;
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

            chkSysRestore.Enabled = SysRestoreAvailable();

            using (IniFile ini = new IniFile(Buster.IniFileName))
            {
                chkSysRestore.Checked = ini.ReadBool("Setup", "CreateCheckPoint", chkSysRestore.Checked);
            }

            if (chkSysRestore.Enabled)
            {
                Console.WriteLine("System Restore Available");
            }
            else
            {
                Console.WriteLine("System Restore Unavailable");
            }

            if (!Buster.IsAdmin())
            {
                AddShieldToButton(RemoveBtn);

                this.UACToolTip.ToolTipTitle = "UAC";
                this.UACToolTip.SetToolTip(RemoveBtn,
                    "\r\nFor Vista and Windows 7 Users:\r\n\r\n" +
                    "Ghostbuster requires admin rights for device removal.\r\n" +
                    "If you click this button GhostBuster will restart and ask for these rights.");
            }

            Buster.HwEntries.CollectionChanged += new NotifyCollectionChangedEventHandler(HwEntries_CollectionChanged);
            Enumerate(false);

            this.InfoToolTip.ToolTipTitle = "Help on Usage";
            this.InfoToolTip.SetToolTip(listView1,
                "\r\nUse the Right Click Context Menu to:\r\n\r\n" +
            "1) Add devices or classes to the removal list (if ghosted)\r\n" +
            "2) Removed devices or classes of the removal list.");
        }

        void HwEntries_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            toolStripStatusLabel1.Text = String.Format("{0} Device(s)", Buster.HwEntries.Count);
            statusStrip1.Update();
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
                if (chkSysRestore.Checked)
                {
                    try
                    {
                        label1.Text = "Creating a System Restore Point...";
                        label1.Update();

                        ManagementBaseObject result = WmiRestorePoint("GhostBuster Restore Point", RestoreType.ApplicationInstall, EventType.BeginSystemChange);

                        if (Int32.Parse(result[ReturnValue].ToString()) == S_OK)
                        {
                            label1.Text = "System Restore Point created successfully.";
                        }
                        else
                        {
                            label1.Text = "Creation of System Restore Point failed!";
                        }
                        label1.Update();

                        Thread.Sleep(1000);
                    }
                    catch (ManagementException err)
                    {
                        MessageBox.Show("An error occurred while trying to execute the WMI method: " + err.Message);
                    }
                }

                label1.Text = "Removing devices...";
                label1.Update();

                Enumerate(true);

                label1.Text = "Devices removed.";
                label1.Update();

                Thread.Sleep(1000);

                label1.Text = String.Empty;
                label1.Update();
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

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern uint SearchPath(string lpPath,
            string lpFileName,
            string lpExtension,
            int nBufferLength,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder lpBuffer,
            string lpFilePart);

        private void hideUnfilteredDevicesToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            Enumerate(false, ((ToolStripMenuItem)(sender)).CheckState == CheckState.Checked);
        }

        private void propertiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems != null &&
                listView1.SelectedItems.Count == 1 &&
               listView1.SelectedItems[0].Tag != null)
            {
                Debug.WriteLine("[Properties]");
                HwEntry he = (HwEntry)listView1.SelectedItems[0].Tag;

                PropertyForm pf = new PropertyForm();

                pf.textBox1.Clear();

                foreach (KeyValuePair<String, String> kvp in he.Properties)
                {
                    String descr = EnumExtensions.GetDescription<SPDRP>((SPDRP)Enum.Parse(typeof(SPDRP), kvp.Key));
                    if (!String.IsNullOrEmpty(descr))
                    {
                        pf.textBox1.Text += String.Format("{0,32} =  {1}\r\n", descr, kvp.Value);
                    }
                }

                pf.ShowDialog();
            }
        }

        #endregion Methods
    }
}