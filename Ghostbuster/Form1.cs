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

//Application Icon from: http://www.iconspedia.com/icon/scream-473-.html
//By: Jojo Mendoza 
//License: Creative Commons Attribution-Noncommercial-No Derivative Works 3.0 License. 

#endregion

#region changelog

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
//----------   ---   -------------------------------------------------------------------------------
//TODO             - SetupDiLoadClassIcon()
//                 - SetupDiLoadDeviceIcon()
//----------   ---   -------------------------------------------------------------------------------

#endregion changelog

namespace Ghostbuster
{
    using System;
    using System.Collections;
    using System.Collections.Specialized;
    using System.Drawing;
    using System.IO;
    using HDEVINFO = System.IntPtr;
    using System.Windows.Forms;

    using Swiss;
    using GhostBuster;
    using System.Security.Principal;
    using System.Diagnostics;
    using System.ComponentModel;
    using System.Runtime.InteropServices;

    public partial class Form1 : Form
    {
        #region Fields

        /// <summary>
        /// Section Name in IniFile
        /// </summary>
        const String ClassKey = "Class.RemoveIfGosted";

        /// <summary>
        /// Section Name in IniFile
        /// </summary>
        const String DeviceKey = "Descr.RemoveIfGosted";

        /// <summary>
        /// IniFileName (Should Automatically use %AppData% when neccesary)!
        /// </summary>
        private String IniFileName = Path.ChangeExtension(Path.GetFileName(Application.ExecutablePath), ".ini");

        /// <summary>
        /// A Handle.
        /// </summary>
        private HDEVINFO aDevInfoSet;

        /// <summary>
        /// A Structure.
        /// </summary>
        private SetupDi.SP_DEVINFO_DATA aDeviceInfoData;

        /// <summary>
        /// The ToolTip used for displaying Context Menu Info.
        /// </summary>
        internal ToolTip InfoToolTip = new ToolTip();

        /// <summary>
        /// The ToolTip used for displaying UAC Info.
        /// </summary>
        internal ToolTip UACToolTip = new ToolTip();

        #endregion Fields

        #region Constructors

        /// <summary>
        /// The Constructor.
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            //TODO Use %APPDATA% here.... (Should be automatic in IniFile if nu path is passed)!
            using (IniFile ini = new IniFile(IniFileName))
            {
                //ini.WriteString(
                //    "RemoveIfGosted",
                //    String.Format("item_{0}", ini.ReadSection("RemoveIfGosted").Count),
                //    "USB-apparaat voor massaopslag");
                ini.UpdateFile();
            }

            DoubleBufferListView.SetExStyles(listView1);
        }

        #endregion Constructors

        #region Event Handlers

        /// <summary>
        /// Enumerate All Devices
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            WindowsPrincipal pricipal = new WindowsPrincipal(WindowsIdentity.GetCurrent());
            bool hasAdministrativeRight = pricipal.IsInRole(WindowsBuiltInRole.Administrator);

            if (!IsAdmin())
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
        /// Enumerate Devices and Delete Ghosts that are marked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveBtn_Click(object sender, EventArgs e)
        {
            if (IsAdmin())
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
        /// Enumerate Devices.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RefreshBtn_Click(object sender, EventArgs e)
        {
            Enumerate(false);
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
                String Class = listView1.SelectedItems[0].Group.ToString();
                String Device = listView1.SelectedItems[0].Text;
                Int32 ndx = listView1.SelectedItems[0].Index;

                using (IniFile ini = new IniFile(IniFileName))
                {
                    ini.WriteString(ClassKey,
                        String.Format("item_{0}", ini.ReadInteger(ClassKey, "Count", 0)),
                        Class);
                    ini.WriteInteger(ClassKey,
                        "Count",
                        ini.ReadInteger(ClassKey, "Count", 0) + 1);
                    ini.UpdateFile();
                }

                ReColorDevices(true);
            }
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
                String Class = listView1.SelectedItems[i].Group.ToString();
                String Device = listView1.SelectedItems[i].Text;
                Int32 ndx = listView1.SelectedItems[i].Index;

                using (IniFile ini = new IniFile(IniFileName))
                {
                    ini.WriteString(DeviceKey,
                        String.Format("item_{0}", ini.ReadInteger(DeviceKey, "Count", 0)),
                        Device);
                    ini.WriteInteger(DeviceKey,
                        "Count",
                        ini.ReadInteger(DeviceKey, "Count", 0) + 1);
                    ini.UpdateFile();
                }
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
            //Int32 fndx = -1;
            for (Int32 i = 0; i < listView1.SelectedItems.Count; i++)
            {
                String Class = listView1.SelectedItems[i].Group.ToString();
                String Device = listView1.SelectedItems[i].Text;
                Int32 ndx = listView1.SelectedItems[i].Index;

                using (IniFile ini = new IniFile(IniFileName))
                {
                    foreach (DictionaryEntry de in ini.ReadSection(DeviceKey))
                    {
                        if (de.Value.ToString() == Device)
                        {
                            ini.DeleteKey(DeviceKey, de.Key.ToString());
                        }
                    }
                    ini.UpdateFile();
                }
            }

            ReColorDevices(true);
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
                String Class = listView1.SelectedItems[i].Group.ToString();
                String Device = listView1.SelectedItems[i].Text;
                Int32 ndx = listView1.SelectedItems[i].Index;

                using (IniFile ini = new IniFile(IniFileName))
                {
                    foreach (DictionaryEntry de in ini.ReadSection(ClassKey))
                    {
                        if (de.Value.ToString() == Class)
                        {
                            ini.DeleteKey(ClassKey, de.Key.ToString());
                        }
                    }
                    ini.UpdateFile();
                }

                ReColorDevices(true);
            }
        }

        #endregion Event Handlers

        #region Methods

        /// <summary>
        /// Enumerate all devices and optionally uninstall ghosted ones.
        /// </summary>
        /// <param name="RemoveGhosts">true if ghosted devices should be uninstalled</param>
        private void Enumerate(Boolean RemoveGhosts)
        {
            using (new WaitCursor())
            {
                Enabled = false;

                toolStripProgressBar1.Value = 0;

                ReColorDevices(true);

                Int32 fndx = -1;

                try
                {
                    listView1.BeginUpdate();

                    listView1.Items.Clear();
                    listView1.Groups.Clear();

                    //Cache all HKLM Services Key Names and DisplayNames
                    SetupDi.EnumServices();

                    aDevInfoSet = SetupDi.SetupDiGetClassDevs(ref SetupDi.NullGuid, 0, IntPtr.Zero, (uint)SetupDi.DIGCF.DIGCF_ALLCLASSES);

                    if (aDevInfoSet != (IntPtr)SetupDi.INVALID_HANDLE_VALUE)
                    {
                        aDeviceInfoData.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(aDeviceInfoData);

                        Int32 i = 0;

                        using (IniFile ini = new IniFile(IniFileName))
                        {
                            while (SetupDi.SetupDiEnumDeviceInfo(aDevInfoSet, i, ref aDeviceInfoData))
                            {
                                SetupDi.DeviceInfo aDeviceInfo = new SetupDi.DeviceInfo();

                                //DeviceClass is used for grouping items...
                                SetupDi.GetClassDescriptionFromGuid(ref aDeviceInfo, aDeviceInfoData.ClassGuid);

                                //veg: Gives Exceptions, remove and return error name...
                                try
                                {
                                    //TODO: Retrieving DeviceName fails.
                                    SetupDi.GetDeviceDescription(ref aDeviceInfo, aDevInfoSet, aDeviceInfoData);
                                    SetupDi.GetDeviceName(ref aDeviceInfo, aDevInfoSet, aDeviceInfoData);

                                    //Use Insert instead of Add...
                                    ListViewItem lvi = listView1.Items.Add(aDeviceInfo.description);

                                    foreach (ListViewGroup lvg in listView1.Groups)
                                    {
                                        if (lvg.Name == aDeviceInfo.deviceclass)
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
                                            if (String.Compare(lvg.Name, aDeviceInfo.deviceclass, true) >= 0)
                                            {
                                                Int32 ndx = listView1.Groups.IndexOf(lvg);
                                                listView1.Groups.Insert(ndx, new ListViewGroup(aDeviceInfo.deviceclass, aDeviceInfo.deviceclass));
                                                listView1.Groups[ndx].Items.Add(lvi);

                                                break;
                                            }
                                        }
                                    }

                                    if (lvi.Group == null)
                                    {
                                        Int32 ndx = listView1.Groups.Add(new ListViewGroup(aDeviceInfo.deviceclass, aDeviceInfo.deviceclass));
                                        listView1.Groups[ndx].Items.Add(lvi);
                                    }
                                    SetupDi.GetDeviceStatus(ref aDeviceInfo, aDevInfoSet, ref aDeviceInfoData);

                                    if (aDeviceInfo.disabled)
                                    {
                                        lvi.SubItems.Add("Disabled");
                                    }
                                    else if (aDeviceInfo.service)
                                    {
                                        lvi.SubItems.Add("Service");
                                    }
                                    else if (aDeviceInfo.ghosted)
                                    {
                                        lvi.SubItems.Add("Ghosted");
                                    }
                                    else
                                    {
                                        lvi.SubItems.Add("Ok");
                                    }

                                    //Remove Devices by Description
                                    StringCollection descrtoremove = ini.ReadSectionValues(DeviceKey);

                                    if (descrtoremove.Contains(aDeviceInfo.description.Trim()))
                                    {
                                        if (aDeviceInfo.ghosted && RemoveGhosts)
                                        {
                                            if (SetupDi.SetupDiRemoveDevice(aDevInfoSet, ref aDeviceInfoData))
                                            {
                                                lvi.SubItems[1].Text = "REMOVED";

                                                fndx = lvi.Index;

                                                if (toolStripProgressBar1.Value < toolStripProgressBar1.Maximum)
                                                {
                                                    toolStripProgressBar1.Increment(1);
                                                }

                                                statusStrip1.Invalidate();
                                                Application.DoEvents();
                                            }
                                        }

                                        //ReColorDevices();

                                        //listView1.EnsureVisible(lvi.Index);
                                    }

                                    //Remove Devices by DeviceClass
                                    StringCollection classtoremove = ini.ReadSectionValues(ClassKey);

                                    if (classtoremove.Contains(aDeviceInfo.deviceclass.Trim()))
                                    {
                                        if (aDeviceInfo.ghosted && RemoveGhosts)
                                        {
                                            if (SetupDi.SetupDiRemoveDevice(aDevInfoSet, ref aDeviceInfoData))
                                            {
                                                lvi.SubItems[1].Text = "REMOVED";

                                                toolStripProgressBar1.Value = toolStripProgressBar1.Value + 1;

                                                statusStrip1.Invalidate();
                                                Application.DoEvents();
                                            }
                                        }

                                        //listView1.EnsureVisible(lvi.Index);
                                    }
                                }
                                finally
                                {
                                    //
                                }

                                i++;

                                Application.DoEvents();
                            }
                        }
                    }
                }
                finally
                {
                    Enabled = true;
                }

                listView1.Columns[0].Width = -1;
                listView1.Columns[1].Width = -1;

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

        public void ReColorDevices(Boolean updatemax)
        {
            Int32 cnt = 0;
            Int32 watched = 0;
            Int32 ghosted = 0;
            Int32 removed = 0;

            using (IniFile ini = new IniFile(IniFileName))
            {
                //Remove Devices by Description
                StringCollection descrtoremove = ini.ReadSectionValues(DeviceKey);
                StringCollection classtoremove = ini.ReadSectionValues(ClassKey);

                foreach (ListViewItem lvi in listView1.Items)
                {
                    String grp = lvi.Group.ToString().Trim();

                    if (classtoremove.Contains(grp) ||
                        descrtoremove.Contains(lvi.Text.Trim()))
                    {
                        if (lvi.SubItems[1].Text.Equals("Ghosted"))
                        {
                            lvi.BackColor = Color.LightSalmon;

                            ghosted++;

                            if (updatemax)
                            {
                                toolStripProgressBar1.Maximum = ghosted;
                            }
                        }
                        else if (lvi.SubItems[1].Text.Equals("REMOVED"))
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
            }

            toolStripStatusLabel1.Text = String.Format("{0} Device(s)", cnt - removed);
            toolStripStatusLabel2.Text = String.Format("{0} Filtered", watched + ghosted);
            toolStripStatusLabel3.Text = String.Format("{0} to be removed", ghosted);
        }

        [DllImport("user32")]
        public static extern UInt32 SendMessage
            (IntPtr hWnd, UInt32 msg, UInt32 wParam, UInt32 lParam);

        internal const int BCM_FIRST = 0x1600; //Normal button
        internal const int BCM_SETSHIELD = (BCM_FIRST + 0x000C); //Elevated button

        public static bool IsAdmin()
        {
            WindowsIdentity user = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(user);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static internal void AddShieldToButton(Button b)
        {
            b.FlatStyle = FlatStyle.System;
            SendMessage(b.Handle, BCM_SETSHIELD, 0, 0xFFFFFFFF);
        }

        #endregion Methods

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

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                using (IniFile ini = new IniFile(IniFileName))
                {
                    AddDeviceMnu.Enabled = true;
                    RemoveDeviceMnu.Enabled = false;

                    String Device = listView1.SelectedItems[0].Text;

                    foreach (DictionaryEntry de in ini.ReadSection(DeviceKey))
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

                    foreach (DictionaryEntry de in ini.ReadSection(ClassKey))
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
            }
        }
    }
}