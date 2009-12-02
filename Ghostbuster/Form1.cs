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
        private String IniFileName = Path.ChangeExtension(Application.ExecutablePath, ".ini");

        /// <summary>
        /// A Handle.
        /// </summary>
        private HDEVINFO aDevInfoSet;

        /// <summary>
        /// A Structure.
        /// </summary>
        private SetupDi.SP_DEVINFO_DATA aDeviceInfoData;

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
            Enumerate(false);
        }

        /// <summary>
        /// Enumerate Devices and Delete Ghosts that are marked.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Enumerate(true);
        }

        /// <summary>
        /// Enumerate Devices.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Enumerate(false);
        }
                
        /// <summary>
        /// Disable Manual Checking of CheckBoxes.
        /// </summary>
        /// <param name="sender">-</param>
        /// <param name="e">-</param>
        private void listView1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            e.NewValue = e.CurrentValue;
        }

        /// <summary>
        /// Add a DeviceClass to the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                String Class = listView1.SelectedItems[0].Group.ToString();
                String Device = listView1.SelectedItems[0].Text;

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
                Enumerate(false);
            }
        }

        /// <summary>
        /// Add a Device to the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                String Class = listView1.SelectedItems[0].Group.ToString();
                String Device = listView1.SelectedItems[0].Text;

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
                Enumerate(false);
            }
        }

        /// <summary>
        /// Remove a DeviceClass from the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                String Class = listView1.SelectedItems[0].Group.ToString();
                String Device = listView1.SelectedItems[0].Text;

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
                Enumerate(false);
            }
        }

        /// <summary>
        /// Remove a DeviceClass from the Ghost Removal List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count != 0)
            {
                String Class = listView1.SelectedItems[0].Group.ToString();
                String Device = listView1.SelectedItems[0].Text;

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
                Enumerate(false);
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

                            //lvi.SubItems.Add(String.Format("0x{0:x8}", aDeviceInfo.status));
                            //lvi.SubItems.Add(String.Format("0x{0:x8}", (Int32)aDeviceInfo.problem));

                            //TODO Weak Code. Pass deviceinfo to above calls as ref.

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
                            foreach (String desc in descrtoremove)
                            {
                                if (aDeviceInfo.description.Equals(desc))
                                {
                                    if (aDeviceInfo.ghosted && RemoveGhosts)
                                    {
                                        if (SetupDi.SetupDiRemoveDevice(aDevInfoSet, ref aDeviceInfoData))
                                        {
                                            lvi.SubItems[1].Text = "REMOVED";
                                        }
                                    }
                                    lvi.Checked = aDeviceInfo.ghosted;

                                    if (aDeviceInfo.ghosted)
                                    {
                                        lvi.ForeColor = SystemColors.GrayText;
                                    }

                                    lvi.BackColor = SystemColors.Info;
                                }
                            }

                            //Remove Devices by DeviceClass
                            StringCollection classtoremove = ini.ReadSectionValues(ClassKey);
                            foreach (String name in classtoremove)
                            {
                                if (aDeviceInfo.deviceclass.Equals(name))
                                {
                                    if (aDeviceInfo.ghosted && RemoveGhosts)
                                    {
                                        if (SetupDi.SetupDiRemoveDevice(aDevInfoSet, ref aDeviceInfoData))
                                        {
                                            lvi.SubItems[1].Text = "REMOVED";
                                        }
                                    }
                                    lvi.Checked = aDeviceInfo.ghosted;

                                    if (aDeviceInfo.ghosted)
                                    {
                                        lvi.ForeColor = SystemColors.GrayText;
                                    }
                                    lvi.BackColor = SystemColors.Info;
                                }
                            }
                        }
                        finally
                        {
                            //
                        }

                        i++;
                    }
                }
            }

            listView1.Columns[0].Width = -1;
            listView1.Columns[1].Width = -1;

            listView1.EndUpdate();

            //ListViewGroupSorter fails, removes certain items.
            //((ListViewGroupSorter)listView1).SortGroups(true);
        }

        #endregion Methods

        #region Other

        //

        #endregion Other
    }
}