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
    using System.Text.RegularExpressions;
    using System.Collections.Generic;
    using System.Text;

    public partial class Form1 : Form
    {
        #region Fields

        /// <summary>
        /// Section Name in IniFile
        /// </summary>
        const String ClassKey = "Class.RemoveIfGosted";
        private const string WildcardKey = "Wildcards";

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

        internal List<Wildcard> wildcards = new List<Wildcard>();

        #endregion Fields

        public enum LVC
        {
            DeviceCol = 0,
            StatusCol,
            MatchTypeCol,
            DescriptionCol
        }

        public class GUID_DEVCLASS
        {
            //http://stackoverflow.com/questions/304986/how-do-i-get-the-friendly-name-of-a-com-port-in-windows
            public static readonly Guid _1394 = new Guid("{0x6bdd1fc1, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid _1394DEBUG = new Guid("{0x66f250d6, 0x7801, 0x4a64, {0xb1, 0x39, 0xee, 0xa8, 0x0a, 0x45, 0x0b, 0x24}}");
            public static readonly Guid _61883 = new Guid("{0x7ebefbc0, 0x3200, 0x11d2, {0xb4, 0xc2, 0x00, 0xa0, 0xc9, 0x69, 0x7d, 0x07}}");
            public static readonly Guid ADAPTER = new Guid("{0x4d36e964, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid APMSUPPORT = new Guid("{0xd45b1c18, 0xc8fa, 0x11d1, {0x9f, 0x77, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");
            public static readonly Guid AVC = new Guid("{0xc06ff265, 0xae09, 0x48f0, {0x81, 0x2c, 0x16, 0x75, 0x3d, 0x7c, 0xba, 0x83}}");
            public static readonly Guid BATTERY = new Guid("{0x72631e54, 0x78a4, 0x11d0, {0xbc, 0xf7, 0x00, 0xaa, 0x00, 0xb7, 0xb3, 0x2a}}");
            public static readonly Guid BIOMETRIC = new Guid("{0x53d29ef7, 0x377c, 0x4d14, {0x86, 0x4b, 0xeb, 0x3a, 0x85, 0x76, 0x93, 0x59}}");
            public static readonly Guid BLUETOOTH = new Guid("{0xe0cbf06c, 0xcd8b, 0x4647, {0xbb, 0x8a, 0x26, 0x3b, 0x43, 0xf0, 0xf9, 0x74}}");
            public static readonly Guid CDROM = new Guid("{0x4d36e965, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid COMPUTER = new Guid("{0x4d36e966, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid DECODER = new Guid("{0x6bdd1fc2, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid DISKDRIVE = new Guid("{0x4d36e967, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid DISPLAY = new Guid("{0x4d36e968, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid DOT4 = new Guid("{0x48721b56, 0x6795, 0x11d2, {0xb1, 0xa8, 0x00, 0x80, 0xc7, 0x2e, 0x74, 0xa2}}");
            public static readonly Guid DOT4PRINT = new Guid("{0x49ce6ac8, 0x6f86, 0x11d2, {0xb1, 0xe5, 0x00, 0x80, 0xc7, 0x2e, 0x74, 0xa2}}");
            public static readonly Guid ENUM1394 = new Guid("{0xc459df55, 0xdb08, 0x11d1, {0xb0, 0x09, 0x00, 0xa0, 0xc9, 0x08, 0x1f, 0xf6}}");
            public static readonly Guid FDC = new Guid("{0x4d36e969, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid FLOPPYDISK = new Guid("{0x4d36e980, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid GPS = new Guid("{0x6bdd1fc3, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid HDC = new Guid("{0x4d36e96a, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid HIDCLASS = new Guid("{0x745a17a0, 0x74d3, 0x11d0, {0xb6, 0xfe, 0x00, 0xa0, 0xc9, 0x0f, 0x57, 0xda}}");
            public static readonly Guid IMAGE = new Guid("{0x6bdd1fc6, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid INFINIBAND = new Guid("{0x30ef7132, 0xd858, 0x4a0c, {0xac, 0x24, 0xb9, 0x02, 0x8a, 0x5c, 0xca, 0x3f}}");
            public static readonly Guid INFRARED = new Guid("{0x6bdd1fc5, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid KEYBOARD = new Guid("{0x4d36e96b, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid LEGACYDRIVER = new Guid("{0x8ecc055d, 0x047f, 0x11d1, {0xa5, 0x37, 0x00, 0x00, 0xf8, 0x75, 0x3e, 0xd1}}");
            public static readonly Guid MEDIA = new Guid("{0x4d36e96c, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MEDIUM_CHANGER = new Guid("{0xce5939ae, 0xebde, 0x11d0, {0xb1, 0x81, 0x00, 0x00, 0xf8, 0x75, 0x3e, 0xc4}}");
            public static readonly Guid MODEM = new Guid("{0x4d36e96d, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MONITOR = new Guid("{0x4d36e96e, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MOUSE = new Guid("{0x4d36e96f, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MTD = new Guid("{0x4d36e970, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MULTIFUNCTION = new Guid("{0x4d36e971, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid MULTIPORTSERIAL = new Guid("{0x50906cb8, 0xba12, 0x11d1, {0xbf, 0x5d, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");
            public static readonly Guid NET = new Guid("{0x4d36e972, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid NETCLIENT = new Guid("{0x4d36e973, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid NETSERVICE = new Guid("{0x4d36e974, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid NETTRANS = new Guid("{0x4d36e975, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid NODRIVER = new Guid("{0x4d36e976, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid PCMCIA = new Guid("{0x4d36e977, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid PNPPRINTERS = new Guid("{0x4658ee7e, 0xf050, 0x11d1, {0xb6, 0xbd, 0x00, 0xc0, 0x4f, 0xa3, 0x72, 0xa7}}");
            public static readonly Guid PORTS = new Guid("{0x4d36e978, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid PRINTER = new Guid("{0x4d36e979, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid PRINTERUPGRADE = new Guid("{0x4d36e97a, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid PROCESSOR = new Guid("{0x50127dc3, 0x0f36, 0x415e, {0xa6, 0xcc, 0x4c, 0xb3, 0xbe, 0x91, 0x0B, 0x65}}");
            public static readonly Guid SBP2 = new Guid("{0xd48179be, 0xec20, 0x11d1, {0xb6, 0xb8, 0x00, 0xc0, 0x4f, 0xa3, 0x72, 0xa7}}");
            public static readonly Guid SCSIADAPTER = new Guid("{0x4d36e97b, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid SECURITYACCELERATOR = new Guid("{0x268c95a1, 0xedfe, 0x11d3, {0x95, 0xc3, 0x00, 0x10, 0xdc, 0x40, 0x50, 0xa5}}");
            public static readonly Guid SMARTCARDREADER = new Guid("{0x50dd5230, 0xba8a, 0x11d1, {0xbf, 0x5d, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");
            public static readonly Guid SOUND = new Guid("{0x4d36e97c, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid SYSTEM = new Guid("{0x4d36e97d, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid TAPEDRIVE = new Guid("{0x6d807884, 0x7d21, 0x11cf, {0x80, 0x1c, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid UNKNOWN = new Guid("{0x4d36e97e, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");
            public static readonly Guid USB = new Guid("{0x36fc9e60, 0xc465, 0x11cf, {0x80, 0x56, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00}}");
            public static readonly Guid VOLUME = new Guid("{0x71a27cdd, 0x812a, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");
            public static readonly Guid VOLUMESNAPSHOT = new Guid("{0x533c5b84, 0xec70, 0x11d2, {0x95, 0x05, 0x00, 0xc0, 0x4f, 0x79, 0xde, 0xaf}}");
            public static readonly Guid WCEUSBS = new Guid("{0x25dbce51, 0x6c8f, 0x4a72, {0x8a, 0x6d, 0xb5, 0x4c, 0x2b, 0x4f, 0xc8, 0x35}}");
            public static readonly Guid FSFILTER_ACTIVITYMONITOR = new Guid("{0xb86dff51, 0xa31e, 0x4bac, {0xb3, 0xcf, 0xe8, 0xcf, 0xe7, 0x5c, 0x9f, 0xc2}}");
            public static readonly Guid FSFILTER_UNDELETE = new Guid("{0xfe8f1572, 0xc67a, 0x48c0, {0xbb, 0xac, 0x0b, 0x5c, 0x6d, 0x66, 0xca, 0xfb}}");
            public static readonly Guid FSFILTER_ANTIVIRUS = new Guid("{0xb1d1a169, 0xc54f, 0x4379, {0x81, 0xdb, 0xbe, 0xe7, 0xd8, 0x8d, 0x74, 0x54}}");
            public static readonly Guid FSFILTER_REPLICATION = new Guid("{0x48d3ebc4, 0x4cf8, 0x48ff, {0xb8, 0x69, 0x9c, 0x68, 0xad, 0x42, 0xeb, 0x9f}}");
            public static readonly Guid FSFILTER_CONTINUOUSBACKUP = new Guid("{0x71aa14f8, 0x6fad, 0x4622, {0xad, 0x77, 0x92, 0xbb, 0x9d, 0x7e, 0x69, 0x47}}");
            public static readonly Guid FSFILTER_CONTENTSCREENER = new Guid("{0x3e3f0674, 0xc83c, 0x4558, {0xbb, 0x26, 0x98, 0x20, 0xe1, 0xeb, 0xa5, 0xc5}}");
            public static readonly Guid FSFILTER_QUOTAMANAGEMENT = new Guid("{0x8503c911, 0xa6c7, 0x4919, {0x8f, 0x79, 0x50, 0x28, 0xf5, 0x86, 0x6b, 0x0c}}");
            public static readonly Guid FSFILTER_SYSTEMRECOVERY = new Guid("{0x2db15374, 0x706e, 0x4131, {0xa0, 0xc7, 0xd7, 0xc7, 0x8e, 0xb0, 0x28, 0x9a}}");
            public static readonly Guid FSFILTER_CFSMETADATASERVER = new Guid("{0xcdcf0939, 0xb75b, 0x4630, {0xbf, 0x76, 0x80, 0xf7, 0xba, 0x65, 0x58, 0x84}}");
            public static readonly Guid FSFILTER_HSM = new Guid("{0xd546500a, 0x2aeb, 0x45f6, {0x94, 0x82, 0xf4, 0xb1, 0x79, 0x9c, 0x31, 0x77}}");
            public static readonly Guid FSFILTER_COMPRESSION = new Guid("{0xf3586baf, 0xb5aa, 0x49b5, {0x8d, 0x6c, 0x05, 0x69, 0x28, 0x4c, 0x63, 0x9f}}");
            public static readonly Guid FSFILTER_ENCRYPTION = new Guid("{0xa0a701c0, 0xa511, 0x42ff, {0xaa, 0x6c, 0x06, 0xdc, 0x03, 0x95, 0x57, 0x6f}}");
            public static readonly Guid FSFILTER_PHYSICALQUOTAMANAGEMENT = new Guid("{0x6a0a8e78, 0xbba6, 0x4fc4, {0xa7, 0x09, 0x1e, 0x33, 0xcd, 0x09, 0xd6, 0x7e}}");
            public static readonly Guid FSFILTER_OPENFILEBACKUP = new Guid("{0xf8ecafa6, 0x66d1, 0x41a5, {0x89, 0x9b, 0x66, 0x58, 0x5d, 0x72, 0x16, 0xb7}}");
            public static readonly Guid FSFILTER_SECURITYENHANCER = new Guid("{0xd02bc3da, 0x0c8e, 0x4945, {0x9b, 0xd5, 0xf1, 0x88, 0x3c, 0x22, 0x6c, 0x8c}}");
            public static readonly Guid FSFILTER_COPYPROTECTION = new Guid("{0x89786ff1, 0x9c12, 0x402f, {0x9c, 0x9e, 0x17, 0x75, 0x3c, 0x7f, 0x43, 0x75}}");
            public static readonly Guid FSFILTER_SYSTEM = new Guid("{0x5d1b9aaa, 0x01e2, 0x46af, {0x84, 0x9f, 0x27, 0x2b, 0x3f, 0x32, 0x4c, 0x46}}");
            public static readonly Guid FSFILTER_INFRASTRUCTURE = new Guid("{0xe55fa6f9, 0x128c, 0x4d04, {0xab, 0xab, 0x63, 0x0c, 0x74, 0xb1, 0x45, 0x3a}}");
        }

        //http://stackoverflow.com/questions/304986/how-do-i-get-the-friendly-name-of-a-com-port-in-windows
        public enum SPDRP
        {
            DEVICEDESC = 0x00000000,
            HARDWAREID = 0x00000001,
            COMPATIBLEIDS = 0x00000002,
            NTDEVICEPATHS = 0x00000003,
            SERVICE = 0x00000004,
            CONFIGURATION = 0x00000005,
            CONFIGURATIONVECTOR = 0x00000006,
            CLASS = 0x00000007,
            CLASSGUID = 0x00000008,
            DRIVER = 0x00000009,
            CONFIGFLAGS = 0x0000000A,
            MFG = 0x0000000B,
            FRIENDLYNAME = 0x0000000C,
            LOCATION_INFORMATION = 0x0000000D,
            PHYSICAL_DEVICE_OBJECT_NAME = 0x0000000E,
            CAPABILITIES = 0x0000000F,
            UI_NUMBER = 0x00000010,
            UPPERFILTERS = 0x00000011,
            LOWERFILTERS = 0x00000012,
            MAXIMUM_PROPERTY = 0x00000013,
        }

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
                                    for (int j = 1; j < listView1.Columns.Count; j++)
                                    {
                                        lvi.SubItems.Add("");
                                    }

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
                                        lvi.SubItems[(int)LVC.StatusCol].Text = "Disabled";
                                    }
                                    else if (aDeviceInfo.service)
                                    {
                                        lvi.SubItems[(int)LVC.StatusCol].Text = "Service";
                                    }
                                    else if (aDeviceInfo.ghosted)
                                    {
                                        lvi.SubItems[(int)LVC.StatusCol].Text = "Ghosted";
                                    }
                                    else
                                    {
                                        lvi.SubItems[(int)LVC.StatusCol].Text = "Ok";
                                    }

                                    uint PropertyRegDataType;

                                    uint nBytes = 512;
                                    StringBuilder sb = new StringBuilder();
                                    sb.Length = (int)nBytes;
                                    uint RequiredSize = 0;

                                    SetupDi.SetupDiGetDeviceRegistryProperty(aDevInfoSet, ref aDeviceInfoData,
                                        (uint)SPDRP.FRIENDLYNAME, out PropertyRegDataType,
                                        sb, nBytes, out RequiredSize);

                                    lvi.SubItems[(int)LVC.DescriptionCol].Text = sb.ToString();

                                    //Remove Devices by Description
                                    StringCollection descrtoremove = ini.ReadSectionValues(DeviceKey);

                                    if (descrtoremove.Contains(aDeviceInfo.description.Trim()))
                                    {
                                        if (aDeviceInfo.ghosted && RemoveGhosts)
                                        {
                                            if (SetupDi.SetupDiRemoveDevice(aDevInfoSet, ref aDeviceInfoData))
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
                                                lvi.SubItems[(int)LVC.StatusCol].Text = "REMOVED";

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

        public void ReColorDevices(Boolean updatemax)
        {
            Int32 cnt = 0;
            Int32 watched = 0;
            Int32 ghosted = 0;
            Int32 removed = 0;

            using (IniFile ini = new IniFile(IniFileName))
            {
                //Remove Devices by Description.
                StringCollection descrtoremove = ini.ReadSectionValues(DeviceKey);
                StringCollection classtoremove = ini.ReadSectionValues(ClassKey);

                //Read Wildcards.
                LoadWildcards(ini);

                foreach (ListViewItem lvi in listView1.Items)
                {
                    String grp = lvi.Group.ToString().Trim();

                    lvi.SubItems[(int)LVC.MatchTypeCol].Text = "";

                    if (classtoremove.Contains(grp))
                    {
                        lvi.SubItems[(int)LVC.MatchTypeCol].Text += "[Class]";
                    }

                    if (descrtoremove.Contains(lvi.Text.Trim()))
                    {
                        lvi.SubItems[(int)LVC.MatchTypeCol].Text += "[Device]";
                    }

                    foreach (Wildcard w in wildcards)
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
            }

            toolStripStatusLabel1.Text = String.Format("{0} Device(s)", cnt - removed);
            toolStripStatusLabel2.Text = String.Format("{0} Filtered", watched + ghosted);
            toolStripStatusLabel3.Text = String.Format("{0} to be removed", ghosted);
        }

        private void LoadWildcards(IniFile ini)
        {
            StringDictionary sd = ini.ReadSection(WildcardKey);
            sd.Remove("Count");

            wildcards.Clear();
            foreach (String key in sd.Keys)
            {
                wildcards.Add(new Wildcard(ini.ReadString(WildcardKey, key, "xyyz")));
            }
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

                removeToolStripMenuItem.DropDownItems.Clear();

                foreach (Wildcard wildcard in wildcards)
                {
                    removeToolStripMenuItem.DropDownItems.Add(wildcard.Pattern.Replace("&", "&&"), null, RemoveWildcardToolStripMenuItem_Click).Tag = wildcard.Pattern;
                }

                removeToolStripMenuItem.Enabled = (removeToolStripMenuItem.DropDownItems.Count != 0);
            }
        }

        private void AddWildCardMnu_Click(object sender, EventArgs e)
        {
            String Device = String.Empty;

            if (listView1.SelectedItems.Count != 0)
            {
                using (IniFile ini = new IniFile(IniFileName))
                {
                    Device = listView1.SelectedItems[0].Text;

                    using (InputDialog dlg = new InputDialog("Wildcard", "Enter WildCard", Device))
                    {
                        if (dlg.ShowDialog() == DialogResult.OK)
                        {
                            //TODO Try / Except.
                            wildcards.Add(new Wildcard(dlg.Input));

                            SaveWildcards(ini);

                            ReColorDevices(false);
                        }
                    }
                }
            }
        }

        private void SaveWildcards(IniFile ini)
        {
            ini.EraseSection(WildcardKey);
            ini.WriteInteger(WildcardKey, "Count", wildcards.Count);
            foreach (Wildcard w in wildcards)
            {
                ini.WriteString(WildcardKey,
                    "item_" + wildcards.IndexOf(w).ToString(), w.Pattern);
            }
            ini.UpdateFile();
        }

        private void RemoveWildcardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String pattern = (String)((ToolStripMenuItem)sender).Tag;

            for (Int32 i = wildcards.Count - 1; i >= 0; i--)
            {
                if (wildcards[i].Pattern.Equals(pattern))
                {
                    wildcards.RemoveAt(i);
                }
            }

            SaveWildcards(new IniFile(null));
            LoadWildcards(new IniFile(null));

            ReColorDevices(false);
        }

    }
}