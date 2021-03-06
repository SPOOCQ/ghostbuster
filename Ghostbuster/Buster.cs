﻿#region License

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
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.IO;
    using System.Text;
    using System.Linq;
    using System.Windows.Forms;

    using GhostBuster;

    using Swiss;

    using HDEVINFO = System.IntPtr;
    using System.Runtime.InteropServices;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.Security.Principal;
    using System.Collections.ObjectModel;
    using System.Text.RegularExpressions;

    #region Enumerations

    /// <summary>
    /// Needed to retrieve the Friendly Name.
    /// 
    /// See http://stackoverflow.com/questions/304986/how-do-i-get-the-friendly-name-of-a-com-port-in-
    /// windows
    /// 
    /// See http://msdn.microsoft.com/en-us/library/windows/hardware/ff542360(v=vs.85).aspx Look into
    /// SPDRP_BASE_CONTAINERID (missing)
    /// </summary>
    public enum SPDRP
    {
        /// <summary>
        /// An enum constant representing the devicedesc option.
        /// </summary>
        [Description("Device Description")]
        DEVICEDESC = 0x00000000,

        /// <summary>
        /// An enum constant representing the hardwareid option.
        /// </summary>
        [Description("Hardware Id")]
        HARDWAREID = 0x00000001,

        /// <summary>
        /// An enum constant representing the compatibleids option.
        /// </summary>
        [Description("Compatiblel ID's")]
        COMPATIBLEIDS = 0x00000002,

        /// <summary>
        /// An enum constant representing the ntdevicepaths option.
        /// </summary>
        [Description("NT Device Paths")]
        NTDEVICEPATHS = 0x00000003,

        /// <summary>
        /// An enum constant representing the service option.
        /// </summary>
        [Description("Service")]
        SERVICE = 0x00000004,

        /// <summary>
        /// An enum constant representing the configuration option.
        /// </summary>
        [Description("Configuration")]
        CONFIGURATION = 0x00000005,

        /// <summary>
        /// An enum constant representing the configurationvector option.
        /// </summary>
        [Description("Configuration Vector")]
        CONFIGURATIONVECTOR = 0x00000006,

        /// <summary>
        /// An enum constant representing the class option.
        /// </summary>
        [Description("Class")]
        CLASS = 0x00000007,

        /// <summary>
        /// An enum constant representing the classguid option.
        /// </summary>
        [Description("Class GUID")]
        CLASSGUID = 0x00000008,

        /// <summary>
        /// An enum constant representing the driver option.
        /// </summary>
        [Description("Driver")]
        DRIVER = 0x00000009,

        /// <summary>
        /// An enum constant representing the configflags option.
        /// </summary>
        [Description("Config Flags")]
        CONFIGFLAGS = 0x0000000A,

        /// <summary>
        /// An enum constant representing the mfg option.
        /// </summary>
        [Description("Manufacturer")]
        MFG = 0x0000000B,

        /// <summary>
        /// An enum constant representing the friendlyname option.
        /// </summary>
        [Description("Friendly Name")]
        FRIENDLYNAME = 0x0000000C,

        /// <summary>
        /// An enum constant representing the location information option.
        /// </summary>
        [Description("Location Information")]
        LOCATION_INFORMATION = 0x0000000D,

        /// <summary>
        /// An enum constant representing the physical device location option.
        /// </summary>
        [Description("Physical Device Location")]
        PHYSICAL_DEVICE_LOCATION = 0x0000000E,

        /// <summary>
        /// An enum constant representing the capabilities option.
        /// </summary>
        [Description("Capabilities")]
        CAPABILITIES = 0x0000000F,

        /// <summary>
        /// An enum constant representing the number option.
        /// </summary>
        [Description("UI Number")]
        UI_NUMBER = 0x00000010,

        /// <summary>
        /// An enum constant representing the upperfilters option.
        /// </summary>
        [Description("Upper Filters")]
        UPPERFILTERS = 0x00000011,

        /// <summary>
        /// An enum constant representing the lowerfilters option.
        /// </summary>
        [Description("Lower Filters")]
        LOWERFILTERS = 0x00000012,

        /// <summary>
        /// An enum constant representing the bustypeguid option.
        /// </summary>
        [Description("BusTypeGUID (R)")]
        BUSTYPEGUID = 0x00000013,

        /// <summary>
        /// An enum constant representing the legacybustype option.
        /// </summary>
        [Description("LegacyBusType (R)")]
        LEGACYBUSTYPE = 0x00000014,

        /// <summary>
        /// An enum constant representing the busnumber option.
        /// </summary>
        [Description("BusNumber (R)")]
        BUSNUMBER = 0x00000015,

        /// <summary>
        /// An enum constant representing the enumerator name option.
        /// </summary>
        [Description("Enumerator Name (R)")]
        ENUMERATOR_NAME = 0x00000016,

        /// <summary>
        /// An enum constant representing the security option.
        /// </summary>
        [Description("Security (R/W, binary form)")]
        SECURITY = 0x00000017,

        /// <summary>
        /// An enum constant representing the security sds option.
        /// </summary>
        [Description("Security (W, SDS form)")]
        SECURITY_SDS = 0x00000018,

        /// <summary>
        /// An enum constant representing the devtype option.
        /// </summary>
        [Description("Device Type (R/W)")]
        DEVTYPE = 0x00000019,

        /// <summary>
        /// An enum constant representing the exclusive option.
        /// </summary>
        [Description("Device is exclusive-access (R/W)")]
        EXCLUSIVE = 0x0000001A,

        /// <summary>
        /// An enum constant representing the characteristics option.
        /// </summary>
        [Description("Device Characteristics (R/W)")]
        CHARACTERISTICS = 0x0000001B,

        /// <summary>
        /// An enum constant representing the address option.
        /// </summary>
        [Description("Device Address (R)")]
        ADDRESS = 0x0000001C,

        /// <summary>
        /// An enum constant representing the number description format option.
        /// </summary>
        [Description("UiNumberDescFormat (R/W)")]
        UI_NUMBER_DESC_FORMAT = 0X0000001D,

        /// <summary>
        /// An enum constant representing the device power data option.
        /// </summary>
        [Description("Device Power Data (R)")]
        DEVICE_POWER_DATA = 0x0000001E,

        /// <summary>
        /// An enum constant representing the removal policy option.
        /// </summary>
        [Description("Removal Policy (R)")]
        REMOVAL_POLICY = 0x0000001F,

        /// <summary>
        /// An enum constant representing the removal policy hardware default option.
        /// </summary>
        [Description("Hardware Removal Policy (R)")]
        REMOVAL_POLICY_HW_DEFAULT = 0x00000020,

        /// <summary>
        /// An enum constant representing the removal policy override option.
        /// </summary>
        [Description("Removal Policy Override (RW)")]
        REMOVAL_POLICY_OVERRIDE = 0x00000021,

        /// <summary>
        /// An enum constant representing the install state option.
        /// </summary>
        [Description("Device Install State (R)")]
        INSTALL_STATE = 0x00000022,

        /// <summary>
        /// An enum constant representing the location paths option.
        /// </summary>
        [Description("Device Location Paths (R)")]
        LOCATION_PATHS = 0x00000023,

        /// <summary>
        /// An enum constant representing the base containerid option.
        /// </summary>
        [Description("Base Container Id")]
        BASE_CONTAINERID = 0x00000024,

        /// <summary>
        /// An enum constant representing the maximum property option.
        /// </summary>
        [Description("")]
        MAXIMUM_PROPERTY = 0x00000025,
    }

    #endregion Enumerations

    /// <summary>
    /// A buster.
    /// </summary>
    public class Buster : IDisposable
    {
        #region Fields

        /// <summary>
        /// The title.
        /// </summary>
        public static readonly String S_TITLE = "GhostBuster";

        /// <summary>
        /// Section Name in IniFile.
        /// </summary>
        public const String ClassKey = "Class.RemoveIfGosted";

        /// <summary>
        /// Section Name in IniFile.
        /// </summary>
        public const String DeviceKey = "Descr.RemoveIfGosted";

        /// <summary>
        /// Section Name in IniFile.
        /// </summary>
        public const String WildcardKey = "Wildcards";

        /// <summary>
        /// A Handle.
        /// </summary>
        internal static HDEVINFO DevInfoSet;

        /// <summary>
        /// The Classes to remove.
        /// </summary>
        public static StringCollection Classes = new StringCollection();

        /// <summary>
        /// The Devices to remove.
        /// </summary>
        public static StringCollection Devices = new StringCollection();

        /// <summary>
        /// The Devices. //.
        /// </summary>
        public static ObservableCollection<HwEntry> HwEntries = new ObservableCollection<HwEntry>();

        /// <summary>
        /// IniFileName (Should Automatically use %AppData% when neccesary)!
        /// </summary>
        public static String IniFileName = String.Empty;

        /// <summary>
        /// The serial ports.
        /// </summary>
        public static Dictionary<String, Int32> SerialPorts = new Dictionary<String, Int32>();

        /// <summary>
        /// The Wildcards to remove.
        /// </summary>
        public static List<Wildcard> Wildcards = new List<Wildcard>();

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Constructor.
        /// 
        /// Defaults to AppIni.
        /// </summary>
        public Buster()
            : this(IniFile.AppIni)
        {
            //Nothing
        }

        /// <summary>
        /// Constructor.
        /// 
        /// Loads static data from am Ini File.
        /// </summary>
        ///
        /// <param name="IniFileName"> . </param>
        public Buster(String IniFileName)
        {
            //! Take this one from %appdata% or the CommandLine.
            Buster.IniFileName = IniFileName;

            LoadDevicesAndClasses(new IniFile(IniFileName));

            LoadWildcards(new IniFile(IniFileName));
        }

        #endregion Constructors

        #region Methods

        /// <summary>
        /// Adds the class key.
        /// </summary>
        ///
        /// <param name="Class"> The class. </param>
        public static void AddClassKey(String Class)
        {
            using (IniFile ini = new IniFile(Buster.IniFileName))
            {
                ini.WriteString(Buster.ClassKey,
                    String.Format("item_{0}", ini.ReadInteger(Buster.ClassKey, "Count", 0)),
                    Class);
                ini.WriteInteger(Buster.ClassKey,
                    "Count",
                    ini.ReadInteger(Buster.ClassKey, "Count", 0) + 1);
                ini.UpdateFile();
            }

            LoadDevicesAndClasses(new IniFile(IniFileName));
        }

        /// <summary>
        /// Adds a device key.
        /// </summary>
        ///
        /// <param name="Device"> The device. </param>
        public static void AddDeviceKey(String Device)
        {
            using (IniFile ini = new IniFile(Buster.IniFileName))
            {
                ini.WriteString(Buster.DeviceKey,
                    String.Format("item_{0}", ini.ReadInteger(Buster.DeviceKey, "Count", 0)),
                    Device);
                ini.WriteInteger(Buster.DeviceKey,
                    "Count",
                    ini.ReadInteger(Buster.DeviceKey, "Count", 0) + 1);
                ini.UpdateFile();
            }

            LoadDevicesAndClasses(new IniFile(IniFileName));
        }

        /// <summary>
        /// Enumerate all devices and optionally uninstall ghosted ones.
        /// </summary>
        public static void Enumerate()
        {
            SetupDi.SP_DEVINFO_DATA aDeviceInfoData = new SetupDi.SP_DEVINFO_DATA();

            HwEntries.Clear();
            SerialPorts.Clear();

            try
            {
                //Cache all HKLM Services Key Names and DisplayNames
                SetupDi.EnumServices();

                DevInfoSet = SetupDi.SetupDiGetClassDevs(ref SetupDi.NullGuid, 0, IntPtr.Zero, (uint)SetupDi.DIGCF.DIGCF_ALLCLASSES);

                if (DevInfoSet != (IntPtr)SetupDi.INVALID_HANDLE_VALUE)
                {
                    aDeviceInfoData.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(aDeviceInfoData);

                    Int32 i = 0;

                    using (IniFile ini = new IniFile(IniFileName))
                    {
                        while (SetupDi.SetupDiEnumDeviceInfo(DevInfoSet, i, ref aDeviceInfoData))
                        {
                            SetupDi.DeviceInfo aDeviceInfo = new SetupDi.DeviceInfo();

                            //DeviceClass is used for grouping items...
                            SetupDi.GetClassDescriptionFromGuid(ref aDeviceInfo, aDeviceInfoData.ClassGuid);

                            //veg: Gives Exceptions, remove and return error name...
                            try
                            {
                                //TODO: Retrieving DeviceName fails.
                                SetupDi.GetDeviceDescription(ref aDeviceInfo, DevInfoSet, aDeviceInfoData);
                                SetupDi.GetDeviceRegistryKey(ref aDeviceInfo, DevInfoSet, aDeviceInfoData);
                                SetupDi.GetDeviceName(ref aDeviceInfo, DevInfoSet, aDeviceInfoData);

                                SetupDi.GetDeviceStatus(ref aDeviceInfo, DevInfoSet, ref aDeviceInfoData);

                                uint PropertyRegDataType;

                                uint nBytes = 512;
                                StringBuilder sb = new StringBuilder();
                                sb.Length = (int)nBytes;
                                uint RequiredSize = 0;

                                SetupDi.SetupDiGetDeviceRegistryProperty(DevInfoSet, ref aDeviceInfoData,
                                    (uint)SPDRP.FRIENDLYNAME, out PropertyRegDataType,
                                    sb, nBytes, out RequiredSize);

                                // Debug.WriteLine(SetupDi.GetProviderName(DevInfoSet, ref aDeviceInfoData));

                                HwEntries.Add(new HwEntry(DevInfoSet, aDeviceInfo, aDeviceInfoData, sb.ToString()));

                                SetupDi.SetupDiGetDeviceRegistryProperty(DevInfoSet, ref aDeviceInfoData,
                                  (uint)SPDRP.MFG, out PropertyRegDataType,
                                  sb, nBytes, out RequiredSize);
                                // Debug.WriteLine(sb);

                            }
                            finally
                            {
                                //Nothing
                            }

                            i++;
                        }
                    }
                }
            }
            finally
            {
                //Nothing
            }

            Debug.Print("[SerialPorts]");
            foreach (KeyValuePair<String, Int32> kvp in SerialPorts)
            {
                Debug.Print("{0}={1}", kvp.Value, kvp.Key);
            }
        }

        /// <summary>
        /// Loads devices and classes.
        /// </summary>
        ///
        /// <param name="ini"> The initialise. </param>
        public static void LoadDevicesAndClasses(IniFile ini)
        {
            Devices = ini.ReadSectionValues(DeviceKey);
            Classes = ini.ReadSectionValues(ClassKey);
        }

        /// <summary>
        /// Removes the class key described by Class.
        /// </summary>
        ///
        /// <param name="Class"> The class. </param>
        public static void RemoveClassKey(String Class)
        {
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

            LoadDevicesAndClasses(new IniFile(IniFileName));
        }

        /// <summary>
        /// Removes the device key described by Device.
        /// </summary>
        ///
        /// <param name="Device"> The device. </param>
        public static void RemoveDeviceKey(String Device)
        {
            using (IniFile ini = new IniFile(Buster.IniFileName))
            {
                foreach (DictionaryEntry de in ini.ReadSection(Buster.DeviceKey))
                {
                    if (de.Value.ToString() == Device)
                    {
                        ini.DeleteKey(Buster.DeviceKey, de.Key.ToString());
                    }
                }
                ini.UpdateFile();
            }

            LoadDevicesAndClasses(new IniFile(IniFileName));
        }

        /// <summary>
        /// Adds a wild card.
        /// </summary>
        ///
        /// <param name="wildcard"> The wildcard. </param>
        internal static void AddWildCard(String wildcard)
        {
            Wildcards.Add(new Wildcard(wildcard));

            SaveWildcards(new IniFile(IniFileName));
        }

        /// <summary>
        /// Removes the wild card described by pattern.
        /// </summary>
        ///
        /// <param name="pattern"> Specifies the pattern. </param>
        internal static void RemoveWildCard(String pattern)
        {
            for (Int32 i = Wildcards.Count - 1; i >= 0; i--)
            {
                if (Wildcards[i].Pattern.Equals(pattern))
                {
                    Wildcards.RemoveAt(i);
                }
            }

            SaveWildcards(new IniFile(IniFileName));
        }

        /// <summary>
        /// Loads the wildcards.
        /// </summary>
        ///
        /// <param name="ini"> The initialise. </param>
        private static void LoadWildcards(IniFile ini)
        {
            StringDictionary sd = ini.ReadSection(WildcardKey);
            sd.Remove("Count");

            Wildcards.Clear();
            foreach (String key in sd.Keys)
            {
                Wildcards.Add(new Wildcard(ini.ReadString(WildcardKey, key, "xyyz")));
            }
        }

        /// <summary>
        /// Saves the wildcards.
        /// </summary>
        ///
        /// <param name="ini"> The initialise. </param>
        private static void SaveWildcards(IniFile ini)
        {
            ini.EraseSection(WildcardKey);
            ini.WriteInteger(WildcardKey, "Count", Wildcards.Count);
            foreach (Wildcard w in Wildcards)
            {
                ini.WriteString(WildcardKey,
                    "item_" + Wildcards.IndexOf(w).ToString(), w.Pattern);
            }
            ini.UpdateFile();
        }

        /// <summary>
        /// Query if this object is admin.
        /// </summary>
        ///
        /// <returns>
        /// true if admin, false if not.
        /// </returns>
        public static bool IsAdmin()
        {
            WindowsIdentity user = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(user);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        #endregion Methods

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged
        /// resources.
        /// </summary>
        public void Dispose()
        {
            Devices.Clear();
            Classes.Clear();
            Wildcards.Clear();

            HwEntries.Clear();
        }

        /// <summary>
        /// Writes to event log.
        /// </summary>
        ///
        /// <param name="message"> The message. </param>
        /// <param name="elet">    The elet. </param>
        public static void WriteToEventLog(string message, EventLogEntryType elet)
        {
            string cs = S_TITLE;

            Debug.WriteLine(message);

            EventLog log = new EventLog();

            if (Buster.IsAdmin())
            {
                //if (!EventLog.SourceExists(cs))
                //{
                //    EventLog.CreateEventSource(cs, cs);
                //}

                if (EventLog.SourceExists(cs))
                {
                    log.Source = cs;
                    log.EnableRaisingEvents = true;
                    log.WriteEntry(message, elet);
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// Needed to retrieve the Friendly Name.
    /// </summary>
    public class GUID_DEVCLASS
    {
        #region Fields

        /// <summary>
        /// The adapter.
        /// </summary>
        public static readonly Guid ADAPTER = new Guid("{0x4d36e964, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The apmsupport.
        /// </summary>
        public static readonly Guid APMSUPPORT = new Guid("{0xd45b1c18, 0xc8fa, 0x11d1, {0x9f, 0x77, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");

        /// <summary>
        /// The avc.
        /// </summary>
        public static readonly Guid AVC = new Guid("{0xc06ff265, 0xae09, 0x48f0, {0x81, 0x2c, 0x16, 0x75, 0x3d, 0x7c, 0xba, 0x83}}");

        /// <summary>
        /// The battery.
        /// </summary>
        public static readonly Guid BATTERY = new Guid("{0x72631e54, 0x78a4, 0x11d0, {0xbc, 0xf7, 0x00, 0xaa, 0x00, 0xb7, 0xb3, 0x2a}}");

        /// <summary>
        /// The biometric.
        /// </summary>
        public static readonly Guid BIOMETRIC = new Guid("{0x53d29ef7, 0x377c, 0x4d14, {0x86, 0x4b, 0xeb, 0x3a, 0x85, 0x76, 0x93, 0x59}}");

        /// <summary>
        /// The bluetooth.
        /// </summary>
        public static readonly Guid BLUETOOTH = new Guid("{0xe0cbf06c, 0xcd8b, 0x4647, {0xbb, 0x8a, 0x26, 0x3b, 0x43, 0xf0, 0xf9, 0x74}}");

        /// <summary>
        /// The cdrom.
        /// </summary>
        public static readonly Guid CDROM = new Guid("{0x4d36e965, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The computer.
        /// </summary>
        public static readonly Guid COMPUTER = new Guid("{0x4d36e966, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The decoder.
        /// </summary>
        public static readonly Guid DECODER = new Guid("{0x6bdd1fc2, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The diskdrive.
        /// </summary>
        public static readonly Guid DISKDRIVE = new Guid("{0x4d36e967, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The display.
        /// </summary>
        public static readonly Guid DISPLAY = new Guid("{0x4d36e968, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The fourth dot.
        /// </summary>
        public static readonly Guid DOT4 = new Guid("{0x48721b56, 0x6795, 0x11d2, {0xb1, 0xa8, 0x00, 0x80, 0xc7, 0x2e, 0x74, 0xa2}}");

        /// <summary>
        /// The dot 4 print.
        /// </summary>
        public static readonly Guid DOT4PRINT = new Guid("{0x49ce6ac8, 0x6f86, 0x11d2, {0xb1, 0xe5, 0x00, 0x80, 0xc7, 0x2e, 0x74, 0xa2}}");

        /// <summary>
        /// The fourth enum 139.
        /// </summary>
        public static readonly Guid ENUM1394 = new Guid("{0xc459df55, 0xdb08, 0x11d1, {0xb0, 0x09, 0x00, 0xa0, 0xc9, 0x08, 0x1f, 0xf6}}");

        /// <summary>
        /// The fdc.
        /// </summary>
        public static readonly Guid FDC = new Guid("{0x4d36e969, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The floppydisk.
        /// </summary>
        public static readonly Guid FLOPPYDISK = new Guid("{0x4d36e980, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The fsfilter activitymonitor.
        /// </summary>
        public static readonly Guid FSFILTER_ACTIVITYMONITOR = new Guid("{0xb86dff51, 0xa31e, 0x4bac, {0xb3, 0xcf, 0xe8, 0xcf, 0xe7, 0x5c, 0x9f, 0xc2}}");

        /// <summary>
        /// The fsfilter antivirus.
        /// </summary>
        public static readonly Guid FSFILTER_ANTIVIRUS = new Guid("{0xb1d1a169, 0xc54f, 0x4379, {0x81, 0xdb, 0xbe, 0xe7, 0xd8, 0x8d, 0x74, 0x54}}");

        /// <summary>
        /// The fsfilter cfsmetadataserver.
        /// </summary>
        public static readonly Guid FSFILTER_CFSMETADATASERVER = new Guid("{0xcdcf0939, 0xb75b, 0x4630, {0xbf, 0x76, 0x80, 0xf7, 0xba, 0x65, 0x58, 0x84}}");

        /// <summary>
        /// The fsfilter compression.
        /// </summary>
        public static readonly Guid FSFILTER_COMPRESSION = new Guid("{0xf3586baf, 0xb5aa, 0x49b5, {0x8d, 0x6c, 0x05, 0x69, 0x28, 0x4c, 0x63, 0x9f}}");

        /// <summary>
        /// The fsfilter contentscreener.
        /// </summary>
        public static readonly Guid FSFILTER_CONTENTSCREENER = new Guid("{0x3e3f0674, 0xc83c, 0x4558, {0xbb, 0x26, 0x98, 0x20, 0xe1, 0xeb, 0xa5, 0xc5}}");

        /// <summary>
        /// The fsfilter continuousbackup.
        /// </summary>
        public static readonly Guid FSFILTER_CONTINUOUSBACKUP = new Guid("{0x71aa14f8, 0x6fad, 0x4622, {0xad, 0x77, 0x92, 0xbb, 0x9d, 0x7e, 0x69, 0x47}}");

        /// <summary>
        /// The fsfilter copyprotection.
        /// </summary>
        public static readonly Guid FSFILTER_COPYPROTECTION = new Guid("{0x89786ff1, 0x9c12, 0x402f, {0x9c, 0x9e, 0x17, 0x75, 0x3c, 0x7f, 0x43, 0x75}}");

        /// <summary>
        /// The fsfilter encryption.
        /// </summary>
        public static readonly Guid FSFILTER_ENCRYPTION = new Guid("{0xa0a701c0, 0xa511, 0x42ff, {0xaa, 0x6c, 0x06, 0xdc, 0x03, 0x95, 0x57, 0x6f}}");

        /// <summary>
        /// The fsfilter hsm.
        /// </summary>
        public static readonly Guid FSFILTER_HSM = new Guid("{0xd546500a, 0x2aeb, 0x45f6, {0x94, 0x82, 0xf4, 0xb1, 0x79, 0x9c, 0x31, 0x77}}");

        /// <summary>
        /// The fsfilter infrastructure.
        /// </summary>
        public static readonly Guid FSFILTER_INFRASTRUCTURE = new Guid("{0xe55fa6f9, 0x128c, 0x4d04, {0xab, 0xab, 0x63, 0x0c, 0x74, 0xb1, 0x45, 0x3a}}");

        /// <summary>
        /// The fsfilter openfilebackup.
        /// </summary>
        public static readonly Guid FSFILTER_OPENFILEBACKUP = new Guid("{0xf8ecafa6, 0x66d1, 0x41a5, {0x89, 0x9b, 0x66, 0x58, 0x5d, 0x72, 0x16, 0xb7}}");

        /// <summary>
        /// The fsfilter physicalquotamanagement.
        /// </summary>
        public static readonly Guid FSFILTER_PHYSICALQUOTAMANAGEMENT = new Guid("{0x6a0a8e78, 0xbba6, 0x4fc4, {0xa7, 0x09, 0x1e, 0x33, 0xcd, 0x09, 0xd6, 0x7e}}");

        /// <summary>
        /// The fsfilter quotamanagement.
        /// </summary>
        public static readonly Guid FSFILTER_QUOTAMANAGEMENT = new Guid("{0x8503c911, 0xa6c7, 0x4919, {0x8f, 0x79, 0x50, 0x28, 0xf5, 0x86, 0x6b, 0x0c}}");

        /// <summary>
        /// The fsfilter replication.
        /// </summary>
        public static readonly Guid FSFILTER_REPLICATION = new Guid("{0x48d3ebc4, 0x4cf8, 0x48ff, {0xb8, 0x69, 0x9c, 0x68, 0xad, 0x42, 0xeb, 0x9f}}");

        /// <summary>
        /// The fsfilter securityenhancer.
        /// </summary>
        public static readonly Guid FSFILTER_SECURITYENHANCER = new Guid("{0xd02bc3da, 0x0c8e, 0x4945, {0x9b, 0xd5, 0xf1, 0x88, 0x3c, 0x22, 0x6c, 0x8c}}");

        /// <summary>
        /// The fsfilter system.
        /// </summary>
        public static readonly Guid FSFILTER_SYSTEM = new Guid("{0x5d1b9aaa, 0x01e2, 0x46af, {0x84, 0x9f, 0x27, 0x2b, 0x3f, 0x32, 0x4c, 0x46}}");

        /// <summary>
        /// The fsfilter systemrecovery.
        /// </summary>
        public static readonly Guid FSFILTER_SYSTEMRECOVERY = new Guid("{0x2db15374, 0x706e, 0x4131, {0xa0, 0xc7, 0xd7, 0xc7, 0x8e, 0xb0, 0x28, 0x9a}}");

        /// <summary>
        /// The fsfilter undelete.
        /// </summary>
        public static readonly Guid FSFILTER_UNDELETE = new Guid("{0xfe8f1572, 0xc67a, 0x48c0, {0xbb, 0xac, 0x0b, 0x5c, 0x6d, 0x66, 0xca, 0xfb}}");

        /// <summary>
        /// The GPS.
        /// </summary>
        public static readonly Guid GPS = new Guid("{0x6bdd1fc3, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The hdc.
        /// </summary>
        public static readonly Guid HDC = new Guid("{0x4d36e96a, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The hidclass.
        /// </summary>
        public static readonly Guid HIDCLASS = new Guid("{0x745a17a0, 0x74d3, 0x11d0, {0xb6, 0xfe, 0x00, 0xa0, 0xc9, 0x0f, 0x57, 0xda}}");

        /// <summary>
        /// The image.
        /// </summary>
        public static readonly Guid IMAGE = new Guid("{0x6bdd1fc6, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The infiniband.
        /// </summary>
        public static readonly Guid INFINIBAND = new Guid("{0x30ef7132, 0xd858, 0x4a0c, {0xac, 0x24, 0xb9, 0x02, 0x8a, 0x5c, 0xca, 0x3f}}");

        /// <summary>
        /// The infrared.
        /// </summary>
        public static readonly Guid INFRARED = new Guid("{0x6bdd1fc5, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The keyboard.
        /// </summary>
        public static readonly Guid KEYBOARD = new Guid("{0x4d36e96b, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The legacydriver.
        /// </summary>
        public static readonly Guid LEGACYDRIVER = new Guid("{0x8ecc055d, 0x047f, 0x11d1, {0xa5, 0x37, 0x00, 0x00, 0xf8, 0x75, 0x3e, 0xd1}}");

        /// <summary>
        /// The media.
        /// </summary>
        public static readonly Guid MEDIA = new Guid("{0x4d36e96c, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The medium changer.
        /// </summary>
        public static readonly Guid MEDIUM_CHANGER = new Guid("{0xce5939ae, 0xebde, 0x11d0, {0xb1, 0x81, 0x00, 0x00, 0xf8, 0x75, 0x3e, 0xc4}}");

        /// <summary>
        /// The modem.
        /// </summary>
        public static readonly Guid MODEM = new Guid("{0x4d36e96d, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The monitor.
        /// </summary>
        public static readonly Guid MONITOR = new Guid("{0x4d36e96e, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The mouse.
        /// </summary>
        public static readonly Guid MOUSE = new Guid("{0x4d36e96f, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The mtd.
        /// </summary>
        public static readonly Guid MTD = new Guid("{0x4d36e970, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The multifunction.
        /// </summary>
        public static readonly Guid MULTIFUNCTION = new Guid("{0x4d36e971, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The multiportserial.
        /// </summary>
        public static readonly Guid MULTIPORTSERIAL = new Guid("{0x50906cb8, 0xba12, 0x11d1, {0xbf, 0x5d, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");

        /// <summary>
        /// The net.
        /// </summary>
        public static readonly Guid NET = new Guid("{0x4d36e972, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The netclient.
        /// </summary>
        public static readonly Guid NETCLIENT = new Guid("{0x4d36e973, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The netservice.
        /// </summary>
        public static readonly Guid NETSERVICE = new Guid("{0x4d36e974, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The nettrans.
        /// </summary>
        public static readonly Guid NETTRANS = new Guid("{0x4d36e975, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The nodriver.
        /// </summary>
        public static readonly Guid NODRIVER = new Guid("{0x4d36e976, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The PCMCIA.
        /// </summary>
        public static readonly Guid PCMCIA = new Guid("{0x4d36e977, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The pnpprinters.
        /// </summary>
        public static readonly Guid PNPPRINTERS = new Guid("{0x4658ee7e, 0xf050, 0x11d1, {0xb6, 0xbd, 0x00, 0xc0, 0x4f, 0xa3, 0x72, 0xa7}}");

        /// <summary>
        /// The ports.
        /// </summary>
        public static readonly Guid PORTS = new Guid("{0x4d36e978, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The printer.
        /// </summary>
        public static readonly Guid PRINTER = new Guid("{0x4d36e979, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The printerupgrade.
        /// </summary>
        public static readonly Guid PRINTERUPGRADE = new Guid("{0x4d36e97a, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The processor.
        /// </summary>
        public static readonly Guid PROCESSOR = new Guid("{0x50127dc3, 0x0f36, 0x415e, {0xa6, 0xcc, 0x4c, 0xb3, 0xbe, 0x91, 0x0B, 0x65}}");

        /// <summary>
        /// The second sbp.
        /// </summary>
        public static readonly Guid SBP2 = new Guid("{0xd48179be, 0xec20, 0x11d1, {0xb6, 0xb8, 0x00, 0xc0, 0x4f, 0xa3, 0x72, 0xa7}}");

        /// <summary>
        /// The scsiadapter.
        /// </summary>
        public static readonly Guid SCSIADAPTER = new Guid("{0x4d36e97b, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The securityaccelerator.
        /// </summary>
        public static readonly Guid SECURITYACCELERATOR = new Guid("{0x268c95a1, 0xedfe, 0x11d3, {0x95, 0xc3, 0x00, 0x10, 0xdc, 0x40, 0x50, 0xa5}}");

        /// <summary>
        /// The smartcardreader.
        /// </summary>
        public static readonly Guid SMARTCARDREADER = new Guid("{0x50dd5230, 0xba8a, 0x11d1, {0xbf, 0x5d, 0x00, 0x00, 0xf8, 0x05, 0xf5, 0x30}}");

        /// <summary>
        /// The sound.
        /// </summary>
        public static readonly Guid SOUND = new Guid("{0x4d36e97c, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The system.
        /// </summary>
        public static readonly Guid SYSTEM = new Guid("{0x4d36e97d, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The tapedrive.
        /// </summary>
        public static readonly Guid TAPEDRIVE = new Guid("{0x6d807884, 0x7d21, 0x11cf, {0x80, 0x1c, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The unknown.
        /// </summary>
        public static readonly Guid UNKNOWN = new Guid("{0x4d36e97e, 0xe325, 0x11ce, {0xbf, 0xc1, 0x08, 0x00, 0x2b, 0xe1, 0x03, 0x18}}");

        /// <summary>
        /// The USB.
        /// </summary>
        public static readonly Guid USB = new Guid("{0x36fc9e60, 0xc465, 0x11cf, {0x80, 0x56, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00}}");

        /// <summary>
        /// The volume.
        /// </summary>
        public static readonly Guid VOLUME = new Guid("{0x71a27cdd, 0x812a, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The volumesnapshot.
        /// </summary>
        public static readonly Guid VOLUMESNAPSHOT = new Guid("{0x533c5b84, 0xec70, 0x11d2, {0x95, 0x05, 0x00, 0xc0, 0x4f, 0x79, 0xde, 0xaf}}");

        /// <summary>
        /// The wceusbs.
        /// </summary>
        public static readonly Guid WCEUSBS = new Guid("{0x25dbce51, 0x6c8f, 0x4a72, {0x8a, 0x6d, 0xb5, 0x4c, 0x2b, 0x4f, 0xc8, 0x35}}");

        /// <summary>
        /// http://stackoverflow.com/questions/304986/how-do-i-get-the-friendly-name-of-a-com-port-in-
        /// windows.
        /// </summary>
        public static readonly Guid _1394 = new Guid("{0x6bdd1fc1, 0x810f, 0x11d0, {0xbe, 0xc7, 0x08, 0x00, 0x2b, 0xe2, 0x09, 0x2f}}");

        /// <summary>
        /// The 1394 debug.
        /// </summary>
        public static readonly Guid _1394DEBUG = new Guid("{0x66f250d6, 0x7801, 0x4a64, {0xb1, 0x39, 0xee, 0xa8, 0x0a, 0x45, 0x0b, 0x24}}");

        /// <summary>
        /// The third 6188.
        /// </summary>
        public static readonly Guid _61883 = new Guid("{0x7ebefbc0, 0x3200, 0x11d2, {0xb4, 0xc2, 0x00, 0xa0, 0xc9, 0x69, 0x7d, 0x07}}");

        #endregion Fields
    }

    /// <summary>
    /// A Hardware Entry.
    /// </summary>
    public class HwEntry
    {
        #region Fields

        /// <summary>
        /// Needed for Removal.
        /// </summary>
        private SetupDi.SP_DEVINFO_DATA aDeviceInfoData;

        /// <summary>
        /// Needed for Removal.
        /// </summary>
        private HDEVINFO aDeviceInfoSet;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Creates a HwEntry to add to the list of devices considered for removal.
        /// </summary>
        ///
        /// <param name="aDeviceInfoSet">  Needed for Removal. </param>
        /// <param name="aDeviceInfo">     Needed for Name and Status. </param>
        /// <param name="aDeviceInfoData"> Needed for Removal. </param>
        /// <param name="FriendlyName">    The Friendly name if any. </param>
        public HwEntry(HDEVINFO aDeviceInfoSet, SetupDi.DeviceInfo aDeviceInfo, SetupDi.SP_DEVINFO_DATA aDeviceInfoData, String FriendlyName)
        {
            this.aDeviceInfoSet = aDeviceInfoSet;
            this.Description = aDeviceInfo.description;
            this.DeviceClass = aDeviceInfo.deviceclass;

            this.aDeviceInfoData = aDeviceInfoData;

            if (aDeviceInfo.disabled)
            {
                this.DeviceStatus = "Disabled";
            }
            else if (aDeviceInfo.service)
            {
                //! Make Sure Services are NEVER flagged as Ghosted by checking them first.
                using (IniFile ini = new IniFile(Buster.IniFileName))
                {
                    if (ini.ReadBool("Setup", "DoNotFlagServices", false) == false)
                    {
                        this.DeviceStatus = "Service";
                    }
                }
            }
            else if (aDeviceInfo.friendlyname.Equals(Guid.Empty.ToString()))
            {
                //! Make Sure the Null Guid Devices is NEVER flagged as Ghosted by checking them first.
                this.DeviceStatus = "System";
            }
            else if (aDeviceInfo.ghosted)
            {
                this.DeviceStatus = "Ghosted";
            }
            else
            {
                this.DeviceStatus = "Ok";
            }

            this.FriendlyName = FriendlyName;

            Properties = new Dictionary<String, String>();

            foreach (SPDRP s in Enum.GetValues(typeof(SPDRP)))
            {
                uint PropertyRegDataType;

                StringBuilder sb = new StringBuilder();
                uint RequiredSize = 0;

                //! PropertyBuffer [out, optional]
                //! A pointer to a buffer that receives the property that is being retrieved. 
                //! If this parameter is set to NULL, and PropertyBufferSize is also set to zero, 
                //! the function returns the required size for the buffer in RequiredSize.

                //See http://msdn.microsoft.com/en-us/library/windows/hardware/ff551967(v=vs.85).aspx 
                SetupDi.SetupDiGetDeviceRegistryProperty(aDeviceInfoSet, ref aDeviceInfoData, (uint)s, out PropertyRegDataType, null, 0, out RequiredSize);

                uint nBytes = RequiredSize;
                if (new SPDRP[] { SPDRP.BUSTYPEGUID, SPDRP.BASE_CONTAINERID, SPDRP.CLASSGUID }.Contains(s) && nBytes == 16)
                {
                    Byte[] buf = new Byte[16];

                    IntPtr unmanagedPointer = Marshal.AllocHGlobal(buf.Length);

                    SetupDi.SetupDiGetDeviceRegistryProperty(aDeviceInfoSet, ref aDeviceInfoData,
                         (uint)s, out PropertyRegDataType,
                         unmanagedPointer, nBytes, out RequiredSize);

                    Marshal.Copy(unmanagedPointer, buf, 0, buf.Length);

                    Marshal.FreeHGlobal(unmanagedPointer);

                    sb.Append("{" + new Guid(buf).ToString() + "}");
                }
                else
                {
                    sb.Length = (int)nBytes;
                    //RequiredSize = 0;

                    SetupDi.SetupDiGetDeviceRegistryProperty(aDeviceInfoSet, ref aDeviceInfoData,
                          (uint)s, out PropertyRegDataType,
                          sb, nBytes, out RequiredSize);
                }

                String descr = EnumExtensions.GetDescription<SPDRP>((SPDRP)Enum.Parse(typeof(SPDRP), s.ToString()));

                switch (s)
                {
                    //! StringList.
                    case SPDRP.COMPATIBLEIDS:
                    case SPDRP.LOCATION_PATHS:
                    case SPDRP.LOWERFILTERS:
                    case SPDRP.UPPERFILTERS:
                    case SPDRP.NTDEVICEPATHS:
                        {
                            //! Correct for 2 Byte UniCode
                            if (RequiredSize != 0 && (RequiredSize / 2) - 1 != sb.Length)
                            {
                                IntPtr ptrBuffer = new IntPtr();
                                ptrBuffer = Marshal.AllocHGlobal((int)nBytes);

                                SetupDi.SetupDiGetDeviceRegistryProperty(aDeviceInfoSet, ref aDeviceInfoData,
                                               (uint)s, out PropertyRegDataType,
                                               ptrBuffer, nBytes, out RequiredSize);

                                List<String> vals = new List<String>();

                                String val = Marshal.PtrToStringAuto(ptrBuffer);
                                do
                                {
                                    // Debug.Print("{0}", val);
                                    vals.Add(val);

                                    //! Advance & Correct for 2 Byte UniCode
                                    ptrBuffer = new IntPtr(ptrBuffer.ToInt64() + val.Length * 2 + 2);

                                    val = Marshal.PtrToStringAuto(ptrBuffer);
                                } while (val.Length > 0);

                                SafeAddOrAppendProperty(s, String.Join("; ", vals.ToArray()));
                                //if (!Properties.ContainsKey(s.ToString()))
                                //{
                                //    Properties.Add(s.ToString(), String.Join("; ", vals.ToArray()));
                                //}
                            }
                            else
                            {
                                SafeAddOrAppendProperty(s, sb);
                            }
                        }
                        break;

                    //! String
                    case SPDRP.CLASS:
                    case SPDRP.DEVICEDESC:
                    case SPDRP.DRIVER:
                    case SPDRP.ENUMERATOR_NAME:
                    case SPDRP.FRIENDLYNAME:
                    case SPDRP.HARDWAREID:
                    case SPDRP.LOCATION_INFORMATION:
                    case SPDRP.MFG:
                    case SPDRP.PHYSICAL_DEVICE_LOCATION:
                    case SPDRP.SERVICE:
                    case SPDRP.UI_NUMBER_DESC_FORMAT:
                        SafeAddOrAppendProperty(s, sb);
                        break;

                    //! ? (seem to be empty)
                    case SPDRP.CONFIGURATION:
                    case SPDRP.CONFIGURATIONVECTOR:
                        SafeAddOrAppendProperty(s, sb);
                        break;

                    //! Guid
                    case SPDRP.BASE_CONTAINERID:
                    case SPDRP.BUSTYPEGUID:
                    case SPDRP.CLASSGUID:
                        SafeAddOrAppendProperty(s, sb);
                        break;

                    //! Struct?
                    case SPDRP.DEVICE_POWER_DATA:
                    // Skip for now: contains CM_POWER_DATA struct?
                    case SPDRP.SECURITY:
                    case SPDRP.SECURITY_SDS:
                        //SafeAddOrAppendProperty(s, "structure or binary data");
                        break;

                    //! Boolean
                    case SPDRP.EXCLUSIVE:
                        if (sb.Length == 1)
                        {
                            SafeAddOrAppendProperty(s, sb[0] == 0 ? true.ToString() : false.ToString());
                        }
                        break;

                    //! Int32
                    case SPDRP.ADDRESS:
                    case SPDRP.BUSNUMBER:
                    case SPDRP.CAPABILITIES:
                    case SPDRP.CHARACTERISTICS:
                    case SPDRP.CONFIGFLAGS:
                    case SPDRP.DEVTYPE:
                    case SPDRP.INSTALL_STATE:
                    case SPDRP.LEGACYBUSTYPE:
                    case SPDRP.REMOVAL_POLICY:
                    case SPDRP.REMOVAL_POLICY_HW_DEFAULT:
                    case SPDRP.REMOVAL_POLICY_OVERRIDE:
                    case SPDRP.UI_NUMBER:
                        if (sb.Length == 1 && !Properties.ContainsKey(s.ToString()))
                        {
                            SafeAddOrAppendProperty(s, ((Int32)sb[0]).ToString("X4"));
                        }
                        break;
                    case SPDRP.MAXIMUM_PROPERTY:
                        //Should not have a value
                        break;
                    default:
                        Debug.Print("Missing SPDRP: {0}", s);
                        SafeAddOrAppendProperty(s, sb);
                        break;
                }
            }

            Regex rg = new Regex(@"(.*)\(COM(?<Port>\d+)\)");

            //! Standard Serial over Bluetooth link (COM20)
            if (Properties.ContainsKey(SPDRP.FRIENDLYNAME.ToString()))
            {
                String fname = Properties[SPDRP.FRIENDLYNAME.ToString()];
                if (rg.IsMatch(fname) && !Buster.SerialPorts.ContainsKey(fname))
                {
                    Int32 portnum = Int32.Parse(rg.Match(fname).Groups["Port"].Value);

                    // Debug.Print("{0} = {1}", fname, portnum);

                    Buster.SerialPorts.Add(fname, portnum);
                }
            }
        }

        /// <summary>
        /// Adds an or append property with a value.
        /// </summary>
        ///
        /// <param name="spdrp"> The SPDRP to process. </param>
        /// <param name="value"> The value. </param>
        private void SafeAddOrAppendProperty(SPDRP spdrp, StringBuilder value)
        {
            SafeAddOrAppendProperty(spdrp, value.ToString());
        }

        /// <summary>
        /// Adds an or append property with a value.
        /// </summary>
        ///
        /// <param name="spdrp"> The SPDRP to process. </param>
        /// <param name="spdrp"> The Value. </param>
        private void SafeAddOrAppendProperty(SPDRP spdrp, String value)
        {
            if (!String.IsNullOrEmpty(value.Trim()))
            {
                if (Properties.ContainsKey(spdrp.ToString()))
                {
                    Properties[spdrp.ToString()] += "; " + value.Trim();
                }
                else
                {
                    Properties.Add(spdrp.ToString(), value.Trim());
                }
            }
            else
            {
                //Properties.Add(spdrp.ToString(), String.Empty);
            }
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// The Device Desecription.
        /// </summary>
        ///
        /// <value>
        /// The description.
        /// </value>
        public String Description
        {
            get;
            internal set;
        }

        /// <summary>
        /// The Device Class.
        /// </summary>
        ///
        /// <value>
        /// The device class.
        /// </value>
        public String DeviceClass
        {
            get;
            private set;
        }

        /// <summary>
        /// The Device Status.
        /// </summary>
        ///
        /// <value>
        /// The device status.
        /// </value>
        public String DeviceStatus
        {
            get;
            private set;
        }

        /// <summary>
        /// The Friendly name if any.
        /// </summary>
        ///
        /// <value>
        /// The name of the friendly.
        /// </value>
        public String FriendlyName
        {
            get;
            internal set;
        }

        /// <summary>
        /// Ghosted.
        /// </summary>
        ///
        /// <value>
        /// true if ghosted, false if not.
        /// </value>
        public Boolean Ghosted
        {
            get
            {
                return !String.IsNullOrEmpty(DeviceStatus) && DeviceStatus.Equals("Ghosted");
            }
        }

        /// <summary>
        /// Gets or sets the properties.
        /// </summary>
        ///
        /// <value>
        /// The properties.
        /// </value>
        public Dictionary<String, String> Properties
        {
            get;
            set;
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Removes a Device.
        /// </summary>
        ///
        /// <returns>
        /// true if successfull.
        /// </returns>
        private Boolean RemoveDevice()
        {
            Boolean Result = SetupDi.SetupDiRemoveDevice(Buster.DevInfoSet, ref aDeviceInfoData);

            Int32 lasterror = Marshal.GetLastWin32Error();
            String msg = String.Empty;

            if (Result)
            {
                msg = String.Format("Device '{0}' removed successfully", Description);

                Buster.WriteToEventLog(msg, EventLogEntryType.Information);
            }
            else
            {
                msg = String.Format("Device removal of '{0}' failed", Description);
                msg += "\r\n" + String.Format("; {0}", new Win32Exception(lasterror).Message);

                Buster.WriteToEventLog(msg, EventLogEntryType.Warning);
            }

            return Result;
        }

        /// <summary>
        /// Remove a Ghosted Device if Filtered.
        /// </summary>
        ///
        /// <returns>
        /// true if successfull.
        /// </returns>
        public Boolean RemoveDevice_IF_Ghosted_AND_Marked()
        {
            if (Ghosted
#if CORSAIR
#warning Compiled with code for pesky Corsair Gaming Keyboards that leave numerous non ghosted instances.
                || Description.Equals("Corsair Gaming Keyboard", StringComparison.OrdinalIgnoreCase)
#endif
)
            {
                if (Buster.Devices.Contains(Description.Trim()) && RemoveDevice())
                {
                    return true;
                }

                if (Buster.Classes.Contains(DeviceClass.Trim()) && RemoveDevice())
                {
                    return true;
                }

                foreach (Wildcard w in Buster.Wildcards)
                {
                    if (w.IsMatch(Description.Trim()) && RemoveDevice())
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        #endregion Methods
    }
}