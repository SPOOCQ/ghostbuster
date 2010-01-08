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
#endregion

namespace GhostBuster
{
    using System;
    using System.Collections.Specialized;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using System.Security.Permissions;
    using System.Text;
    using System.Windows.Forms;
    using Microsoft.Win32;

    public class SetupDi
    {

        [Flags]
        internal enum DIGCF : uint
        {
            DIGCF_DEFAULT = 0x00000001,    // only valid with DIGCF_DEVICEINTERFACE
            DIGCF_PRESENT = 0x00000002,
            DIGCF_ALLCLASSES = 0x00000004,
            DIGCF_PROFILE = 0x00000008,
            DIGCF_DEVICEINTERFACE = 0x00000010,
        }

        internal const String SetupApiModuleName = "SetupApi.dll";
        internal const String AdvApi32ModuleName = "advapi32.dll";
        internal const String Kernel32ModuleName = "kernel32.dll";

        internal const UInt32 SPDRP_DEVICEDESC = 0x00000000;

        internal const UInt32 DICS_FLAG_GLOBAL = 0x00000001;

        internal const UInt32 DIREG_DEV = 0x00000001;

        internal const UInt32 KEY_QUERY_VALUE = 0x0001;
        internal const UInt32 KEY_READ = 0x20019;

        internal const Int32 CR_SUCCESS = 0;

        internal const Int32 ERROR_SUCCESS = 0;
        internal const Int32 ERROR_FILE_NOT_FOUND = 2;
        internal const Int32 ERROR_ACCESS_DENIED = 5;
        internal const Int32 ERROR_MORE_DATA = 234;

        internal const Int64 INVALID_HANDLE_VALUE = -1;

        internal static System.Guid NullGuid = System.Guid.Empty;

        //veg: Added (Not sure all these values are correct).
        internal enum DN : int
        {
            DN_ROOT_ENUMERATED = 0x00000001,
            DN_DRIVER_LOADED = 0x00000002,
            DN_ENUM_LOADED = 0x00000004,
            DN_STARTED = 0x00000008,
            DN_MANUAL = 0x00000010,
            DN_NEED_TO_ENUM = 0x00000020,
            DN_NOT_FIRST_TIME = 0x00000040,//?
            DN_HARDWARE_ENUM = 0x00000080,//?
            DN_LIAR = 0x00000100,//?
            DN_HAS_MASK = 0x00000200,
            DN_HAS_PROBLEM = 0x00000400,
            DN_FILTERED = 0x00000800,
            DN_MOVED = 0x00001000,
            DN_DISABLEABLE = 0x00002000,
            DN_REMOVABLE = 0x00004000,//?
            DN_PRIVATE_PROBLEM = 0x00008000,
            DN_MF_PARENT = 0x00010000,
            DN_MF_CHILD = 0x00020000,
            DN_WILL_BE_REMOVED = 0x00040000,
            DN_NOT_FIRST_TIMEE = 0x00080000,
            DN_STOP_FREE_RES = 0x00100000,
            DN_REBAL_CANDIDATE = 0x00200000,//?
            DN_BAD_PARTIAL = 0x00400000,
            DN_NT_ENUMERATOR = 0x00800000,
            DN_NT_DRIVER = 0x01000000,
            DN_NEEDS_LOCKING = 0x02000000,
            DN_ARM_WAKEUP = 0x04000000,
            DN_APM_DRIVER = 0x08000000,
            DN_SILENT_INSTALL = 0x10000000,
            DN_NO_SHOW_IN_DM = 0x20000000,
            DN_BOOT_LOG_PROB = 0x40000000
        }

        //veg: Added
        public enum CM_PROB : uint
        {
            CM_PROB_NOT_CONFIGURED = 1,
            CM_PROB_OUT_OF_MEMORY = 3,
            CM_PROB_INVALID_DATA = 9,
            CM_PROB_FAILED_START = 10,
            CM_PROB_NORMAL_CONFLICT = 12,
            CM_PROB_NEED_RESTART = 14,
            CM_PROB_PARTIAL_LOG_CONF = 16,
            CM_PROB_REINSTALL = 18,
            CM_PROB_REGISTRY = 19,
            CM_PROB_WILL_BE_REMOVED = 21,
            CM_PROB_DISABLED = 22,
            CM_PROB_DEVICE_NOT_THERE = 24,
            CM_PROB_FAILED_INSTALL = 28,
            CM_PROB_HARDWARE_DISABLED = 29,
            CM_PROB_FAILED_ADD = 31,
            CM_PROB_DISABLED_SERVICE = 32,
            CM_PROB_TRANSLATION_FAILED = 33,
            CM_PROB_NO_SOFTCONFIG = 34,
            CM_PROB_BIOS_TABLE = 35,
            CM_PROB_IRQ_TRANSLATION_FAILED = 36,
            CM_PROB_FAILED_DRIVER_ENTRY = 37,
            CM_PROB_DRIVER_FAILED_PRIOR_UNLOAD = 38,
            CM_PROB_DRIVER_FAILED_LOAD = 39,
            CM_PROB_DRIVER_SERVICE_KEY_INVALID = 40,
            CM_PROB_LEGACY_SERVICE_NO_DEVICES = 41,
            CM_PROB_DUPLICATE_DEVICE = 42,
            CM_PROB_FAILED_POST_START = 43,
            CM_PROB_HALTED = 44,
            CM_PROB_PHANTOM = 45,
            CM_PROB_SYSTEM_SHUTDOWN = 46,
            CM_PROB_HELD_FOR_EJECT = 47,
            CM_PROB_DRIVER_BLOCKED = 48,
            CM_PROB_REGISTRY_TOO_LARGE = 49,
            CM_PROB_SETPROPERTIES_FAILED = 50,
            CM_PROB_WAITING_ON_DEPENDENCY = 51,
            CM_PROB_UNSIGNED_DRIVER = 52
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct SP_DEVINFO_DATA
        {
            public Int32 cbSize;
            public Guid ClassGuid;
            public Int32 DevInst;
            public UIntPtr Reserved;
        };

        [DllImport(SetupApiModuleName)]
        internal static extern Int32 SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport(SetupApiModuleName)]
        internal static extern bool SetupDiEnumDeviceInfo(IntPtr DeviceInfoSet, Int32 MemberIndex, ref  SP_DEVINFO_DATA DeviceInterfaceData);

        [DllImport(SetupApiModuleName, CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern bool SetupDiGetDeviceRegistryProperty(IntPtr deviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData,
            uint property, out UInt32 propertyRegDataType, StringBuilder propertyBuffer, uint propertyBufferSize, out UInt32 requiredSize);

        [DllImport(SetupApiModuleName, SetLastError = true)]
        internal static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, UInt32 iEnumerator, IntPtr hParent, UInt32 nFlags);

        [DllImport(SetupApiModuleName, CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern IntPtr SetupDiOpenDevRegKey(IntPtr hDeviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, uint scope,
            uint hwProfile, uint parameterRegistryValueKind, uint samDesired);

        [DllImport(SetupApiModuleName, CharSet = CharSet.Auto, SetLastError = true)]
        public static extern Int32 SetupDiGetClassDescription(ref Guid ClassGuid,
            StringBuilder classDescription, Int32 ClassDescriptionSize, ref Int32 RequiredSize);

        [DllImport(SetupApiModuleName)]
        internal static extern Int32 SetupDiGetDeviceInstanceId(IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            StringBuilder DeviceInstanceId, Int32 DeviceInstanceIdSize, ref Int32 RequiredSize);

        [DllImport(SetupApiModuleName, SetLastError = true)]
        internal static extern int CM_Get_DevNode_Status(ref int pulStatus, ref int pulProblemNumber, int dnDevInst, int ulFlags);

        [DllImport(SetupApiModuleName, SetLastError = true)]
        internal static extern bool SetupDiRemoveDevice(IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData);

        //[DllImport(AdvApi32ModuleName, EntryPoint = "RegEnumKeyEx")]
        //extern private static int RegEnumKeyEx(IntPtr hKey,
        //    uint index,
        //    StringBuilder lpName,
        //    ref uint lpcbName,
        //    IntPtr reserved,
        //    IntPtr lpClass,
        //    IntPtr lpcbClass,
        //    IntPtr lpftLastWriteTime);

        //[DllImport(AdvApi32ModuleName, CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueExW", SetLastError = true)]
        //public static extern int RegQueryValueEx(IntPtr hKey, string lpValueName, int lpReserved, out uint lpType,
        //    StringBuilder lpData, ref uint lpcbData);

        [DllImport(AdvApi32ModuleName, CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int RegCloseKey(IntPtr hKey);

        //[DllImport(Kernel32ModuleName)]
        //public static extern Int32 GetLastError();

        /// <summary>
        /// Structure to store the state of a device.
        /// </summary>
        public struct DeviceInfo
        {
            public string deviceclass;
            public string name;
            public string friendlyname;
            public string description;

            public bool disabled;
            public bool service;
            public bool ghosted;
            public int status;
            public CM_PROB problem;
        }

        /// <summary>
        /// Stores all known Services and their DisplayNames.
        /// </summary>
        internal static StringCollection serviceslist = new StringCollection();

        /// <summary>
        /// Get Name of Device and try to retrieve its FriendlyName.
        /// 
        /// NOTE: The FriendlyName fails.
        /// </summary>
        /// <param name="deviceinfo"></param>
        /// <param name="pdevinfoset"></param>
        /// <param name="deviceinfodata"></param>
        /// <returns></returns>
        internal static Boolean GetDeviceName(ref DeviceInfo deviceinfo, IntPtr pdevinfoset, SP_DEVINFO_DATA deviceinfodata)
        {
            deviceinfo.name = deviceinfo.description;
            deviceinfo.friendlyname = deviceinfo.name;

            IntPtr hDeviceRegistryKey = SetupDiOpenDevRegKey(pdevinfoset, ref deviceinfodata,
                DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_READ);
            if ((long)hDeviceRegistryKey == INVALID_HANDLE_VALUE)
            {
                //throw new Exception("Failed to open a registry key for device-specific configuration information");
                return false;
            }

            StringBuilder deviceNameBuf = new StringBuilder(256);
            try
            {
                //uint lpRegKeyType;
                //uint length = (uint)deviceNameBuf.Capacity;
                //int result = RegQueryValueEx(hDeviceRegistryKey, "DisplayName", 0, out lpRegKeyType, deviceNameBuf, ref length);

                ////int result = RegEnumKeyEx(hDeviceRegistryKey, 0, deviceNameBuf, ref length, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                //if (result == ERROR_SUCCESS)
                //{
                //    //throw new Exception("Can not read registry value PortName for device " + deviceInfoData.ClassGuid);
                //    dev.friendlyname = deviceNameBuf.ToString();
                //}
            }

            finally
            {
                RegCloseKey(hDeviceRegistryKey);
            }

            return String.IsNullOrEmpty(deviceinfo.friendlyname);
        }

        /// <summary>
        /// Retrieve the devices decription.
        /// </summary>
        /// <param name="deviceinfo"></param>
        /// <param name="hdeviceinfoset"></param>
        /// <param name="deviceinfodata"></param>
        /// <returns></returns>
        internal static Boolean GetDeviceDescription(ref DeviceInfo deviceinfo, IntPtr hdeviceinfoset, SP_DEVINFO_DATA deviceinfodata)
        {
            StringBuilder descriptionBuf = new StringBuilder(256);
            uint propRegDataType;
            uint length = (uint)descriptionBuf.Capacity;

            if (SetupDiGetDeviceRegistryProperty(hdeviceinfoset, ref deviceinfodata, SPDRP_DEVICEDESC,
                out propRegDataType, descriptionBuf, length, out length))
            {
                deviceinfo.description = descriptionBuf.ToString();
                return true;
            }
            else
            {
                deviceinfo.description = deviceinfodata.ClassGuid.ToString();
                return false;
            }
        }

        /// <summary>
        /// Retrieve and decode the device status.
        /// </summary>
        /// <param name="deviceinfo"></param>
        /// <param name="deviceinfoset"></param>
        /// <param name="deviceinfodata"></param>
        /// <returns></returns>
        internal static Boolean GetDeviceStatus(ref DeviceInfo deviceinfo, IntPtr deviceinfoset, ref SetupDi.SP_DEVINFO_DATA deviceinfodata)
        {
            StringBuilder descriptionBuf = new StringBuilder(256);
            int length = descriptionBuf.Capacity;

            //Assume All is ok
            deviceinfo.status = 0;
            deviceinfo.problem = 0;
            deviceinfo.service = false;
            deviceinfo.disabled = false;
            deviceinfo.ghosted = false;

            if (SetupDi.SetupDiGetDeviceInstanceId(deviceinfoset, ref deviceinfodata, descriptionBuf, length, ref length) != 0)
            {
                int aDevInst = deviceinfodata.DevInst;
                int aStatus = 0;
                int aProblem = 0;

                if (CM_Get_DevNode_Status(ref aStatus, ref aProblem, aDevInst, (int)0) == CR_SUCCESS)
                {
                    deviceinfo.status = aStatus;

                    if ((aStatus & (int)DN.DN_HAS_PROBLEM) != 0)
                    {
                        switch (aProblem)
                        {
                            case (int)CM_PROB.CM_PROB_DISABLED:
                            case (int)CM_PROB.CM_PROB_DISABLED_SERVICE:
                                deviceinfo.disabled = true;
                                break;
                            default:
                                break;
                        }
                    }
                    else if ((aStatus & ((int)DN.DN_DRIVER_LOADED | (int)DN.DN_STARTED)) != 0)
                    {
                        //return String.Format("0x{0:x8} - {1}", aStatus, "OK");
                    }
                    else if ((aStatus & ((int)DN.DN_DISABLEABLE)) != 0)
                    {
                        //
                    }
                    else
                    {
                        //
                    }
                }
                else
                {
                    //using (RegistryKey services = Registry.LocalMachine.OpenSubKey(@"System\CurrentControlSet\Services"))
                    //{
                    //    if (services.OpenSubKey(deviceinfo.name) != null)
                    //        return String.Format("{0}", "Service");

                    //    //Also Scan All Services for DisplayNames....
                    //}
                    if (serviceslist.Contains(deviceinfo.name.ToLower()))
                    {
                        deviceinfo.service = true;
                    }
                    else
                    {
                        deviceinfo.ghosted = true;
                    }
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Retrieve the Description of a DeviceClass.
        /// </summary>
        /// <param name="deviceinfo"></param>
        /// <param name="guid"></param>
        /// <returns></returns>
        internal static Boolean GetClassDescriptionFromGuid(ref DeviceInfo deviceinfo, Guid guid)
        {
            deviceinfo.deviceclass = String.Empty;

            StringBuilder strClassDesc = new StringBuilder(0);
            Int32 iRequiredSize = 0;
            Int32 iSize = 0;
            Int32 iRet = SetupDi.SetupDiGetClassDescription(ref guid, strClassDesc, iSize, ref iRequiredSize);
            strClassDesc = new StringBuilder(iRequiredSize);
            iSize = iRequiredSize;
            iRet = SetupDi.SetupDiGetClassDescription(ref guid, strClassDesc, iSize, ref iRequiredSize);
            if (iRet == 1)
            {
                deviceinfo.deviceclass = strClassDesc.ToString();
            }

            return String.IsNullOrEmpty(deviceinfo.deviceclass);
        }

        /// <summary>
        /// Get Device Instance Id.
        /// </summary>
        /// <param name="DeviceInfoSet"></param>
        /// <param name="DeviceInfoData"></param>
        /// <returns></returns>
        internal String GetDeviceInstanceId(IntPtr DeviceInfoSet, SetupDi.SP_DEVINFO_DATA DeviceInfoData)
        {
            StringBuilder strId = new StringBuilder(0);
            Int32 iRequiredSize = 0;
            Int32 iSize = 0;

            /*Int32 iRet = */
            SetupDi.SetupDiGetDeviceInstanceId(DeviceInfoSet, ref DeviceInfoData, strId, iSize, ref iRequiredSize);
            strId = new StringBuilder(iRequiredSize);
            iSize = iRequiredSize;

            if (SetupDi.SetupDiGetDeviceInstanceId(DeviceInfoSet, ref DeviceInfoData, strId, iSize, ref iRequiredSize) == 1)
            {
                return strId.ToString();
            }
            return String.Empty;
        }

        /// <summary>
        /// Enumerate Names of Services and their DisplayNames.
        /// </summary>
        //[RegistryPermissionAttribute(SecurityAction.Demand, Unrestricted = true)]
        internal static void EnumServices()
        {
            //We need AllAccess for Enumeration.            
            RegistryPermission f = new RegistryPermission(RegistryPermissionAccess.AllAccess, @"HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\");
            using (RegistryKey services = Registry.LocalMachine.OpenSubKey(@"System\CurrentControlSet\Services", false))
            {
                foreach (String name in services.GetSubKeyNames())
                {
                    try
                    {
                        serviceslist.Add(name.ToLower());

                        f.AddPathList(RegistryPermissionAccess.Read, @"System\CurrentControlSet\Services\" + name);
                        using (RegistryKey service = services.OpenSubKey(name, false))
                        {
                            Object displayname = service.GetValue("DisplayName");
                            if (displayname != null)
                            {
                                //Debug.WriteLine(displayname.ToString());
                                serviceslist.Add(displayname.ToString().ToLower());
                            }
                        }
                    }
                    catch
                    {
                        //SBCore fails on WHS And propably W2K3 Server too.
                        //MessageBox.Show(name);
                    }
                }
            }
        }
    }
}