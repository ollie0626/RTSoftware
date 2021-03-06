using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

class Dbt
{
    #region Dbt Class - Constants

    public const ushort WM_DEVICECHANGE = 0x0219;
    public const ushort DBT_DEVICEARRIVAL = 0x8000;
    public const ushort DBT_DEVICEREMOVECOMPLETE = 0x8004;
    public const ushort DBT_DEVTYP_DEVICEINTERFACE = 0x0005;
    public const int DEVICE_NOTIFY_WINDOW_HANDLE = 0x0000;
    public const string WINUSB_DEMO_GUID_STRING = "{427A71EC-44c2-82DE-87D8D8FF6C1E}";

    #endregion

    #region Dbt Class - Device Change Structures

    [StructLayout(LayoutKind.Sequential)]
    public class DEV_BROADCAST_DEVICEINTERFACE
    {
        public int dbcc_size;
        public int dbcc_devicetype;
        public int dbcc_reserved;
        public Guid dbcc_classguid;
        public char dbcc_name;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public class DEV_BROADCAST_DEVICEINTERFACE_1
    {
        public int dbcc_size;
        public int dbcc_devicetype;
        public int dbcc_reserved;
        public Guid dbcc_classguid;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 255)]
        public char[] dbcc_name;
    }

    [StructLayout(LayoutKind.Sequential)]
    public class DEV_BROADCAST_HDR
    {
        public int dbch_size;
        public int dbch_devicetype;
        public int dbch_reserved;
    }

    #endregion

    #region DLL Imports

    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr RegisterDeviceNotification(IntPtr hRecipient, IntPtr NotificationFilter, uint Flags);

    [DllImport("user32.dll")]
    public static extern uint UnregisterDeviceNotification(IntPtr Handle);

    [DllImport("hid.dll", SetLastError = true)]
    public static extern void HidD_GetHidGuid(ref System.Guid HidGoid);
    #endregion
}
