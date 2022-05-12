using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtSecurityDataModule : IBaseModule
    {
        int RTBB_EXTSECURITYDATA_GetData(ref int nSize, byte[] pData);
        string RTBB_EXTSECURITYDATA_GetString();
    }

    public class ExtSecurityDataModule : GlobalVariable, IExtSecurityDataModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtSecurityDataModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtSecurityData";
        }

        ///<summary>
        ///Description: ExtSecurityData get data.
        ///Intput/Output Parameters: nSize ->  IN and OUT parameter, IN: maximal buffer size, OUT: security data size
        ///Output Parameters: pData -> data buffer to read security data.
        ///Return value :  > 0: buffer is not enough.
        ///                = 0: function succeeds.
        ///                < 0: function fails. To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTSECURITYDATA_GetData(ref int nSize, byte[] pData)
        {
            return native_RTBB_EXTSECURITYDATA_GetData(hDev, ref nSize, pData);
        }

        ///<summary>
        ///Description: ExtSecurityData get string.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public string RTBB_EXTSECURITYDATA_GetString()
        {
            return trans_RTBB_EXTSECURITYDATA_GetString(hDev);
        }

        /* ExtIOConfig Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSECURITYDATA_GetData")]
        private static extern int native_RTBB_EXTSECURITYDATA_GetData(IntPtr hDevice, ref int nSize, byte[] pData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSECURITYDATA_GetString")]
        private static extern IntPtr native_RTBB_EXTSECURITYDATA_GetString(IntPtr hDevice);

        private static string trans_RTBB_EXTSECURITYDATA_GetString(IntPtr hDevice)
        {
            return Marshal.PtrToStringAnsi(native_RTBB_EXTSECURITYDATA_GetString(hDevice));
        }
    }
}
