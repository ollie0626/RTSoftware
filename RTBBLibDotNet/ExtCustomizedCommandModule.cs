using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtCustomizedCommandModule : IBaseModule
    {
        int RTBB_EXTCFW_GetCFWVersion();
        string RTBB_EXTCFW_GetCFWVendor();
        int RTBB_EXTCFW_Transact(ref int pCmdIn, ref int pDataInCount, byte[] pDataIn, ref int pCmdOut, ref int pDataOutCount, byte[] pDataOut);
    }

    public class ExtCustomizedCommandModule : GlobalVariable, IExtCustomizedCommandModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtCustomizedCommandModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtCustomizedCommand";
        }

        ///<summary>
        ///Description: get the custom firmware version.
        ///If the function succeeds, the return value is Version of CFW.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTCFW_GetCFWVersion()
        {
            return native_RTBB_EXTCFW_GetCFWVersion(hDev);
        }

        ///<summary>
        ///Description: get the custom firmware vendor
        ///If the function succeeds, the return value is Vendor Information of CFW.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public string RTBB_EXTCFW_GetCFWVendor()
        {
            return trans_RTBB_EXTCFW_GetCFWVendor(hDev);
        }

        ///<summary>
        ///Description: custom firmware transact.
        ///Input Parameters: pCmdIn -> User command data which from host to bridge board.
        ///Input Parameters: pDataInCount -> the size of user transact data in pDataIn which from host to bridge board.
        ///Input Parameters: pDataIn -> User data array which from host to bridge board.
        ///Output Parameters: pCmdOut -> Command data which return from bridge board to host.
        ///Output Parameters: pDataOutCount -> the size of transact data in pDataOut which return from bridge board to host.
        ///Output Parameters: pDataOut -> The data array which return from bridge board to host.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTCFW_Transact(ref int pCmdIn, ref int pDataInCount, byte[] pDataIn, ref int pCmdOut, ref int pDataOutCount, byte[] pDataOut)
        {
            return native_RTBB_EXTCFW_Transact(hDev, ref pCmdIn, ref pDataInCount, pDataIn, ref pCmdOut, ref pDataOutCount, pDataOut);
        }

        /* ExtCustomizedCommand Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTCFW_GetCFWVersion")]
        private static extern int native_RTBB_EXTCFW_GetCFWVersion(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTCFW_GetCFWVendor")]
        private static extern IntPtr native_RTBB_EXTCFW_GetCFWVendor(IntPtr hDevice);

        
        private string trans_RTBB_EXTCFW_GetCFWVendor(IntPtr hDevice)
        {
            return Marshal.PtrToStringAnsi(native_RTBB_EXTCFW_GetCFWVendor(hDevice));
        }

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTCFW_Transact")]
        private static extern int native_RTBB_EXTCFW_Transact(IntPtr hDevice, ref int pCmdIn, ref int pDataInCount, byte[] pDataIn, ref int pCmdOut, ref int pDataOutCount, byte[] pDataOut);
    }
}
