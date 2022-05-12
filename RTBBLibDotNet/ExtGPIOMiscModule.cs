using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtGPIOMiscMdoule : IBaseModule
    {
        int RTBB_EXTGPIOMISC_GetUniversalGPIOCount(ref uint pUniversaPinCount);
        int RTBB_EXTGPIOMISC_GetUniversalGPIOMapping(uint nUniversalPinNumber, ref uint pGPIOPinNumber);
        int RTBB_EXTGPIOMISC_GetSPICSPinCount(ref uint pSPICSPinCount);
        int RTBB_EXTGPIOMISC_GetSPIPinMapping(uint nSPICSPinNumber, ref uint pGPIOPinNumber);
        int RTBB_EXTGPIOMISC_GetGPIOPinMode(uint nGPIOPinNumber, ref uint pPinMode);
        int RTBB_EXTGPIOMISC_SetGPIOPinMode(uint nGPIOPinNumber, uint pPinMode);
        int RTBB_EXTGPIOMISC_GetGPIOPinSel(uint nGPIOPinNumber, ref uint pPinSel);
        int RTBB_EXTGPIOMISC_SetGPIOPinSel(uint nGPIOPinNumber, uint pPinSel);
    }

    public class ExtGPIOMiscMdoule : GlobalVariable, IExtGPIOMiscMdoule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtGPIOMiscMdoule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtGPIOMisc";
        }

        ///<summary>
        ///Description: Get the universal GPIO count.
        ///Output Parameters: pUniversaPinCount -> A pointer to unsigned integer which indicates the universal GPIO count of bridge board.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetUniversalGPIOCount(ref uint pUniversaPinCount)
        {
            return native_RTBB_EXTGPIOMISC_GetUniversalGPIOCount(hDev, ref pUniversaPinCount);
        }

        ///<summary>
        ///Description: Get the universal GPIO mapping.
        ///Input Parameters: pUniversaPinCount -> Index of universal pin.
        ///Output Parameters: pGPIOPinNumber -> A pointer to unsigned integer which indicates the mapping GPIO pin of nUniversalPinNumber.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetUniversalGPIOMapping(uint nUniversalPinNumber, ref uint pGPIOPinNumber)
        {
            return native_RTBB_EXTGPIOMISC_GetUniversalGPIOMapping(hDev, nUniversalPinNumber, ref pGPIOPinNumber);
        }

        ///<summary>
        ///Description: Get SPI CS Pin Count.
        ///Output Parameters: pSPICSPinCount -> A pointer to unsigned integer which indicates the SPI chip select pin count of bridge board.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetSPICSPinCount(ref uint pSPICSPinCount)
        {
            return native_RTBB_EXTGPIOMISC_GetSPICSPinCount(hDev, ref pSPICSPinCount);
        }

        ///<summary>
        ///Description: Get SPI Pin mapping.
        ///Input Parameters: nSPICSPinNumber -> Index of chip select pin.
        ///Output Parameters: pGPIOPinNumber -> A pointer to unsigned integer which indicates the mapping GPIO pin of nSPICSPinNumber.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetSPIPinMapping(uint nSPICSPinNumber, ref uint pGPIOPinNumber)
        {
            return native_RTBB_EXTGPIOMISC_GetSPIPinMapping(hDev, nSPICSPinNumber, ref pGPIOPinNumber);
        }

        ///<summary>
        ///Description: Get GPIO Pin mode.
        ///Input Parameters: nGPIOPinNumber -> Index of GPIO pin
        ///Output Parameters: pPinMode -> A pointer to unsigned integer which indicates the pin mode of nGPIOPinNumber.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetGPIOPinMode(uint nGPIOPinNumber, ref uint pPinMode)
        {
            return native_RTBB_EXTGPIOMISC_GetGPIOPinMode(hDev, nGPIOPinNumber, ref pPinMode);
        }

        ///<summary>
        ///Description: Set GPIO Pin mode.
        ///Input Parameters: nGPIOPinNumber -> Index of GPIO pin
        ///Input Parameters: pPinMode -> The new pin mode setting of GPIO pin.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_SetGPIOPinMode(uint nGPIOPinNumber, uint pPinMode)
        {
            return native_RTBB_EXTGPIOMISC_SetGPIOPinMode(hDev, nGPIOPinNumber, pPinMode);
        }

        ///<summary>
        ///Description: Get GPIO Pin selection.
        ///Input Parameters: nGPIOPinNumber -> Index of GPIO pin
        ///Output Parameters: pPinSel -> A pointer to unsigned integer which indicates the function mode of nGPIOPinNumber.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_GetGPIOPinSel(uint nGPIOPinNumber, ref uint pPinSel)
        {
            return native_RTBB_EXTGPIOMISC_GetGPIOPinSel(hDev, nGPIOPinNumber, ref pPinSel);
        }

        //<summary>
        ///Description: Set GPIO Pin selection.
        ///Input Parameters: nGPIOPinNumber -> Index of GPIO pin
        ///Input Parameters: pPinSel -> The new function mode setting of GPIO pin.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGPIOMISC_SetGPIOPinSel(uint nGPIOPinNumber, uint pPinSel)
        {
            return native_RTBB_EXTGPIOMISC_SetGPIOPinSel(hDev, nGPIOPinNumber, pPinSel);
        }

        /* extGPIOMisc control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetUniversalGPIOCount")]
        private static extern int native_RTBB_EXTGPIOMISC_GetUniversalGPIOCount(IntPtr hDevice, ref uint pUniversaPinCount);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetUniversalGPIOMapping")]
        private static extern int native_RTBB_EXTGPIOMISC_GetUniversalGPIOMapping(IntPtr hDevice, uint nUniversalPinNumber, ref uint pGPIOPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetSPICSPinCount")]
        private static extern int native_RTBB_EXTGPIOMISC_GetSPICSPinCount(IntPtr hDevice, ref uint pSPICSPinCount);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetSPIPinMapping")]
        private static extern int native_RTBB_EXTGPIOMISC_GetSPIPinMapping(IntPtr hDevice, uint nSPICSPinNumber, ref uint pGPIOPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetGPIOPinMode")]
        private static extern int native_RTBB_EXTGPIOMISC_GetGPIOPinMode(IntPtr hDevice, uint nGPIOPinNumber, ref uint pPinMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_SetGPIOPinMode")]
        private static extern int native_RTBB_EXTGPIOMISC_SetGPIOPinMode(IntPtr hDevice, uint nGPIOPinNumber, uint pPinMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_GetGPIOPinSel")]
        private static extern int native_RTBB_EXTGPIOMISC_GetGPIOPinSel(IntPtr hDevice, uint nGPIOPinNumber, ref uint pPinSel);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGPIOMISC_SetGPIOPinSel")]
        private static extern int native_RTBB_EXTGPIOMISC_SetGPIOPinSel(IntPtr hDevice, uint nGPIOPinNumber, uint pPinSel);
    }
}
