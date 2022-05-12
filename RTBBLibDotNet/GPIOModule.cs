using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IGPIOModule : IBaseModule
    {
        int RTBB_GPIOSetIODirection(int nPort, UInt32 nMask, UInt32 nValue);
        int RTBB_GPIOGetIODirection(int nPort, ref UInt32 pValue);
        int RTBB_GPIOWrite(int nPort, UInt32 nMask, UInt32 nValue);
        int RTBB_GPIORead(int nPort, ref UInt32 pValue);
        int RTBB_GPIOSingleSetIODirection(int nPinNumber, bool bOutput);
        int RTBB_GPIOSingleWrite(int nPinNumber, bool bValue);
        int RTBB_GPIOSingleRead(int nPinNumber, ref bool bValue);
        int RTBB_GPIOPinNumber2PortNumber(int nPinNumber, ref int pPortNumber);
        int RTBB_GPIOGetPinCount(ref int pCount);
        string RTBB_GPIOGetPinNameAndMode(int nPinNumber, ref int pMode);
    }

    public class GPIOModule : GlobalVariable, IGPIOModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public GPIOModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "GPIO";
        }

        ///<summary>
        ///Description: GPIO Set IO Direction.
        ///Input Parameters: nPort -> gpio port number
        ///Input Parameters: nMask -> just choose the bit that the gpio number you want to set.
        ///Input Parameters: nValue -> if a bit is set to 1, means output mode, otherwise, means input mode.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOSetIODirection(int nPort, UInt32 nMask, UInt32 nValue)
        {
            return native_RTBB_GPIOSetIODirection(hDev, nPort, nMask, nValue);
        }

        ///<summary>
        ///Description: GPIO Get IO Direction.
        ///Input Parameters: nPort -> gpio port number
        ///Input Parameters: pValue -> return value. if a bit is set to 1, means output mode, otherwise, means input mode.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOGetIODirection(int nPort, ref UInt32 pValue)
        {
            return native_RTBB_GPIOGetIODirection(hDev, nPort, ref pValue);
        }

        ///<summary>
        ///Description: GPIO Write.
        ///Input Parameters: nPort -> gpio port number
        ///Input Parameters: nMask -> just choose the bit that the gpio number you want to set.
        ///Input Parameters: nValue -> if a bit is set to 1, means output high, otherwise, means output low.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOWrite(int nPort, UInt32 nMask, UInt32 nValue)
        {
            return native_RTBB_GPIOWrite(hDev, nPort, nMask, nValue);
        }

        ///<summary>
        ///Description: GPIO Read.
        ///Input Parameters: nPort -> gpio port number
        ///Input Parameters: pValue -> return value. if a bit is set to 1, means output/input high, otherwise, means output/input low.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIORead(int nPort, ref UInt32 pValue)
        {
            return native_RTBB_GPIORead(hDev, nPort, ref pValue);
        }

        ///<summary>
        ///Description: Single GPIO Set IO Direction.
        ///Input Parameters: hDevice -> Bridget board device handle.
        ///Input Parameters: nPinNumber -> gpio pin number
        ///Input Parameters: bOutput -> if true, output mode, otherwise, input mode.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOSingleSetIODirection(int nPinNumber, bool bOutput)
        {
            return native_RTBB_GPIOSingleSetIODirection(hDev, nPinNumber, bOutput);
        }

        ///<summary>
        ///Description: Single GPIO Write.
        ///Input Parameters: nPinNumber -> gpio pin number
        ///Input Parameters: bValue -> if true, output high, otherwise, output low.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOSingleWrite(int nPinNumber, bool bValue)
        {
            return native_RTBB_GPIOSingleWrite(hDev, nPinNumber, bValue);
        }

        ///<summary>
        ///Description: Single GPIO Read.
        ///Input Parameters: nPinNumber -> gpio pin number
        ///Input Parameters: bValue -> if true, output high, otherwise, output low.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOSingleRead(int nPinNumber, ref bool bValue)
        {
            return native_RTBB_GPIOSingleRead(hDev, nPinNumber, ref bValue);
        }

        ///<summary>
        ///Description: GIPO Pin Number 2 Port Number.
        ///Input Parameters: nPinNumber -> gpio pin number
        ///Input Parameters: pPortNumber -> return the specified GPIO pin port number.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOPinNumber2PortNumber(int nPinNumber, ref int pPortNumber)
        {
            return native_RTBB_GPIOPinNumber2PortNumber(hDev, nPinNumber, ref nPinNumber);
        }

        ///<summary>
        ///Description: GIPO Get Pin Count.
        ///Input Parameters: pCount -> return total gpio pin count.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOGetPinCount(ref int pCount)
        {
            return native_RTBB_GPIOGetPinCount(hDev, ref pCount);
        }

        ///<summary>
        ///Description: GIPO Get Pin Name and Mode.
        ///Input Parameters: nPinNumber -> gpio pin number.
        ///Input Parameters: pMode -> return the specified mode for the specified gpio pin.
        ///Output value: If success, return a lpcstr string pointer.
        ///              else return NULL;
        ///</summary>
        public string RTBB_GPIOGetPinNameAndMode(int nPinNumber, ref int pMode)
        {
            return native_RTBB_GPIOGetPinNameAndMode(hDev, nPinNumber, ref pMode);
        }

        /* GPIO Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOSetIODirection")]
        private static extern int native_RTBB_GPIOSetIODirection(IntPtr hDevice, int nPort, UInt32 nMask, UInt32 nValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOGetIODirection")]
        private static extern int native_RTBB_GPIOGetIODirection(IntPtr hDevice, int nPort, ref UInt32 pValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOWrite")]
        private static extern int native_RTBB_GPIOWrite(IntPtr hDevice, int nPort, UInt32 nMask, UInt32 nValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIORead")]
        private static extern int native_RTBB_GPIORead(IntPtr hDevice, int nPort, ref UInt32 pValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOSingleSetIODirection")]
        private static extern int native_RTBB_GPIOSingleSetIODirection(IntPtr hDevice, int nPinNumber, bool bOutput);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOSingleWrite")]
        private static extern int native_RTBB_GPIOSingleWrite(IntPtr hDevice, int nPinNumber, bool bValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOSingleRead")]
        private static extern int native_RTBB_GPIOSingleRead(IntPtr hDevice, int nPinNumber, ref bool bValue);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOPinNumber2PortNumber")]
        private static extern int native_RTBB_GPIOPinNumber2PortNumber(IntPtr hDevice, int nPinNumber, ref int pPortNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOGetPinCount")]
        private static extern int native_RTBB_GPIOGetPinCount(IntPtr hDevice, ref int pCount);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOGetPinNameAndMode")]
        private static extern string native_RTBB_GPIOGetPinNameAndMode(IntPtr hDevice, int nPinNumber, ref int pMode);
    }
}
