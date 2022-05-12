using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IGPIOExtModule : IBaseModule
    {
        int RTBB_GPIOExtSetPinMode(int nPin, UInt32 nMode);
        int RTBB_GPIOExtGetCurrentPinMode(int nPin, ref UInt32 nMode);
        int RTBB_GPIOExtGetPinModeCount(ref UInt32 pCount);
        string RTBB_GPIOExtGetPinModeName(uint nMode);
        int RTBB_GPIOExtGetPinODMode(int nPin, ref bool pOpenDrain);
        int RTBB_GPIOExtSetPinODMode(int nPin, bool pOpenDrain);
        int RTBB_GPIOExtGetPinSelCount(ref UInt32 pCount);
        int RTBB_GPIOExtSetPinSel(int nPin, UInt32 nSel);
        string RTBB_GPIOExtGetPinSelName(UInt32 nPinNumber, UInt32 nSel);
    }

    public class GPIOExtModule : GlobalVariable, IGPIOExtModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public GPIOExtModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "GPIOExt";
        }

        ///<summary>
        ///Description: GPIOExt Set Pin Mode.
        ///Input Parameters: nPin -> pin number
        ///Input Parameters: nMode -> just choose the bit that the gpio number you want to set.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtSetPinMode(int nPin, UInt32 nMode)
        {
            return native_RTBB_GPIOExtSetPinMode(hDev, nPin, nMode);
        }

        ///<summary>
        ///Description: GPIOExt Get Current Pin Mode.
        ///Input Parameters: nPin -> pin number
        ///Input Parameters: nMode -> Return the current mode.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtGetCurrentPinMode(int nPin, ref UInt32 nMode)
        {
            return native_RTBB_GPIOExtGetCurrentPinMode(hDev, nPin, ref nMode);
        }

        ///<summary>
        ///Description: GPIOExt Get Pin Mode Count.
        ///Input Parameters: pCount -> Return the pinCount.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtGetPinModeCount(ref UInt32 pCount)
        {
            return native_RTBB_GPIOExtGetPinModeCount(hDev, ref pCount);
        }

        ///<summary>
        ///Description: GPIOExt Get Pin Mode Name.
        ///Input Parameters: nMode -> Return the current mode.
        ///Return string if success, other non-zero value.
        ///Else, return null string.
        ///</summary>
        public string RTBB_GPIOExtGetPinModeName(uint nMode)
        {
            return trans_RTBB_GPIOExtGetPinModeName(hDev, nMode);
        }

        ///<summary>
        ///Description: GPIOExt Get Pin OD Mode.
        ///Input Parameters: nPin -> pin number
        ///Input Parameters: pOpenDrain -> Return the Open Drain status for the specified pin.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtGetPinODMode(int nPin, ref bool pOpenDrain)
        {
            return native_RTBB_GPIOExtGetPinODMode(hDev, nPin, ref pOpenDrain);
        }

        ///<summary>
        ///Description: GPIOExt Set Pin OD Mode.
        ///Input Parameters: nPin -> pin number
        ///Input Parameters: pOpenDrain -> If true, set Open Drain, otherwise not.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtSetPinODMode(int nPin, bool pOpenDrain)
        {
            return native_RTBB_GPIOExtSetPinODMode(hDev, nPin, pOpenDrain);
        }

        ///<summary>
        ///Description: GPIOExt Get Pin Selection Count.
        ///Input Parameters: pCount -> Indicate the function mode count.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtGetPinSelCount(ref UInt32 pCount)
        {
            return native_RTBB_GPIOExtGetPinSelCount(hDev, ref pCount);
        }

        ///<summary>
        ///Description: GPIOExt Set Pin Selection.
        ///Input Parameters: nPin -> Index of GPIO pin, Start number is 0.
        ///Input Parameters: nSel -> The new function mode setting of GPIO pin.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_GPIOExtSetPinSel(int nPin, UInt32 nSel)
        {
            return native_RTBB_GPIOExtSetPinSel(hDev, nPin, nSel);
        }

        ///<summary>
        ///Description: GPIOExt Get Pin Selection.
        ///Input Parameters: nPin -> Index of GPIO pin, Start number is 0.
        ///Input Parameters: nSel -> The function mode index to be retrieved, start number is 0.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public string RTBB_GPIOExtGetPinSelName(UInt32 nPinNumber, UInt32 nSel)
        {
            return trans_RTBB_GPIOExtGetPinSelName(hDev, nPinNumber, nSel);
        }

        /* GPIO Ext control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtSetPinMode")]
        private static extern int native_RTBB_GPIOExtSetPinMode(IntPtr hDevice, int nPin, UInt32 nMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetCurrentPinMode")]
        public static extern int native_RTBB_GPIOExtGetCurrentPinMode(IntPtr hDevice, int nPin, ref UInt32 nMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetPinModeCount")]
        public static extern int native_RTBB_GPIOExtGetPinModeCount(IntPtr hDevice, ref UInt32 pCount);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetPinModeName")]
        private static extern IntPtr native_RTBB_GPIOExtGetPinModeName(IntPtr hDevice, uint nMode);

        private static string trans_RTBB_GPIOExtGetPinModeName(IntPtr hDevice, uint nMode)
        {
            return Marshal.PtrToStringAnsi(native_RTBB_GPIOExtGetPinModeName(hDevice, nMode));
        }

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetPinODMode")]
        private static extern int native_RTBB_GPIOExtGetPinODMode(IntPtr hDevice, int nPin, ref bool pOpenDrain);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtSetPinODMode")]
        private static extern int native_RTBB_GPIOExtSetPinODMode(IntPtr hDevice, int nPin, bool pOpenDrain);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetPinSelCount")]
        private static extern int native_RTBB_GPIOExtGetPinSelCount(IntPtr hDevice, ref UInt32 pCount);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtSetPinSel")]
        private static extern int native_RTBB_GPIOExtSetPinSel(IntPtr hDevice, int nPin, UInt32 nSel);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GPIOExtGetPinSelName")]
        private static extern IntPtr native_RTBB_GPIOExtGetPinSelName(IntPtr hDevice, UInt32 nPinNumber, UInt32 nSel);

        private static string trans_RTBB_GPIOExtGetPinSelName(IntPtr hDevice, UInt32 nPinNumber, UInt32 nSel)
        {
            return Marshal.PtrToStringAnsi(native_RTBB_GPIOExtGetPinSelName(hDevice, nPinNumber, nSel));
        }
    }
}
