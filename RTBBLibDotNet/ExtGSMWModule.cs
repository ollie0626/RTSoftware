using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtGSMWModule : IBaseModule
    {
        int RTBB_EXTGSMW_SendData(uint nLength, uint nMask, UInt16[] pBuffer, bool bCheckAck);
        int RTBB_EXTGSMW_GetDataMaxCount();
        int RTBB_EXTGSMW_SetBaseClk(uint nClkNs, byte nClkMode);
        int RTBB_EXTGSMW_GetBaseClk();
        int RTBB_EXTGSMW_GetPinCount();
        int RTBB_EXTGSMW_GetAllPinsMask();
    }

    public class ExtGSMWModule : GlobalVariable, IExtGSMWModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtGSMWModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtGSMW";
        }

        ///<summary>
        ///Description: GSMW send data.
        ///Input Parameters: nLength -> the length of pBuffer data to be output on GSOW wires.
        ///Input Parameters: nMask -> if pattern will be output on specified pins.
        ///                        -> 0: disable output.
        ///                        -> 1: enable output.
        ///                        -> this value has no effect in GSOW, it is valid in GSMW module.
        ///Input Parameters: pBuffer -> A pointer to the buffer that containing the data to be written to the one-wire.
        ///Input Parameters: bCheckAck -> the default value is FALSE.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_SendData(uint nLength, uint nMask, UInt16[] pBuffer, bool bCheckAck)
        {
            return native_RTBB_EXTGSMW_SendData(hDev, nLength, nMask, pBuffer, bCheckAck);
        }

        ///<summary>
        ///Description: GSMW get data max count.
        ///If the function succeeds, the return value is the quantities of max available data length.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_GetDataMaxCount()
        {
            return native_RTBB_EXTGSMW_GetDataMaxCount(hDev);
        }

        ///<summary>
        ///Description: GSMW set base clock.
        ///Input Parameters: nClkNs -> user base clk time for pattern data, unit is ns per bit..
        ///Input Parameters: nClkMode -> the default value is RTGS_CLK_ABOVE_EQUAL.
        ///If the function succeeds, the return value is the real base clk time, the unit is ns.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_SetBaseClk(uint nClkNs, byte nClkMode)
        {
            return native_RTBB_EXTGSMW_SetBaseClk(hDev, nClkNs, nClkMode);
        }

        ///<summary>
        ///Description: GSMW get base clock.
        ///If the function succeeds, the return value is the real base clk time, the unit is ns.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_GetBaseClk()
        {
            return native_RTBB_EXTGSMW_GetBaseClk(hDev);
        }

        ///<summary>
        ///Description: GSMW get pin count.
        ///If the function succeeds, the return value is the qualities of pin for GSMW module.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_GetPinCount()
        {
            return native_RTBB_EXTGSMW_GetPinCount(hDev);
        }

        ///<summary>
        ///Description: GSMW get all pin mask.
        ///If the function succeeds, the return value is the mask value for all pins in GSMW module.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTGSMW_GetAllPinsMask()
        {
            return native_RTBB_EXTGSMW_GetAllPinsMask(hDev);
        }

        /* ExtGSMW Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_SendData")]
        private static extern int native_RTBB_EXTGSMW_SendData(IntPtr hDevice, uint nLength, uint nMask, UInt16[] pBuffer, bool bCheckAck);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_GetDataMaxCount")]
        private static extern int native_RTBB_EXTGSMW_GetDataMaxCount(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_SetBaseClk")]
        private static extern int native_RTBB_EXTGSMW_SetBaseClk(IntPtr hDevice, uint nClkNs, byte nClkMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_GetBaseClk")]
        private static extern int native_RTBB_EXTGSMW_GetBaseClk(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_GetPinCount")]
        private static extern int native_RTBB_EXTGSMW_GetPinCount(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTGSMW_GetAllPinsMask")]
        private static extern int native_RTBB_EXTGSMW_GetAllPinsMask(IntPtr hDevice);
    }
}
