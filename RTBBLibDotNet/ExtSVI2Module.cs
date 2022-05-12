using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtSVI2CModule : IBaseModule
    {
        int RTBB_EXTSVI2_SVI2GetFrqeuencyCapability(ref int pFrequencyCapability);
        int RTBB_EXTSVI2_SVI2GetCurrentFrequency(ref GlobalVariable.ERTExtSVI2Frequency pFrequency);
        int RTBB_EXTSVI2_SVI2SetFrequency(GlobalVariable.ERTExtSVI2Frequency nFrequencyMode);
        int RTBB_EXTSVI2_SVI2PowerUp(GlobalVariable.ERTSVI2BootVID eBootVIDCode);
        int RTBB_EXTSVI2_SVI2PowerDown();
        int RTBB_EXTSVI2_SVI2IsPowerUp();
        int RTBB_EXTSVI2_SVI2SendCmd(uint VDD_SEL, uint VDDNB_SEL, uint PSI0_L, uint PSI1_L, uint VID_CODE, uint TFN, uint LoadLineSlopeTrim, uint OffsetTrim);
        int RTBB_EXTSVI2_SVI2RecvData(ref byte nCount, uint[] RecvData);
    }

    public class ExtSVI2CModule : GlobalVariable, IExtSVI2CModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtSVI2CModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtSVI2C";
        }

        ///<summary>
        ///Description: ExtSVI2 get frequency capability.
        ///Output Parameters: pFrequency -> Current frequency of SVI2 device which has 20, 15, 10, 25, 30 or 40 MHz.
        ///                              -> Please refer to the enum ERTExtSVI2Frequency.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2GetFrqeuencyCapability(ref int pFrequencyCapability)
        {
            return native_RTBB_EXTSVI2_SVI2GetFrqeuencyCapability(hDev, ref pFrequencyCapability);
        }

        ///<summary>
        ///Description: ExtSVI2 get current frequency.
        ///Output Parameters: pFrequencyCapability -> Please refer to the enum ERTExtSVI2Frequency as the bitmask.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2GetCurrentFrequency(ref ERTExtSVI2Frequency pFrequency)
        {
            return native_RTBB_EXTSVI2_SVI2GetCurrentFrequency(hDev, ref pFrequency);
        }

        ///<summary>
        ///Description: ExtSVI2 set frequency.
        ///Input Parameters: nFrequencyMode -> Please refer to the enum ERTExtSVI2Frequency as the bitmask.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2SetFrequency(ERTExtSVI2Frequency nFrequencyMode)
        {
            return native_RTBB_EXTSVI2_SVI2SetFrequency(hDev, nFrequencyMode);
        }

        ///<summary>
        ///Description: ExtSVI2 Power up function.
        ///Input Parameters: eBootVIDCode -> Please refer to the enum ERTSVI2BootVID.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2PowerUp(GlobalVariable.ERTSVI2BootVID eBootVIDCode)
        {
            return native_RTBB_EXTSVI2_SVI2PowerUp(hDev, eBootVIDCode);
        }

        ///<summary>
        ///Description: Check SVI2 power down or not.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2PowerDown()
        {
            return native_RTBB_EXTSVI2_SVI2PowerDown(hDev);
        }

        ///<summary>
        ///Description: Check SVI2 is Powered up or not.
        ///If the function succeeds, the return value will be
        ///0: status is power off
        ///1: status is power up
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2IsPowerUp()
        {
            return native_RTBB_EXTSVI2_SVI2IsPowerUp(hDev);
        }

        ///<summary>
        ///Description: SVI2 Send command.
        ///Input Parameters: VDD_SEL ->VDD domain selector bit.
        ///Input Parameters: VDDNB_SEL -> VDDNB domain selector bit.
        ///Input Parameters: PSI0_L -> Power state indicate level 0 bit. This signal is active low.
        ///Input Parameters: PSI1_L -> Power state indicate level 1 bit. This signal is active low.
        ///Input Parameters: VID_CODE -> VID code bits[7:0].
        ///Input Parameters: TFN -> Telemetry functionality bit.
        ///Input Parameters: LoadLineSlopeTrim -> bit [2:0] Description
        ///                                    -> 000      Remove all LL droop from output
        ///                                    -> 001      Initial LL Slope -40%
        ///                                    -> 010      Initial LL Slope -20%
        ///                                    -> 011      Initial LL Slope (Default Value)
        ///                                    -> 100      Initial LL Slope +20%
        ///                                    -> 101      Initial LL Slope +40%
        ///                                    -> 110      Initial LL Slope +60%
        ///                                    -> 111      Initial LL Slope +80%
        ///Input Parameters: OffsetTrim -> bit [1:0]       Description
        ///                             -> 00              Remove all Offset from output1
        ///                             -> 01              Initial Offset -25 mV
        ///                             -> 10              Use Initial Offset(Default Value)
        ///                             -> 11              Initial Offset +25 mV
        ///If the function succeeds, the return value is the ACK data which could be zero or positive value.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2SendCmd(uint VDD_SEL, uint VDDNB_SEL, uint PSI0_L, uint PSI1_L, uint VID_CODE, uint TFN, uint LoadLineSlopeTrim, uint OffsetTrim)
        {
            return native_RTBB_EXTSVI2_SVI2SendCmd(hDev, VDD_SEL, VDDNB_SEL, PSI0_L, PSI1_L, VID_CODE, TFN, LoadLineSlopeTrim, OffsetTrim);
        }

        ///<summary>
        ///Description: SVI2 receive data.
        ///Input Parameters: nCount -> Data length received from he SVI2 bus.
        ///Output Parameters: RecvData ->  A pointer to the buffer that receives the data from the SVI2 bus.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSVI2_SVI2RecvData(ref byte nCount, uint[] RecvData)
        {
            return native_RTBB_EXTSVI2_SVI2RecvData(hDev, ref nCount, RecvData);
        }

        /* ExtSVI2 Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2GetFrqeuencyCapability")]
        private static extern int native_RTBB_EXTSVI2_SVI2GetFrqeuencyCapability(IntPtr hDevice, ref int pFrequencyCapability);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2GetCurrentFrequency")]
        private static extern int native_RTBB_EXTSVI2_SVI2GetCurrentFrequency(IntPtr hDevice, ref ERTExtSVI2Frequency pFrequency);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2SetFrequency")]
        private static extern int native_RTBB_EXTSVI2_SVI2SetFrequency(IntPtr hDevice, ERTExtSVI2Frequency nFrequencyMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2PowerUp")]
        private static extern int native_RTBB_EXTSVI2_SVI2PowerUp(IntPtr hDevice, ERTSVI2BootVID eBootVIDCode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2PowerDown")]
        private static extern int native_RTBB_EXTSVI2_SVI2PowerDown(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2IsPowerUp")]
        private static extern int native_RTBB_EXTSVI2_SVI2IsPowerUp(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2SendCmd")]
        private static extern int native_RTBB_EXTSVI2_SVI2SendCmd(IntPtr hDevice, uint VDD_SEL, uint VDDNB_SEL, uint PSI0_L, uint PSI1_L, uint VID_CODE, uint TFN, uint LoadLineSlopeTrim, uint OffsetTrim);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSVI2_SVI2RecvData")]
        private static extern int native_RTBB_EXTSVI2_SVI2RecvData(IntPtr hDevice, ref byte nCount, uint[] RecvData);
    }
}
