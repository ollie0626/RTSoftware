using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtIOConfigModule : IBaseModule
    {
        int RTBB_EXTIOCONF_GetIOConfigurableType();
        int RTBB_EXTIOCONF_GetIOConfigurableList(int nType, int IOPort_Size, int[] pIOPort);
        int RTBB_EXTIOCONF_GetIOVoltageList(int nType, int nIOPort, int nVoltageRange, int[] pVoltageRange);
        int RTBB_EXTIOCONF_SetIOVoltageRange(int nType, int nIOPort, int min_mV, int max_mV);
        int RTBB_EXTIOCONF_SetIOVoltage(int nType, int nIOPort, int mV);
        int RTBB_EXTIOCONF_GetIOVoltage(int nType, int nIOPort);
        int RTBB_EXTIOCONF_GetIOVoltageSel(int nType, int nIOPort);
    }

    public class ExtIOConfigModule : GlobalVariable, IExtIOConfigModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtIOConfigModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtIOConfig";
        }

        ///<summary>
        ///Description: Get ExtIO IO Config type.
        ///If the function succeeds, the return value is Configuration Type of RTBridgeboard.
        ///it defines as followings,
        ///eRT_IOConfType_I2C = (1<<0),
        ///eRT_IOConfType_SPI = (1<<1),
        ///eRT_IOConfType_GPIO = (1<<2),
        ///eRT_IOConfType_PWM = (1<<3),
        ///eRT_IOConfType_GSOW = (1<<4),
        ///eRT_IOConfType_GSMW = (1<<5),
        ///eRT_IOConfType_DAC = (1<<6),
        ///eRT_IOConfType_ADC = (1<<7),
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTIOCONF_GetIOConfigurableType()
        {
            return native_RTBB_EXTIOCONF_GetIOConfigurableType(hDev);
        }

        ///<summary>
        ///Description: ExtIO get IO Config List.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: IOPort_Size -> the size of IOPort.
        ///Output Parameters: pIOPort -> Pointer to the buffer that receives the IO port number value from the interface type of bridge board configuration.
        ///If the function succeeds, the return value is the size of IO port buffer.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_GetIOConfigurableList(int nType, int IOPort_Size, int[] pIOPort)
        {
            return trans_RTBB_EXTIOCONF_GetIOConfigurableList(hDev, nType, IOPort_Size, pIOPort);
        }

        ///<summary>
        ///Description: ExtIO get IO Voltage List.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: nIOPort -> the IO port number in interface type of bridge board configuration.
        ///Input Parameters: nVoltageRange -> the size of pVoltageRange.
        ///Output Parameters: pVoltageRange -> Pointer to the buffer that receives the voltage value for IO port number in interface type of bridge board configuration.
        ///If the function succeeds, the return value is the size of voltage value buffer.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_GetIOVoltageList(int nType, int nIOPort, int nVoltageRange, int[] pVoltageRange)
        {
            return trans_RTBB_EXTIOCONF_GetIOVoltageList(hDev, nType, nIOPort, nVoltageRange, pVoltageRange);
        }

        ///<summary>
        ///Description: ExtIO set IO voltage range.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: nIOPort -> the IO port number in interface type of bridge board configuration.
        ///Input Parameters: min_mV -> the min voltage value to be set to the IO port number in interface type of bridge board configuration. the unit is mV.
        ///Input Parameters: max_mV -> the max voltage value to be set to the IO port number in interface type of bridge board configuration. the unit is mV.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_SetIOVoltageRange(int nType, int nIOPort, int min_mV, int max_mV)
        {
            return native_RTBB_EXTIOCONF_SetIOVoltageRange(hDev, nType, nIOPort, min_mV, max_mV);
        }

        ///<summary>
        ///Description: ExtIO set IO Voltage.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: nIOPort -> the IO port number in interface type of bridge board configuration.
        ///Input Parameters: mV -> set voltage value to the IO port number in interface type of bridge board configuration.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_SetIOVoltage(int nType, int nIOPort, int mV)
        {
            return native_RTBB_EXTIOCONF_SetIOVoltage(hDev, nType, nIOPort, mV);
        }

        ///<summary>
        ///Description: ExtIO Get IO voltage.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: nIOPort -> the IO port number in interface type of bridge board configuration.
        ///If the function succeeds, the return value is the voltage of IO port number in interface type of bridge board configuration. the unit is mV.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_GetIOVoltage(int nType, int nIOPort)
        {
            return native_RTBB_EXTIOCONF_GetIOVoltage(hDev, nType, nIOPort);
        }

        ///<summary>
        ///Description: ExtIO Get IO voltage selection.
        ///Input Parameters: nType -> the interface type of bridge board configuration.
        ///Input Parameters: nIOPort -> the IO port number in interface type of bridge board configuration.
        ///If the function succeeds, the return value is the voltage of IO port number in interface type of bridge board configuration. the unit is mV.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String()
        ///</summary>
        public int RTBB_EXTIOCONF_GetIOVoltageSel(int nType, int nIOPort)
        {
            return native_RTBB_EXTIOCONF_GetIOVoltageSel(hDev, nType, nIOPort);
        }

        /* ExtIOConfig Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_GetIOConfigurableType")]
        private static extern int native_RTBB_EXTIOCONF_GetIOConfigurableType(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_GetIOConfigurableList")]
        private static extern int native_RTBB_EXTIOCONF_GetIOConfigurableList(IntPtr hDevice, int nType, ref IntPtr ppIOPort);

        private int trans_RTBB_EXTIOCONF_GetIOConfigurableList(IntPtr hDevice, int nType, int IOPort_Size, int[] pIOPort)
        {
            IntPtr tmpIOPort = IntPtr.Zero;
            int ret = 0;

            ret = native_RTBB_EXTIOCONF_GetIOConfigurableList(hDevice, nType, ref tmpIOPort);
            if (ret > 0)
                Marshal.Copy(tmpIOPort, pIOPort, 0, ret);
            return ret;
        }

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_GetIOVoltageList")]
        private static extern int native_RTBB_EXTIOCONF_GetIOVoltageList(IntPtr hDevice, int nType, int nIOPort, ref IntPtr ppVoltageRange);

        private int trans_RTBB_EXTIOCONF_GetIOVoltageList(IntPtr hDevice, int nType, int nIOPort, int nVoltageRange, int[] pVoltageRange)
        {
            IntPtr tmpVoltageRange = IntPtr.Zero;
            int ret = 0;

            ret = native_RTBB_EXTIOCONF_GetIOVoltageList(hDevice, nType, nIOPort, ref tmpVoltageRange);
            if (ret > 0)
                Marshal.Copy(tmpVoltageRange, pVoltageRange, 0, ret);
            return ret;
        }

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_SetIOVoltageRange")]
        private static extern int native_RTBB_EXTIOCONF_SetIOVoltageRange(IntPtr hDevice, int nType, int nIOPort, int min_mV, int max_mV);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_SetIOVoltage")]
        private static extern int native_RTBB_EXTIOCONF_SetIOVoltage(IntPtr hDevice, int nType, int nIOPort, int mV);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_GetIOVoltage")]
        private static extern int native_RTBB_EXTIOCONF_GetIOVoltage(IntPtr hDevice, int nType, int nIOPort);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTIOCONF_GetIOVoltageSel")]
        private static extern int native_RTBB_EXTIOCONF_GetIOVoltageSel(IntPtr hDevice, int nType, int nIOPort);
    }
}
