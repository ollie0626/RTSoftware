using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public class GlobalVariable
    {
        protected const string dll_path = "";

        /* BB capability */
        public const uint RT_BRIDGE_I2C = (1 << 0);
        public const uint RT_BRIDGE_COMBO_COMMAND = (1 << 1);
        public const uint RT_BRIDGE_SPI = (1 << 2);
        public const uint RT_BRIDGE_GPIO = (1 << 3);
        public const uint RT_BRIDGE_GPIO_EXT = (1 << 4);
        public const uint RT_BRIDGE_PWM = (1 << 5);
        public const uint RT_BRIDGE_ADC = (1 << 6);
        public const uint RT_BRIDGE_LED = (1 << 7);
        public const uint RT_BRIDGE_UART = (1 << 8);
        public const uint RT_BRIDGE_SPI_FULL_DUPLEX = (1 << 9);
        public const uint RT_BRIDGE_EXTENDED = (1 << 10);
        public const uint RT_BRIDGE_CUSTOMIZED_COMMAND = (1 << 11);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct RTBBInfo
        {
            [MarshalAs(UnmanagedType.LPStr)]
            public String strVendorName;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strControllerName;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strLibraryName;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strLibraryPath;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strFirmwareInfo;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strDevicePath;
            [MarshalAs(UnmanagedType.LPStr)]
            public String strBoardName;
            public int nIndexOfDevice;
            public uint nVID;
            public uint nPID;
            public UInt32 nCapability;
            public UInt32 nGPIOBitsType;
            public UInt32 nGPIOPinCount;
            public UInt32 nI2CCount;
            public UInt32 nSPICount;
            public UInt32 nUARTCount;
        };

        /* I2C frequency enum */
        public enum ERTI2CFrequency
        {
            eRTI2CFreqFast = 1,
            eRTI2CFreqStd = 2,
            eRTI2CFreq83KHz = 4,
            eRTI2CFreq71KHz = 8,
            eRTI2CFreq62KHz = 16,
            eRTI2CFreq50KHz = 32,
            eRTI2CFreq25KHz = 64,
            eRTI2CFreq10KHz = 128,
            eRTI2CFreq5KHz = 256,
            eRTI2CFreq2KHz = 512,
            eRTI2CFreq1MHz = 1024,
            eRTI2CFreqHS = 2048,
            eRTI2CFreqCustom = (1 << 30),
            eRTI2CFreqUnknow = (1 << 31),
        };

        /* I2C slave addr struct */
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public struct I2CSLAVEADDR
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public uint[] addr;
            public bool I2CSlave_Exist_Device()
            {
                for (int i = 0; i < 4; i++)
                    if (addr[i] != 0)
                        return true;
                return false;
            }
        };

        public enum ERTSPIFrequency
        {
            eRTSPIFreq400KHz = (1 << 0),
            eRTSPIFreq200KHz = (1 << 1),
            eRTSPIFreq100KHz = (1 << 2),
            eRTSPIFreq83KHz = (1 << 3),
            eRTSPIFreq71KHz = (1 << 4),
            eRTSPIFreq62KHz = (1 << 5),
            eRTSPIFreq50KHz = (1 << 6),
            eRTSPIFreq25KHz = (1 << 7),
            eRTSPIFreq10KHz = (1 << 8),
            eRTSPIFreq5KHz = (1 << 9),
            eRTSPIFreq2KHz = (1 << 10),
            eRTSPIFreq1MHz = (1 << 11),
            eRTSPIFreq2MHz = (1 << 12),
            eRTSPIFreq3MHz = (1 << 13),
            eRTSPIFreq5MHz = (1 << 14),
            eRTSPIFreq6MHz = (1 << 15),
            eRTSPIFreq10MHz = (1 << 16),
            eRTSPIFreq12MHz = (1 << 17),
            eRTSPIFreq15MHz = (1 << 18),
            eRTSPIFreq20MHz = (1 << 19),
            eRTSPIFreq24MHz = (1 << 20),
            eRTSPIFreq36MHz = (1 << 21),
            eRTSPIFreq40MHz = (1 << 22),
            eRTSPIFreq60MHz = (1 << 23),
            eRTSPIFreqCustom = (1 << 30),
            eRTSPIFreqUnknow = (1 << 31),
        };

        public enum ERTSPIMode
        {
            eSPIModeIdleLowLeadingEdge = 0,
            eSPIModeCPHA0CPOL0 = eSPIModeIdleLowLeadingEdge,
            eSPIModeIdleHighLeadingEdge = 1,
            eSPIModeCPHA0CPOL1 = eSPIModeIdleHighLeadingEdge,
            eSPIModeIdleLowTrailingEdge = 2,
            eSPIModeCPHA1CPOL0 = eSPIModeIdleLowTrailingEdge,
            eSPIModeIdleHighTrailingEdge = 3,
            eSPIModeCPHA1CPOL1 = eSPIModeIdleHighTrailingEdge,
        };

        public enum ERTExtSVI2Frequency
        {
            eRTExtSVI2Freq20MHz = 1,
            eRTExtSVI2Freq15MHz = 2,
            eRTExtSVI2Freq10MHz = 4,
            eRTExtSVI2Freq25MHz = 8,
            eRTExtSVI2Freq30MHz = 16,
            eRTExtSVI2Freq40MHz = 32,
        };

        public enum ERTSVI2BootVID
        {
            eRTExtSVI2BootVID1_1 = 0,
            eRTExtSVI2BootVID1_0 = 1,
            eRTExtSVI2BootVID0_9 = 2,
            eRTExtSVI2BootVID0_8 = 3,
        };




        public enum ERTQSPIFiledMode
        {
            eRTQSPIFieldO1A1I1D1 = (1 << 0),
            eRTQSPIFieldO1A1I1D2 = (1 << 1),
            eRTQSPIFieldO1A1I1D4 = (1 << 2),
            eRTQSPIFieldO1A1I2D2 = (1 << 3),
            eRTQSPIFieldO1A1I2D4 = (1 << 4),
            eRTQSPIFieldO1A1I4D4 = (1 << 5),
            eRTQSPIFieldO1A2I2D2 = (1 << 6),
            eRTQSPIFieldO1A2I2D4 = (1 << 7),
            eRTQSPIFieldO1A2I4D4 = (1 << 8),
            eRTQSPIFieldO1A4I4D4 = (1 << 9),
            eRTQSPIFieldO2A2I2D2 = (1 << 10),
            eRTQSPIFieldO2A2I2D4 = (1 << 11),
            eRTQSPIFieldO2A2I4D4 = (1 << 12),
            eRTQSPIFieldO2A4I4D4 = (1 << 13),
            eRTQSPIFieldO4A4I4D4 = (1 << 14),
        };

        public enum ERTQSPIdMode
        {
            RT_QSPI_SCK_MODE_CPOL0_CPHA0 = 0,
            RT_QSPI_SCK_MODE_CPOL0_CPHA1 = 1 << 1,
            RT_QSPI_SCK_MODE_CPOL1_CPHA0 = 1 << 2,
            RT_QSPI_SCK_MODE_CPOL1_CPHA1 = 1 << 3,
        };



        /* Return Err Integer Definition */
        public const int RT_BB_SUCCESS = 0;
        public const int RT_BB_BAD_PARAMETER = -1;
        public const int RT_BB_HARDWARE_NOT_FOUND = -2;
        public const int RT_BB_SLAVE_DEVICE_NOT_FOUND = -3;
        public const int RT_BB_TRANSACTION_FAILED = -4;
        public const int RT_BB_SLAVE_OPENNING_FOR_WRITE_FAILED = -5;
        public const int RT_BB_SLAVE_OPENNING_FOR_READ_FAILED = -6;
        public const int RT_BB_SENDING_MEMORY_ADDRESS_FAILED = -7;
        public const int RT_BB_SENDING_DATA_FAILED = -8;
        public const int RT_BB_NOT_IMPLEMENTED = -9;
        public const int RT_BB_NO_ACK = -10;
        public const int RT_BB_DEVICE_BUSY = -11;
        public const int RT_BB_MEMORY_ERROR = -12;
        public const int RT_BB_UNKNOWN_ERROR = -13;
        public const int RT_BB_I2C_TIMEOUT = -14;
        public const int RT_BB_IDLE = -15;
        public const int RT_BB_NO_DATA = -16;
        public const int RT_BB_BUFFER_OVERFLOW = -17;
        public const int RT_BB_HARDWARE_NOT_SUPPORT = -18;
        public const int RT_BB_MEMORY_ACCESS_ERROR = -19;
        public const int RT_BB_GSMW_PIN_MASK_ERROR = -20;
        public const int RT_BB_SVI2_TIMEOUT = -21;
        public const int RT_BB_SVI2_NO_POWER_OK = -22;
        public const int RT_BB_SVI2_ALREADY_BOOTUP = -23;
        public const int RT_BB_SVI2_ALREADY_POWEROFF = -24;
        public const int RT_BB_USB_COMMUNICATION_FAILED = -25;



        /* General Misc Functions */
        ///<summary>
        ///Input Parameters: nResult -> int result that need to be coverted.
        ///Convert result code to ANSI string.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_Result2String")]
        public static extern string RTBB_Result2String(int nResult);

        ///<summary>
        ///Input Parameters: dwMilliseconds -> milliseconds toi sleep.
        ///Suspends the execution of the current thread until the time-out interval elapses.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_MISC_Sleep")]
        public static extern void RTBB_MISC_Sleep(UInt16 dwMilliseconds);

        ///<summary>
        ///Input Parameters: nVerBCD -> firmware version bcd number.
        ///Output Parameters: majorVer -> Major version.
        ///Output Parameters: minorVer -> Minor version.
        ///Output Parameters: patchVer -> Patch version.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_ParseFirmwareVersion")]
        public static extern void RTBB_ParseFirmwareVersion(uint nVerBCD, ref byte majorVer, ref byte minorVer, ref byte patchVer);
    }
}
