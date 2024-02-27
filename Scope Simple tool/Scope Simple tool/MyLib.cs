using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scope_Simple_tool
{
    public class MyLib
    {
        public static Dictionary<int, string> RTDictionary = new Dictionary<int, string>();
        string[] RTTable = new string[] {
                "RT_BB_BAD_PARAMETER",
                "RT_BB_BUFFER_OVERFLOW",
                "RT_BB_DEVICE_BUSY",
                "RT_BB_GSMW_PIN_MASK_ERROR",
                "RT_BB_HARDWARE_NOT_FOUND",
                "RT_BB_HARDWARE_NOT_SUPPORT",
                "RT_BB_I2C_TIMEOUT",
                "RT_BB_IDLE",
                "RT_BB_MEMORY_ACCESS_ERROR",
                "RT_BB_MEMORY_ERROR",
                "RT_BB_NOT_IMPLEMENTED",
                "RT_BB_NO_ACK",
                "RT_BB_NO_DATA",
                "RT_BB_SENDING_DATA_FAILED",
                "RT_BB_SENDING_MEMORY_ADDRESS_FAILED",
                "RT_BB_SLAVE_DEVICE_NOT_FOUND",
                "RT_BB_SLAVE_OPENNING_FOR_READ_FAILED",
                "RT_BB_SLAVE_OPENNING_FOR_WRITE_FAILED",
                "RT_BB_SUCCESS",
                "RT_BB_SVI2_ALREADY_BOOTUP",
                "RT_BB_SVI2_ALREADY_POWEROFF",
                "RT_BB_SVI2_NO_POWER_OK",
                "RT_BB_SVI2_TIMEOUT",
                "RT_BB_TRANSACTION_FAILED",
                "RT_BB_UNKNOWN_ERROR",
                "RT_BB_USB_COMMUNICATION_FAILED",
                "RT_BRIDGE_ADC",
                "RT_BRIDGE_COMBO_COMMAND",
                "RT_BRIDGE_CUSTOMIZED_COMMAND",
                "RT_BRIDGE_EXTENDED",
                "RT_BRIDGE_GPIO",
                "RT_BRIDGE_GPIO_EXT",
                "RT_BRIDGE_I2C",
                "RT_BRIDGE_LED",
                "RT_BRIDGE_PWM",
                "RT_BRIDGE_SPI",
                "RT_BRIDGE_SPI_FULL_DUPLEX",
                "RT_BRIDGE_UART"
            };


        int[] RTMacro = new int[]{
            -1, -17, -11, -20, -2, -18, -14, -15, -19, -12, -9, -10, -16, -8, -7, -3, -6, -5, 0,
            -23, -24, -22, -21, -4, -13, -25, 64, 2, 2048, 1024, 8, 16, 1, 128, 32, 4, 512, 256
        };

        public MyLib()
        {
            for(int idx = 0; idx < RTTable.Length; idx++)
            {
                RTDictionary.Add(RTMacro[idx], RTTable[idx]);
            }
        }

    }
}
