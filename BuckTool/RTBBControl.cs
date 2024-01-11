using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using System.IO;
//using RTBBLibDotNet;
using RTBBLibDotNet;


namespace BuckTool
{
    public class RTBBControl
    {
        public const int GPIO2_0 = 32; // relay Iin
        public const int GPIO2_1 = 33; // relay Iout
        public const int GPIO2_2 = 34; // freq switch
        private static BridgeBoard hDevice;
        private static BridgeBoardEnum hEnum;
        private static GPIOModule gpioModule;


        public static void BoardInit()
        {
            hEnum = BridgeBoardEnum.GetBoardEnum();
            hDevice = BridgeBoard.ConnectByDefault(hEnum);
            if (hDevice != null)
                gpioModule = hDevice.GetGPIOModule();
        }


        public static void GpioInit()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_0, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_1, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_2, true);

            gpioModule.RTBB_GPIOSingleWrite(GPIO2_0, false);
            // gpio low relay 10A
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_1, false);
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_2, false);
        }

        public static void Gpio_Enable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_2, true);
        }

        public static void Gpio_Disable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_2, false);
        }

        public static void GpioEn_Enable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_1, true);
        }

        public static void GpioEn_Disable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_1, false);
        }

        public static void Meter400mA(int port)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(port, true);
        }

        public static void Meter10A(int port)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(port, false);
        }


    }
}
