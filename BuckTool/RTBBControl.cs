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
        private const int GPIO2_0 = 32;
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
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_0, false);
        }

        public static void Gpio_Enable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_0, true);
        }

        public static void Gpio_Disable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_0, false);
        }

        public static void RelayOn(int num)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(num, false);
        }

        public static void RelayOff(int num)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(num, true);
        }

    }
}
