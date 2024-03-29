﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using RTBBLibDotNet;

namespace OLEDLite
{
    public class RTBBControl
    {
        private const int GPIO2_0 = 32;
        private const int GPIO2_1 = 33;
        private const int GPIO2_2 = 34;

        public static int ESwire = GPIO2_0;
        public static int ASwire = GPIO2_1;
        public static int ENVO4 = GPIO2_2;

        private static BridgeBoard hDevice;
        private static BridgeBoardEnum hEnum;
        private static I2CModule i2cModule;
        private static GPIOModule gpioModule;
        private static ExtCustomizedCommandModule customizedMdoule;

        public static int[] in_gpio_table = new int[] { 32, 33, 36, 40, 41, 42, 46, 47 };
        public static int[] out_gpio_table = new int[] { 48, 49, 50, 51, 52, 53, 54, 55 };

        public RTBBControl()
        {
            Console.WriteLine("RTBBControl construct");
        }

        public void BoadInit()
        {
            hEnum = BridgeBoardEnum.GetBoardEnum();
            hDevice = BridgeBoard.ConnectByIndex(hEnum, 0);
            if(hDevice != null)
            {
                i2cModule = hDevice.GetI2CModule();
                gpioModule = hDevice.GetGPIOModule();
                customizedMdoule = hDevice.GetExtCustomizedCommandModule();
            }
        }

        public void GpioInit()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSetIODirection(1, 0xffff, 1);
            gpioModule.RTBB_GPIOWrite(1, 0xffff, 1);

            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_0, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_1, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(GPIO2_2, true);
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_0, false);
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_1, false);
            gpioModule.RTBB_GPIOSingleWrite(GPIO2_2, false);
        }

        public static void Swire_Control(int swire, bool state)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(swire, state);
        }
        
        public void RelayOn(int num)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(num, false);
        }

        public void RelayOff(int num)
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(num, true);
        }


        public int I2C_Write(byte slave, byte addr, byte[] buf)
        {
            if (i2cModule == null) return -1;
            int ret = i2cModule.RTBB_I2CWrite(slave, 0x01, addr, buf.Length, buf);
            return ret;
        }

        public int I2C_Read(byte slave, byte addr, byte[] buf)
        {
            if (i2cModule == null) return -1;
            int ret = i2cModule.RTBB_I2CRead(slave, 0x01, addr, buf.Length, buf);
            return ret;
        }

        public int I2C_WriteBin(byte slave, byte addr, string bin_file)
        {
            if (i2cModule == null) return -1;
            int ret = 0;
            FileStream file;
            BinaryReader reader;
            byte[] binData;

            file = File.Open(bin_file, FileMode.Open);
            reader = new BinaryReader(file);
            binData = reader.ReadBytes((int)file.Length);
            I2C_Write(slave, addr, binData);

            file.Close();
            reader.Close();
            return ret;
        }


        public int I2C_WriteBinAndGPIO(byte slave, byte addr, string bin_file)
        {
            if (customizedMdoule == null) return -1;
            int ret = 0;
            FileStream file;
            BinaryReader reader;
            byte[] binData;
            file = File.Open(bin_file, FileMode.Open);
            reader = new BinaryReader(file);
            binData = reader.ReadBytes((int)file.Length);

            // customized transation
            int pCmdIn = 1;
            int pDataInCount = (int)file.Length + 3;
            byte[] pDataIn = new byte[pDataInCount];
            pDataIn[0] = slave;
            pDataIn[1] = addr;
            pDataIn[2] = (byte)file.Length;
            for (int i = 3; i < (int)pDataInCount ; i++)
            {
                pDataIn[i] = binData[i - 3];
            }
            int pCmdOut = 0;
            int pDataOutCount = 0;
            customizedMdoule.RTBB_EXTCFW_Transact(ref pCmdIn, ref pDataInCount, pDataIn, ref pCmdOut, ref pDataOutCount, null);

            file.Close();
            reader.Close();
            return ret;
        }


        public static void SwirePulse(bool EASwire, int num)
        {
            if (customizedMdoule == null) return;
            // customized transation
            int pCmdIn = EASwire ? 1 : 2;
            int pDataInCount = 16;
            byte[] pDataIn = new byte[pDataInCount];
            byte[] pDataOut = new byte[pDataInCount];

            // wait time
            pDataIn[0] = 500 & 0xff;
            pDataIn[1] = (500 & 0xff00) >> 8;
            pDataIn[2] = 0x00;
            pDataIn[3] = 0x00;

            // high time
            pDataIn[4] = 5;
            pDataIn[5] = 0x00;
            pDataIn[6] = 0x00;
            pDataIn[7] = 0x00;

            // low time
            pDataIn[8] = 5;
            pDataIn[9] = 0x00;
            pDataIn[10] = 0x00;
            pDataIn[11] = 0x00;

            // pulse
            pDataIn[12] = (byte)(num & 0xff);
            pDataIn[13] = (byte)((num & 0xff00) >> 8);
            pDataIn[14] = 0x00;
            pDataIn[15] = 0x00;

            int pCmdOut = 0;
            int pDataOutCount = 0;
            int ret;
            if (num != 0) ret = customizedMdoule.RTBB_EXTCFW_Transact(ref pCmdIn, ref pDataInCount, pDataIn, ref pCmdOut, ref pDataOutCount, pDataOut);

            //Console.WriteLine("ret = {0}", ret);
        }

    }
}
