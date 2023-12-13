using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using RTBBLibDotNet;
using System.IO;


namespace MulanLite
{
    public class RTBBControl
    {
        private BridgeBoardEnum hEnum;
        public BridgeBoard hDevice;
        private SPIModule spiModule;
        private GPIOModule gpioModule;
        private PWMModule pwmModule;
        private I2CModule i2cMoudle;

        private const int CS_Pin = 2;
        private const int GPIO_2_0 = 32;        /* trans en */
        private const int GPIO_2_1 = 33;        /* CO enable */
        private const int GPIO_2_2 = 34;        /* for POR */
        private const int SPI_BUF_LEN = 30;

        private const int Trans_en = GPIO_2_0;
        private const int Co_en = GPIO_2_1;
        private const int POR = GPIO_2_2;
        public static bool CRC_En = true;


        public void CiEnable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(Co_en, true);
        }

        public void CiDisable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(Co_en, false);
        }

        public void POREnable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(POR, true);
        }

        public void PORDisable()
        {
            if (gpioModule == null) return;
            gpioModule.RTBB_GPIOSingleWrite(POR, false);
        }

        public void I2cWrite(byte slave, byte addr, byte[] data)
        {
            if (i2cMoudle == null) return;
            i2cMoudle.RTBB_I2CWrite(slave, 0x01, addr, data.Length, data);
        }

        public void SPIWrite(byte[] buf)
        {
            if (spiModule == null) return;
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x00, (ushort)buf.Length, 0x00, buf);
        }

        /* 
         * idx = 0, Ci = 1%
         * idx = 1, Ci = 2%
         * idx = 2, Ci = 3%
         * idx = 3, Ci = 4%
         * idx = 4, Ci = 5%
         * idx = 5, Ci = 6%
         * idx = 6, Ci = 7%
         * idx = 7, Ci = 8%
         * idx = 8, Ci = 9%
         * 
         */
        public void SetCiClock(int idx)
        {
            if (pwmModule == null) return;
            if (spiModule == null) return;

            byte[] buf = new byte[2];
            byte CmdSize = 0x01;
            uint Cmd = 0xAD;

            switch (idx)
            {

                case 0:
                    // 6M (1%)
                    buf[0] = 0xBA;
                    buf[1] = 0x61;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, 2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 1:
                    buf[0] = 0xBA;
                    buf[1] = 0x62;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);

                    break;
                case 2:
                    pwmModule.RTBB_PWMSetPeriod(0, 3);
                    buf[0] = 0xBA;
                    buf[1] = 0x63;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 3:
                    buf[0] = 0xBA;
                    buf[1] = 0x64;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 4:
                    buf[0] = 0xBA;
                    buf[1] = 0x65;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 5:
                    buf[0] = 0xBA;
                    buf[1] = 0x66;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 6:
                    buf[0] = 0xBA;
                    buf[1] = 0x67;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 7:
                    buf[0] = 0xBA;
                    buf[1] = 0x68;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;                
                case 8:
                    buf[0] = 0xBA;
                    buf[1] = 0x69;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 9:
                    buf[0] = 0xBA;
                    buf[1] = 0x6A;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 10:
                    buf[0] = 0xBA;
                    buf[1] = 0x6B;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 11:
                    buf[0] = 0xBA;
                    buf[1] = 0x6C;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 12:
                    buf[0] = 0xBA;
                    buf[1] = 0x6D;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 13:
                    buf[0] = 0xBA;
                    buf[1] = 0x6E;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 14:
                    buf[0] = 0xBA;
                    buf[1] = 0x6F;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 15:
                    buf[0] = 0xBA;
                    buf[1] = 0x70;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 16:
                    buf[0] = 0xBA;
                    buf[1] = 0x71;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                case 17:
                    buf[0] = 0xBA;
                    buf[1] = 0x72;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
                default:
                    buf[0] = 0xBA;
                    buf[1] = 0x75;
                    spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
                    spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)2, Cmd, buf);
                    System.Threading.Thread.Sleep(20);
                    break;
            }
        }

        public bool BoardInit()
        {
            hEnum = null;
            hDevice = null;

            hEnum = BridgeBoardEnum.GetBoardEnum();
            hDevice = BridgeBoard.ConnectByIndex(hEnum, 0);

            if (hDevice == null) return false;

            spiModule = hDevice.GetSPIModule();
            gpioModule = hDevice.GetGPIOModule();
            pwmModule = hDevice.GetPWMModule();
            i2cMoudle = hDevice.GetI2CModule();

            gpioModule.RTBB_GPIOSingleSetIODirection(Trans_en, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(Co_en, true);
            gpioModule.RTBB_GPIOSingleSetIODirection(POR, true);

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
            gpioModule.RTBB_GPIOSingleWrite(Co_en, true);
            gpioModule.RTBB_GPIOSingleWrite(POR, false);
            pwmModule.RTBB_PWMSetPeriod(0, 1); /* pwm 36MHz */
            pwmModule.RTBB_PWMSetDutyCycle(0, 0.5);
            pwmModule.RTBB_PWMStart(0);
            return true;

        }

        public void BoardRemove()
        {
            hDevice = null;
            hEnum = null;
            spiModule = null;
            i2cMoudle = null;
            pwmModule = null;
            GC.Collect();
        }

        public byte CRC_8(byte[] buffer)
        {
            byte crc = 0xFF;
            for (int j = 0; j < buffer.Length - 2; j++)
            {
                crc ^= buffer[j];

                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x80) != 0)
                    {
                        crc <<= 1;
                        crc ^= 0x07;
                    }
                    else
                    {
                        crc <<= 1;
                    }
                }
            }

            return CRC_En ? crc : (byte)0;
        }

        public void WriteBin(byte id, string Path)
        {
            if (!File.Exists(Path)) return;

            FileStream fsFile = new FileStream(Path, FileMode.Open);
            BinaryReader bReader = new BinaryReader(fsFile);
            byte[] WData = new byte[SPI_BUF_LEN];
            for (int i = 0; i < 0x63; i++)
                WData[i] = bReader.ReadByte();

            WriteFunc(id, 0x2D, 0x00, 0x62 - 1, WData);

            fsFile.Close();
            bReader.Close();

            fsFile.Dispose();
            bReader.Dispose();
        }

        public void Identify(byte id)
        {
            if (spiModule == null) return;

            byte sync_Cmd = 0xAC;
            byte cmd = 0x78;
            int bit7 = ((id & 0x80) >> 7);
            int bit6 = ((id & 0x40) >> 6);
            int bit5 = ((id & 0x20) >> 5);
            int bit4 = ((id & 0x10) >> 4);
            int bit3 = ((id & 0x08) >> 3);
            int bit2 = ((id & 0x04) >> 2);
            int bit1 = ((id & 0x02) >> 1);
            int bit0 = ((id & 0x01) >> 0);


            byte identify_id = (byte)((bit0 << 7) | (bit1 << 6) | (bit2 << 5) | (bit3 << 4) | (bit4 << 3) | (bit5 << 2) | (bit6 << 1) | (bit7 << 0));
            byte[] tmp = new byte[4];
            tmp[0] = cmd;
            tmp[1] = 0x00;
            tmp[2] = identify_id;
            tmp[3] = 0;

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x01, (ushort)tmp.Length, sync_Cmd, tmp);
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
        }

        public byte[] Inquiry()
        {
            byte[] ReadPacket = new byte[13];
            byte sync_Cmd = 0xAC;
            byte cmd = 0x4B;
            ReadPacket[0] = cmd;
            ReadPacket[1] = 0x00;
            ReadPacket[2] = 0x00;
            ReadPacket[3] = 0x00;
            ReadPacket[4] = 0x00;

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            //System.Threading.Thread.Sleep(2);
            Task.Delay(2).Wait();
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x01, (ushort)0x05, sync_Cmd, ReadPacket);

            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA1CPOL0);
            Task.Delay(100).Wait();
            spiModule.RTBB_SPIHLReadCS(CS_Pin, 0, (ushort)ReadPacket.Length, 0xAC, ReadPacket);
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);

            byte item = 0xac;
            int idx = Array.IndexOf(ReadPacket, item);
            byte[] data = ReadPacket.Skip(idx).ToArray();
            WriteFunc(0x00, 0x00, 0x00, 0x00, new byte[1] { 0x00 }, 0x00);
            return data;
        }

        public byte[] ResponesID(byte flag)
        {
            byte[] ReadPacket = new byte[13];

            byte sync_Cmd = 0xAC;
            byte cmd = 0x69;
            ReadPacket[0] = cmd;
            ReadPacket[1] = flag;
            ReadPacket[2] = 0x00;
            ReadPacket[3] = 0x00;

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            Task.Delay(2).Wait();
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x01, (ushort)0x05, sync_Cmd, ReadPacket);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA1CPOL0);
            Task.Delay(100).Wait();
            spiModule.RTBB_SPIHLReadCS(CS_Pin, 0, (ushort)ReadPacket.Length, 0xAC, ReadPacket);
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);

            byte item = 0xac;
            int idx = Array.IndexOf(ReadPacket, item);
            byte[] data = ReadPacket.Skip(idx).ToArray();

            WriteFunc(0x00, 0x00, 0x00, 0x00, new byte[1] { 0x00 }, 0x00);
            return data;
        }

        public void BLEnable(byte id)
        {
            byte[] WData = new byte[1] { 0x02 };
            WriteFunc(id, 0x2D, 0x03, 0x00, WData);
        }

        public void BLUpdate()
        {
            if (spiModule == null) return;
            byte sync_Cmd = 0xAC;
            byte[] tmp = new byte[0x02];
            tmp[0] = 0x5A;
            tmp[1] = 0x00;

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x01, (ushort)0x02, sync_Cmd, tmp);
            System.Threading.Thread.Sleep(20);
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
        }

        public int WriteFunc(byte id, byte cmd, byte addr, int len, byte[] buf, byte Sync_Cmd = 0xAC)
        {
            if (spiModule == null) return 1;

            //byte Sync_Cmd = 0xAC;
            byte invid = (byte)~id;
            byte addr_M = 0x00;
            byte addr_L = addr;
            byte CmdSize = 0x01;

            byte[] tmp = new byte[9 + len]; // add CRC + FPGA dummy
            tmp[0] = cmd;
            tmp[1] = id;
            tmp[2] = invid;
            tmp[3] = (byte)(len);
            tmp[4] = addr_M;
            tmp[5] = addr_L;

            for (int i = 0; i < len + 1; i++)
            {
                tmp[i + 6] = buf[i];
            }

            byte[] CRC_buf = new byte[9 + len];
            Array.Copy(tmp, CRC_buf, CRC_buf.Length);
            byte CRC8 = CRC_8(CRC_buf);
            tmp[tmp.Length - 2] = CRC8;
            tmp[tmp.Length - 1] = 0;    // for FPGA dummy byte

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            if (Sync_Cmd == 0xAC) spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)tmp.Length, Sync_Cmd, tmp);
            else
            {
                Array.Resize(ref buf, buf.Length + 1);
                spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)buf.Length, Sync_Cmd, buf);
            }
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
            return 0;
        }

        public void LEDPacket(byte len, int addr, int[] buf)
        {
            if (spiModule == null) return;

            byte sysnc_cmd = 0xAC;
            byte cmd = 0x3C;
            byte num_zones = len;
            byte addr_M = (byte)((addr & 0xFF00) >> 8);
            byte addr_L = (byte)(addr & 0xFF);


            List<byte> packet = new List<byte>();
            packet.Add(cmd);
            packet.Add(num_zones);
            packet.Add(addr_M);
            packet.Add(addr_L);
            int shift1 = 9;
            int shift2 = 1;
            int shift3 = 7;
            int idx = 4;
            foreach (int data in buf)
            {
                if (idx >= packet.Count) packet.Add(0);
                packet[idx] = (byte)((data >> shift1) | packet[idx++]); shift1++;
                if (idx >= packet.Count) packet.Add(0);
                packet[idx++] = (byte)(data >> shift2); shift2++;
                if (idx >= packet.Count) packet.Add(0);
                packet[idx] = (byte)(data << shift3); shift3--;

                if (shift3 == -1)
                {
                    shift1 = 9;
                    shift2 = 1;
                    shift3 = 7;
                    idx++;
                }
            }

            byte[] CRC_buf = new byte[packet.Count];
            Array.Copy(packet.ToArray(), CRC_buf, CRC_buf.Length);
            packet.Add(0);
            packet.Add(0);
            byte CRC8 = CRC_8(packet.ToArray());
            packet[packet.Count - 2] = CRC8;
            packet[packet.Count - 1] = 0;


            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, 0x01, (ushort)packet.Count, sysnc_cmd, packet.ToArray());
            System.Threading.Thread.Sleep(20);
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
        }

        public int WriteBADID(byte id, byte cmd, byte addr, byte len, byte[] buf, byte Sync_Cmd = 0xAC)
        {
            if (spiModule == null) return 1;

            byte invid = (byte)(0xAA ^ id);
            byte addr_M = 0x00;
            byte addr_L = addr;
            byte CmdSize = 0x01;

            byte[] tmp = new byte[9 + len]; // add CRC + FPGA dummy
            tmp[0] = cmd;
            tmp[1] = id;
            tmp[2] = invid;
            tmp[3] = (byte)(len);
            tmp[4] = addr_M;
            tmp[5] = addr_L;

            for (int i = 0; i < len + 1; i++)
            {
                tmp[i + 6] = buf[i];
            }

            byte[] CRC_buf = new byte[9 + len];
            Array.Copy(tmp, CRC_buf, CRC_buf.Length);
            byte CRC8 = CRC_8(CRC_buf);
            tmp[tmp.Length - 2] = CRC8;
            tmp[tmp.Length - 1] = 0;    // for FPGA dummy byte

            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            if (Sync_Cmd == 0xAC) spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)tmp.Length, Sync_Cmd, tmp);
            else spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)tmp.Length, Sync_Cmd, new byte[] { 0x00 });
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
            return 0;
        }

        /* len follow packet setting (n + 1) */
        public byte[] ReadFunc(byte id, byte len, byte addr, uint Cmd = 0xAC)
        {
            if (spiModule == null) return new byte[10];
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            byte CmdSize = 0x01;
            //uint Cmd = 0xAC;

            byte[] buf = new byte[7];
            buf[0] = 0x1E;
            buf[1] = id;
            buf[2] = (byte)~id;
            buf[3] = len;
            buf[4] = 0x00;
            buf[5] = addr;
            buf[6] = 0x00; // for FPGA dummy byte
            //Repeat:
            /* write command */
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            //System.Threading.Thread.Sleep(2);
            Task.Delay(2).Wait();
            if (Cmd == 0xAC)
            {
                spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)(buf.Length), Cmd, buf);
            }
            else
            {
                Array.Resize(ref buf, buf.Length + 1);
                spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)(buf.Length), Cmd, buf);
            }
            

            Task.Delay(2).Wait();
            byte[] Buffer_tmp = new byte[13];

            /* read command */
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA1CPOL0);
            Task.Delay(2).Wait();
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA1CPOL0);
            Task.Delay(2).Wait();
            spiModule.RTBB_SPIHLReadCS(CS_Pin, 0, (ushort)(Buffer_tmp.Length), Cmd, Buffer_tmp);
            for (int i = 0; i < Buffer_tmp.Length; i++) Console.Write("{0:X}, ", Buffer_tmp[i]);
            Console.WriteLine();


            //if (Enumerable.SequenceEqual(Buffer_tmp, new byte[13]))
            //{ goto Repeat; }

            byte item = (byte)Cmd;
            int idx = Array.IndexOf(Buffer_tmp, item);
            byte[] data = Buffer_tmp.Skip(idx).ToArray();
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, false);
            WriteFunc(0x00, 0x00, 0x00, 0x00, new byte[1] { 0x00 }, 0x00);

            if (data.Length < 3) data = new byte[12];
            return data;
        }

        public void RDOFlow_Flag(byte id)
        {
            
            byte CmdSize = 0x01;
            uint Cmd = 0xAC;

            byte[] buf = new byte[18];
            buf[0] = 0x1E;
            buf[1] = id;
            buf[2] = (byte)~id;
            buf[3] = 0x00;
            buf[4] = 0x00;
            buf[5] = 0x32;
            buf[6] = 0x00; // for FPGA dummy byte

            buf[7] = 0;
            buf[8] = 0;
            buf[9] = 0;

            buf[10] = 0xAC;
            buf[11] = 0x1E;
            buf[12] = id;
            buf[13] = (byte)~id;
            buf[14] = 0x00;
            buf[15] = 0x00;
            buf[16] = 0x32;
            buf[17] = 0x00; // for FPGA dummy byte

            /* write command */
            gpioModule.RTBB_GPIOSingleWrite(Trans_en, true);
            spiModule.RTBB_SPISetMode((uint)GlobalVariable.ERTSPIMode.eSPIModeCPHA0CPOL0);
            Task.Delay(2).Wait();
            spiModule.RTBB_SPIHLWriteCS(CS_Pin, CmdSize, (ushort)(buf.Length), Cmd, buf);
        }
    }
}
