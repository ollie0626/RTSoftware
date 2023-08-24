using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RTBBLibDotNet;


namespace CS601C
{
    public class RTBBControl
    {
        private BridgeBoard hDevice;
        private BridgeBoardEnum hEnum;
        private I2CModule i2cModule;

        public bool BoadInit()
        {
            hEnum = BridgeBoardEnum.GetBoardEnum();
            hDevice = BridgeBoard.ConnectByIndex(hEnum, 0);
            if (hDevice != null)
            {
                i2cModule = hDevice.GetI2CModule();
                return true;
            }
            return false;
        }


        public List<byte> ScanSlaveID()
        {
            
            if (!BoadInit()) return null;

            GlobalVariable.I2CSLAVEADDR i2CSLAVEADDR = new GlobalVariable.I2CSLAVEADDR();
            List<byte> slave_list = new List<byte>();

            for (int slave = 0; slave < 128; slave += 2)
            {
                i2cModule.RTBB_I2CScanSlaveDevice(ref i2CSLAVEADDR);
                slave = i2cModule.RTBB_I2CGetFirstValidSlaveAddr(ref i2CSLAVEADDR, slave);
                bool valid = i2cModule.I2C_SLAVE_ADDR_IS_VALID(i2CSLAVEADDR, slave);

                if (valid)
                    slave_list.Add((byte)(slave));
                else
                    break;
            }
            return slave_list;
        }

        public int I2C_Write(byte slave, byte addr, byte[] buf)
        {
            if (i2cModule == null) return -1;
            int ret = i2cModule.RTBB_I2CWrite(slave, 0x01, addr, buf.Length, buf);
            return ret;
        }

        public int I2C_Read(byte slave, byte addr, ref byte[] buf)
        {
            if (i2cModule == null) return -1;
            int ret = i2cModule.RTBB_I2CRead(slave, 0x01, addr, buf.Length, buf);
            return ret;
        }
    }
}
