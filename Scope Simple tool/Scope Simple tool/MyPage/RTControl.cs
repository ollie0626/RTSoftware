using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RTBBLibDotNet;
using System.IO;


namespace Scope_Simple_tool.MyPage
{
    public class RTControl
    {
        static public BridgeBoard hBoard = null;
        static public BridgeBoardEnum hEnum;
        static public string DeviceName;
        static public bool IsUse;
        static public I2CModule i2cModule;


        public static bool InitBridgeBoard()
        {
            hEnum = BridgeBoardEnum.GetBoardEnum();
            DeviceName = "No Connected";
            IsUse = false;

            hBoard = BridgeBoard.ConnectByDefault(hEnum);
            if (hBoard == null) return false;
            i2cModule = hBoard.GetI2CModule();

            RTBBLibDotNet.GlobalVariable.RTBBInfo Info = new GlobalVariable.RTBBInfo();
            hEnum.RTBB_GetEnumBoardInfo(0, ref Info);
            DeviceName = Info.strBoardName;

            Pages.LuaWindowViewModel.LuaWindowMessage(DeviceName + " Connect Success !!!");
            return true;
        }


        static public void I2cWriteBinfile(int slave, int addr, string path)
        {
            byte[] BinBuffer;
            FileStream Fio = File.Open(path, FileMode.Open);
            BinaryReader binRead = new BinaryReader(Fio);
            FileInfo fileInfo = new FileInfo(path);
            BinBuffer = binRead.ReadBytes((int)fileInfo.Length);

            if (i2cModule != null)
            {
                int res = i2cModule.RTBB_I2CWrite(slave >> 1, 0x01, 0x00, BinBuffer.Length, BinBuffer);
                Pages.LuaWindowViewModel.LuaWindowMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}", slave, addr));
                Pages.SubWindow.PrintDebugMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}", slave, addr));
            }
            else
            {
                Pages.LuaWindowViewModel.LuaWindowMessage("I2c Model null");
                Pages.SubWindow.PrintDebugMessage("I2c Model null");
            }
            Fio.Close();
            binRead.Close();
        }

        static public void I2c_SingleWrite(int slave, int addr, byte data)
        {
            byte[] DataBuf = new byte[1];
            DataBuf[0] = data;
            if (i2cModule != null)
            {
                int res = i2cModule.RTBB_I2CWrite(slave >> 1, 0x01, addr, 1, DataBuf);
                Pages.LuaWindowViewModel.LuaWindowMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}, Data {2:X}", slave, addr, data));
                Pages.SubWindow.PrintDebugMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}", slave, addr));
            }
            else
            {
                Pages.LuaWindowViewModel.LuaWindowMessage("I2c Model null");
                Pages.SubWindow.PrintDebugMessage("I2c Model null");
            }

        }


        static public void I2c_MultiWrite(int slave, int addr, byte[] buf)
        {
            if (i2cModule != null)
            {
                int res = i2cModule.RTBB_I2CWrite(slave >> 1, 0x01, addr, 1, buf);
                Pages.LuaWindowViewModel.LuaWindowMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}", slave, addr));
                Pages.SubWindow.PrintDebugMessage(string.Format(Scope_Simple_tool.MyLib.RTDictionary[res] + " : I2c Write Slave {0:X}, Addr {1:X}", slave, addr));
            }
            else
            {

            }
        }





    }
}
