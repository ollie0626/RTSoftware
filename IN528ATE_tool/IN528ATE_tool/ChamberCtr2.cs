using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net.Sockets;
using System.IO;
using System.Net;

namespace IN528ATE_tool
{
    public class ChamberCtr
    {
        static string ChamberDirPath = @"\\192.168.10.144\temp\DPBU\Chamber\";
        static public string ChamberName = "Mulan";
        static private bool IsFolderExist;
        static public bool IsTCPConnected;
        static public int FailCnt = 0;
        //static string ChamberTemps = "";
        static string GetMasterIP = "";
        //static string GetMasterIP = "";
        static public TcpListener myTcpListener = null;
        static public Socket mySocket = null;
        //宣告網路資料流變數
        static NetworkStream myNetworkStream = null;
        //宣告 Tcp 用戶端物件
        static TcpClient myTcpClient;
        //Socket mySocket = myTcpListener.AcceptSocket();

        static public void CreatShareChamberFolder()
        {
            System.IO.Directory.CreateDirectory(ChamberDirPath + ChamberName);
            IsFolderExist = Directory.Exists(ChamberDirPath + ChamberName);
            FailCnt = 0;
        }

        static public bool CreatTCPServer()
        {
            if (myTcpListener == null)
            {
                IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
                IPAddress[] addr = ipEntry.AddressList;
                for (int i = 0; i < addr.Length; i++)
                {
                    Console.WriteLine("IP Address {0}: {1} ", i, addr[i].ToString());
                }
                myTcpListener = new TcpListener(addr[addr.Length - 1], 36000);
            }
            myTcpListener.Start();
            return true;
        }

        static public bool CreatSlaveConnect()
        {
            if (IsTCPConnected)
            {
                myTcpClient.Close();
                IsTCPConnected = false;
            }
            try
            {
                myTcpClient = new TcpClient();
                myTcpClient.Connect(GetMasterIP, 36000);
                myTcpClient.Close();
                return true;
            }
            catch
            {
                //  Console.WriteLine ("主機 {0} 通訊埠 {1} 無法連接  !!", GetMasterIP, 360000);
            }
            myTcpClient.Close();
            return false;
        }

        static public void WriteData2Master()
        {
            String strTest = "SlaveTestOK";
            //將字串轉 byte 陣列，使用 ASCII 編碼
            Byte[] myBytes = Encoding.ASCII.GetBytes(strTest);
            myNetworkStream = myTcpClient.GetStream();
            myNetworkStream.Write(myBytes, 0, myBytes.Length);
        }

        static public void DeleteShareChamberFile()
        {
            if(IsFolderExist)
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(ChamberDirPath + ChamberName);
                foreach(FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
        }

        static public void DeleteFolder()
        {
            if (IsFolderExist)
                System.IO.Directory.Delete(ChamberDirPath + ChamberName, true);
        }

        static public string GetIP()
        {
            IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
            //IPAddress[] addr = ;
            return ipEntry.AddressList.Last().ToString();
        }

        static public string ReadTempList()
        {
            //IsTCPConnected = false;
            if (IsFolderExist)
            {
                DirectoryInfo di = new DirectoryInfo(ChamberDirPath + ChamberName);
                FileInfo[] afi = di.GetFiles();
                for (int i = 0; i < afi.Length; ++i)
                {
                    Console.WriteLine(afi[i]);
                    if (afi[i].Name.Contains("TempList"))
                    {
                        GetMasterIP = File.ReadAllText(afi[i].FullName);
                        GetMasterIP = GetMasterIP.Substring(0, GetMasterIP.Length - 2);
                        Console.WriteLine(GetMasterIP);
                        return afi[i].Name.Split('_')[2];
                    }
                }
                return "";
            }
            return "";
        }

        static public void CreatTempList(string TempList)
        {
            if(IsFolderExist)
            {
                string time = DateTime.Now.ToString("yyyyMMddHHmmss");
                string path = string.Format("{0}\\{1}_TempList_{2}", ChamberDirPath + ChamberName, time, TempList);
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine(GetIP());
                    sw.Close();
                }
                FailCnt = 0;
            }
        }

        static public void WriteCurrTemp()
        {
            if (IsFolderExist)
            {
                string time = DateTime.Now.ToString("yyyyMMddHHmmss");
                string path = string.Format("{0}\\{1}_CurrTemp_{2}", ChamberDirPath + ChamberName, time, 10);
                using (StreamWriter sw = File.CreateText(path)) { }
            }
        }

        static public void WriteTestFin(bool IsSlave = true)
        {
            string str = !IsSlave ? "{0}\\{1}_MasterFin_{2}" : "{0}\\{1}_SlaveFin_{2}";
            if (IsFolderExist)
            {
                string time = DateTime.Now.ToString("yyyyMMddHHmmss");
                string path = string.Format(str, ChamberDirPath + ChamberName, time, 10);
                using (StreamWriter sw = File.CreateText(path)) { }
            }
        }

        static public void WriteDeviceAlive(bool IsSlave = true)
        {
            string str = !IsSlave ? "{0}\\{1}_MasterLive_{2}" : "{0}\\{1}_SlaveLive_{2}";
            if (IsFolderExist)
            {
                string time = DateTime.Now.ToString("yyyyMMddHHmmss");
                string path = string.Format(str, ChamberDirPath + ChamberName, time, 10);
                using (StreamWriter sw = File.CreateText(path)) { }
            }
        }

    }
}
