using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace IN528ATE_tool
{
    public class ChamberCtr
    {
        static string ChamberDirPath = @"\\192.168.10.144\temp\DPBU\Chamber\";
        static public string ChamberName = "PMIC2_0";
        static private bool IsFolderExist;
        static public bool IsTCPConnected = false;
        static public int FailCnt = 0;
        static string ChamberTemps = "";
        static string GetMasterIP = "";
        //static string GetMasterIP = "";
        static public TcpListener myTcpListener = null;
        static public Socket mySocket = null;
        //宣告網路資料流變數
        static NetworkStream myNetworkStream = null;
        //宣告 Tcp 用戶端物件
        static TcpClient myTcpClient;
        static byte[] receiveBytes = new byte[512];
        //Socket mySocket = myTcpListener.AcceptSocket();
        static System.Timers.Timer Mytimer;
        static bool MasterFlag = true;
        static int FailCNT = 0;
        static public bool IsTCPNoConnected = false;
        static public double CurrentTemp = 0.0;
        static public string CurrentStateMaster = "Idle,25";
        static public string CurrenStateSlave = "No,25";


        static public void CreatShareChamberFolder()
        {
            System.IO.Directory.CreateDirectory(ChamberDirPath + ChamberName);
            IsFolderExist = Directory.Exists(ChamberDirPath + ChamberName);
            FailCnt = 0;
        }

        static public void InitTCPTimer(bool IsMaster)
        {
            int interval = 5000;
            FailCNT = 0;
            // IsTCPNoConnected = false;
            if (Mytimer == null)
            {
                Mytimer = new System.Timers.Timer(interval);
                //設定重複計時
                Mytimer.AutoReset = true;
                //設定執行System.Timers.Timer.Elapsed事件
                Mytimer.Elapsed += new System.Timers.ElapsedEventHandler(Mytimer_tick);
            }
            Mytimer.Stop();
            MasterFlag = IsMaster;
            Console.WriteLine("Init TCP Timer");
        }

        static public void SetTCPTimerState(bool IsRun)
        {
            FailCNT = 0;
            //IsTCPNoConnected = false;
            if (Mytimer != null)
            {
                if (IsRun) Mytimer.Start();
                else
                {
                    IsTCPNoConnected = true;
                    try
                    {
                        Console.WriteLine("TCP Stop");
                        Mytimer.Stop();
                        mySocket.Close();
                        myTcpListener.Stop();
                    }
                    catch { }
                }
            }
        }

        static private void Mytimer_tick(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                if (MasterFlag)
                    CreatTCPServer();
                else
                    CreatSlaveConnect();
            }
            catch { }
            //Console.WriteLine("IsSameState :" + ChamberCtr.CheckAllIdle() + "   Fail:" + FailCNT.ToString());
            Console.WriteLine("{3}:Master {0}, Slave {1}  --- {2}", CurrentStateMaster, CurrenStateSlave, FailCNT, MasterFlag);

        }

        static public bool CreatTCPServer()
        {
            //if(IsTCPConnected && myTcpListener != null)
            // myTcpListener.Stop();

            //IsTCPConnected = false;
            /* if (mySocket != null)
             {
                 mySocket.Close();
             }*/
            if (FailCNT > 2) CurrenStateSlave = "";
            if (IsTCPNoConnected) return true;
            if (FailCNT > 60) IsTCPNoConnected = true;
            else IsTCPNoConnected = false;
            if (myTcpListener == null)
            {
                IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
                IPAddress[] addr = ipEntry.AddressList;
                for (int i = 0; i < addr.Length; i++)
                {
                    Console.WriteLine("IP Address {0}: {1} ", i, addr[i].ToString());
                }
                // myTcpListener = new TcpListener(addr[addr.Length - 1], 36000);
                myTcpListener = new TcpListener(addr[addr.Length - 1], 36000);
            }
            myTcpListener.Start();

            try
            {
                if (myTcpListener.Pending())
                {
                    FailCNT = 0;
                    mySocket = myTcpListener.AcceptSocket();
                    ///向客戶端傳送一條訊息
                    byte[] date = System.Text.Encoding.UTF8.GetBytes(CurrentStateMaster);//轉換成為bytes陣列
                    mySocket.Send(date);
                    ///接收一條客戶端的訊息
                    //byte[] dateBuffer = new byte[1024];
                    int count = mySocket.Receive(receiveBytes);
                    CurrenStateSlave = System.Text.Encoding.UTF8.GetString(receiveBytes, 0, count);
                    //Console.WriteLine("Back : " + CurrenStateSlave);
                    Thread.Sleep(50);
                    mySocket.Close();
                    myTcpListener.Stop();
                }
                else
                {
                    //   Console.WriteLine("Slave {0} 通訊埠 {1} 無法連接  !!", GetMasterIP, 360000);
                    FailCNT++;
                }
            }
            catch { FailCNT++; }
            return true;
        }

        static public bool CreatSlaveConnect()
        {
            /*  if(myTcpClient == null)

              if(IsTCPConnected)
              {
                  myTcpClient.Close();
                  IsTCPConnected = false;
              }
                  */
            if (FailCNT > 2) CurrentStateMaster = "";
            if (IsTCPNoConnected) return true;

            if (IsTCPConnected)
            {
                myTcpClient.Close();
                IsTCPConnected = false;
            }

            if (FailCNT > 60) IsTCPNoConnected = true;
            else IsTCPNoConnected = false;

            try
            {
                myTcpClient = new TcpClient();
                //測試連線至遠端主機 
                myTcpClient.Connect(GetMasterIP, 36000);
                NetworkStream ns = myTcpClient.GetStream();
                int Len = ns.Read(receiveBytes, 0, receiveBytes.Length);
                CurrentStateMaster = Encoding.Default.GetString(receiveBytes, 0, Len);
                //Console.WriteLine(CurrentStateMaster);
                FailCNT = 0;
                byte[] msgByte = Encoding.Default.GetBytes(CurrenStateSlave);
                ns.Write(msgByte, 0, msgByte.Length);
                Thread.Sleep(50);
                myTcpClient.Close();
                return true;
            }
            catch
            {
                //Console.WriteLine ("主機 {0} 通訊埠 {1} 無法連接  !!", GetMasterIP, 360000);
                FailCNT++;
            }
            myTcpClient.Close();
            return false;
        }

        static public bool CheckAllIdle()
        {
            if (CurrentStateMaster.Contains("Idle") && CurrenStateSlave.Contains("Idle"))
            {
                if (CurrentStateMaster == CurrenStateSlave) return true;
                return false;
            }
            return false;
        }
        //寫入資料
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
            if (IsFolderExist)
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(ChamberDirPath + ChamberName);

                foreach (FileInfo file in di.GetFiles())
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

        static public void CreatTempList(string TempList)
        {
            if (IsFolderExist)
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

        static public bool CheckTCP_ChamberIdle()
        {
            int Time_1S = 9999999;
            string tmpObj1 = "";
            string tmpObj2 = "";
            int CNT = 0;
            int CNT2 = 0;
     
            for (int i = 0; i < Time_1S; ++i)
            {
                if (IsTCPNoConnected)
                {
                    Console.WriteLine("TCP No Connected....");
                    return true;
                }
                Console.WriteLine("*******State 1: {0}----, 2:{1} , {2}", CurrentStateMaster, CurrenStateSlave, CNT2);
                if (!string.IsNullOrEmpty(CurrentStateMaster) && !string.IsNullOrEmpty(CurrenStateSlave))
                {
                    if (CheckAllIdle()) ++CNT;
                    if (CNT > 5)
                    {
                        CNT2++;
                    }
                    if (CNT2 > 10) return true;
                    string[] str1 = CurrentStateMaster.Split(',');
                    string[] str2 = CurrenStateSlave.Split(',');

                    tmpObj1 = "[Slave] : " + CurrenStateSlave + "  " + (i).ToString();
                    tmpObj2 = "[Mater] : " + CurrentStateMaster + "  " + (i).ToString();

                    
                    //MyControl.SendPercentage(0, MasterFlag ? tmpObj1 : tmpObj2);
                    //Console.WriteLine("[Slave] : " + CurrenStateSlave + "  " + CNT.ToString() + "  " + CNT2.ToString());
                    //Console.WriteLine("[Mater] : " + CurrentStateMaster + "  " + CNT.ToString() + "  " + CNT2.ToString());
                }

                try
                {
                    System.Threading.Thread.Sleep(2000);
                }
                catch
                {
                }
                
            }
            return true;
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

        static public string GetIP()
        {
            IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
            //IPAddress[] addr = ;
            return ipEntry.AddressList.Last().ToString();
        }
    }
}
