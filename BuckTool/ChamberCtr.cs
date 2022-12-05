using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Sockets;
using System.Net;
using System.IO;

namespace BuckTool
{
    public class ChamberCtr
    {
        public bool Enable = false;
        private const int PORT = 36000;
        public string Role = ""; // Master or Slave
        private const int BUFFER_SIZE = 2048;
        private static readonly byte[] buffer = new byte[BUFFER_SIZE];
        public string Temperature;

        //server
        public static Socket serverSocket;
        public static List<Socket> Socket_List = new List<Socket>();
        public Dictionary<int, string> ClientNowStatus = new Dictionary<int, string>();
        public string MasterStatus = "WAIT"; // RUN | WAIT | STOP

        //client
        private static Socket clientSocket;
        public string SlaveStatus; // READY or OVER

        private delegate void getInfo();
        private delegate void UpdateUI(string sMessage);

        /******************* PUBLIC *******************/
        public int GetClientNumber()
        {
            return Socket_List.Count;
        }
        public bool CheckAllClientStatus(string input)
        {
            foreach (var d in ClientNowStatus)
            {
                if (d.Value != input)
                    return false;
            }
            return true;
        }
        private void GetInfo_Role()
        {
            //if (this.InvokeRequired)
            //{
            //    getInfo del = new getInfo(GetInfo_Role);
            //    this.Invoke(del);
            //}
            //else
            //{
            //    //Role = $[1];
            //}
        }
        public void Init(string temperature)
        {
            //GetInfo_Role();
            switch (Role)
            {
                case "Master":
                    Temperature = temperature;
                    break;

                case "Slave":
                    break;

                default:
                    break;
            }
        }
        public void End()
        {
            switch (Role)
            {
                case "Master":
                    Socket_List.Clear();
                    ClientNowStatus.Clear();
                    break;

                case "Slave":
                    break;

                default:
                    break;
            }
        }

        /******************* MASTER *******************/

        private void ReceiveCallback(IAsyncResult AR)
        {
            Socket current = (Socket)AR.AsyncState;
            int received;

            try
            {
                received = current.EndReceive(AR);
            }
            catch (SocketException)
            {
                ClientNowStatus.Remove(Socket_List.IndexOf(current));
                Console.WriteLine(string.Format("[Client {0}]: forcefully disconnected.", Socket_List.IndexOf(current)));
                current.Close();
                Socket_List.Remove(current);
                return;
            }

            byte[] recBuf = new byte[received];
            Array.Copy(buffer, recBuf, received);
            string text = Encoding.ASCII.GetString(recBuf);

            if (text.ToLower() == "temperature")
            {
                Console.WriteLine(string.Format("[Client {0}]: request temperature.", Socket_List.IndexOf(current)));
                byte[] data = Encoding.ASCII.GetBytes(Temperature);
                current.Send(data);
            }
            else if (text.ToLower() == "status")
            {
                Console.WriteLine(string.Format("[Client {0}]: request status.", Socket_List.IndexOf(current)));
                byte[] data = Encoding.ASCII.GetBytes(MasterStatus);
                current.Send(data);
            }
            else if (text.ToLower() == "ready")
            {
                ClientNowStatus[Socket_List.IndexOf(current)] = "ready";
                Console.WriteLine(string.Format("[Client {0}]: ready. Return status now: {1}", Socket_List.IndexOf(current), MasterStatus));
                byte[] data = Encoding.ASCII.GetBytes(MasterStatus);
                current.Send(data);
            }
            else if (text.ToLower() == "idle")
            {
                ClientNowStatus[Socket_List.IndexOf(current)] = "idle";
                Console.WriteLine(string.Format("[Client {0}]: idle. Return status now: {1}", Socket_List.IndexOf(current), MasterStatus));
                byte[] data = Encoding.ASCII.GetBytes(MasterStatus);
                current.Send(data);
            }
            else if (text.ToLower() == "exit")
            {
                Console.WriteLine(string.Format("[Client {0}]: Disconnected", Socket_List.IndexOf(current)));
                current.Shutdown(SocketShutdown.Both);
                current.Close();
                Socket_List.Remove(current);
                return;
            }
            else
            {
                Console.WriteLine("Invalid request!");
                byte[] data = Encoding.ASCII.GetBytes("Server Response: Invalid request!");
                current.Send(data);
                current.Shutdown(SocketShutdown.Both);
                current.Close();
            }

            current.BeginReceive(buffer, 0, BUFFER_SIZE, SocketFlags.None, ReceiveCallback, current);
        }

        private void AcceptCallback(IAsyncResult AR)
        {
            Socket socket;

            try
            {
                socket = serverSocket.EndAccept(AR);
            }
            catch (ObjectDisposedException) // I cannot seem to avoid this (on exit when properly closing sockets)
            {
                return;
            }

            Socket_List.Add(socket);
            ClientNowStatus.Add(Socket_List.IndexOf(socket), "idle");
            socket.BeginReceive(buffer, 0, BUFFER_SIZE, SocketFlags.None, ReceiveCallback, socket);
            //string str = string.Format("[Client {0}] connected, waiting for request...", Socket_List.IndexOf(socket));
            //UpdateServerBoard(str);
            Console.WriteLine(string.Format("[Client {0}] connected, waiting for request...", Socket_List.IndexOf(socket)));
            serverSocket.BeginAccept(AcceptCallback, null);
        }

        public void SetupServer()
        {
            //UpdateServerBoard("Setting up server...");
            Console.WriteLine("Setting up server...");
            serverSocket.Bind(new IPEndPoint(IPAddress.Any, PORT));
            serverSocket.Listen(0);
            serverSocket.BeginAccept(AcceptCallback, null);
            //UpdateServerBoard("Server setup complete.");
            Console.WriteLine("Server setup complete.");
        }

        public void MasterLisening()
        {
            IPAddress[] ipa = Dns.GetHostAddresses(Dns.GetHostName());
            serverSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            SetupServer();
        }


        public void MasterStop()
        {
            foreach (var c in Socket_List)
            {
                c.Shutdown(SocketShutdown.Both);
                c.Close();
            }
            //serverSocket.Close();
            ClientNowStatus.Clear();
            Socket_List.Clear();
        }

        public void Dispose()
        {
            serverSocket.Dispose();
        }

        ///******************* SLAVE *******************/
        public void ClientConnect(string ip)
        {
            clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

            while (!clientSocket.Connected)
            {
                try
                {
                    clientSocket.Connect(ip, PORT);
                }
                catch (SocketException)
                {
                    Console.WriteLine("Connect Fail!");
                }
            }

             Console.WriteLine("Connected.");
        }

        public string ReceiveResponse()
        {
            var buffer = new byte[2048];
            int received = clientSocket.Receive(buffer, SocketFlags.None);
            if (received == 0)
                return "Fail";

            var data = new byte[received];
            Array.Copy(buffer, data, received);
            string str = Encoding.ASCII.GetString(data);
            return str;
        }

        public void SendString(string str)
        {
            byte[] buffer = Encoding.ASCII.GetBytes(str);
            clientSocket.Send(buffer, 0, buffer.Length, SocketFlags.None);
        }
        public string SendRequest(string str)
        {
            SendString(str);
            string result = ReceiveResponse();
            return result;
        }
        public void Exit()
        {
            SendString("exit");
            clientSocket.Shutdown(SocketShutdown.Both);
            clientSocket.Close();
        }
    }
}
