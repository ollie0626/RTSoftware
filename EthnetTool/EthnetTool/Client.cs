using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Threading;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.IO;


namespace EthnetTool
{
    public class Client
    {
        public string msg;
        public string file_name;
        private TcpClient m_tcpClient;
        public byte[] trans_buffer;
        public byte[] receiver_buffer;

        public void ConnectToServer(string IP, int port)
        {
            //m_tcpClient = new TcpClient(IP, port);
            Console.WriteLine("Host IP:" + IP);
            Console.WriteLine("Connecting to host .... ");
        }


        public void WaitTCPData(string IP, int port, string path)
        {
            try
            {
                m_tcpClient = new TcpClient(IP, port);
                while (true)
                {
                    byte[] btDatas = new byte[4];
                    NetworkStream stream = m_tcpClient.GetStream();
                    stream.Flush();
                    int i = stream.Read(btDatas, 0, btDatas.Length);
                    if (i != 0)
                    {
                        int size = btDatas[0] | btDatas[1] << 8 | btDatas[2] << 16 | btDatas[3] << 24;
                        Console.WriteLine("File Size 0x{0:X}", size);

                        stream.Read(btDatas, 0, btDatas.Length);
                        size = btDatas[0] | btDatas[1] << 8 | btDatas[2] << 16 | btDatas[3] << 24;
                        Console.WriteLine("File Name Size 0x{0:X}", size);

                        btDatas = new byte[size];
                        i = stream.Read(btDatas, 0, btDatas.Length);
                        string sData = Encoding.ASCII.GetString(btDatas, 0, i);
                        Console.WriteLine("File Name " + sData);


                        Array.Resize(ref receiver_buffer, size);
                        stream.Read(receiver_buffer, 0, size);

                        Console.WriteLine("Write File ok!!!! ");
                        FileProcess.WriteFile(receiver_buffer, path, sData);
                        Thread.Sleep(100);
                    }
                }

            }
            catch (SocketException ex)
            {
                Console.WriteLine("SocketException: {0}", ex);
            }
        }


        public void SendTCPData()
        {
            if (m_tcpClient.Connected)
            {
                int size = trans_buffer.Length;
                byte[] btDatas = new byte[4];
                NetworkStream stream = m_tcpClient.GetStream();
                btDatas[0] = (byte)(size & 0xff);
                btDatas[1] = (byte)((size & 0xff00) >> 8);
                btDatas[2] = (byte)((size & 0xff0000) >> 16);
                btDatas[3] = (byte)((size & 0xff000000) >> 24);
                Console.WriteLine("Send File Size " + size.ToString("X"));
                stream.Write(btDatas, 0, btDatas.Length);

                btDatas = Encoding.ASCII.GetBytes(file_name);
                size = btDatas.Length;
                btDatas[0] = (byte)(size & 0xff);
                btDatas[1] = (byte)((size & 0xff00) >> 8);
                btDatas[2] = (byte)((size & 0xff0000) >> 16);
                btDatas[3] = (byte)((size & 0xff000000) >> 24);
                Console.WriteLine("Send File Name Size " + size.ToString("X"));
                stream.Write(btDatas, 0, btDatas.Length);


                Console.WriteLine("Send File Name " + file_name);
                btDatas = Encoding.ASCII.GetBytes(file_name);
                stream.Write(btDatas, 0, btDatas.Length);

                Console.WriteLine("Send File !!!");
                stream.Write(trans_buffer, 0, trans_buffer.Length);
            }
        }


        public void CloseClient()
        {
            m_tcpClient.Close();
        }

        public void CopyFileBuffer(byte[] buf)
        {
            Array.Resize(ref trans_buffer, buf.Length);
            Array.Copy(buf, trans_buffer, buf.Length);
        }

    }
}
