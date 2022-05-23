using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace EthnetTool
{
    public class Server
    {
        public string msg;
        public string file_name;
        public byte[] trans_buffer;
        public byte[] receiver_buffer;
        private TcpListener m_tcpListener;
        private TcpClient m_client;
        NetworkStream m_stream;

        public Server(string IP, int port)
        {
            // create local IPEndpoint object
            IPEndPoint ipe = new IPEndPoint(IPAddress.Parse(IP), port);
            // create tcpListener
            m_tcpListener = new TcpListener(ipe);
        }

        public void ListenToConnection()
        {
            // start listener port
            m_tcpListener.Start();
            Console.WriteLine("wait for client connecting .... ");
        }


        public void Listening(string path)
        {
            try
            {
                Console.WriteLine("Waiting for connection ... ");

                m_client = m_tcpListener.AcceptTcpClient();
                m_stream = m_client.GetStream();
                WaitforData(path);
            }
            catch (SocketException ex)
            {
                Console.WriteLine("SocketException: {0}", ex);
            }
        }


        public void WaitforData(string path)
        {
            while(true)
            {
                if (m_client.Connected)
                {
                    byte[] btDatas = new byte[4];
                    int i = m_stream.Read(btDatas, 0, btDatas.Length);
                    if(i != 0)
                    {
                        int size = btDatas[0] | btDatas[1] << 8 | btDatas[2] << 16 | btDatas[3] << 24;
                        Console.WriteLine("File Size 0x{0:X}", size);
                        Array.Resize(ref receiver_buffer, size);

                        //m_stream.Read(btDatas, 0, btDatas.Length);
                        //size = btDatas[0] | btDatas[1] << 8 | btDatas[2] << 16 | btDatas[3] << 24;
                        //Console.WriteLine("File Name Size 0x{0:X}", size);

                        //btDatas = new byte[size];
                        //i = m_stream.Read(btDatas, 0, btDatas.Length);
                        //string sData = Encoding.ASCII.GetString(btDatas, 0, i);
                        //Console.WriteLine("File Name " + sData);

                        Console.WriteLine("Write File ok!!!! ");
                        m_stream.Read(receiver_buffer, 0, size);
                        FileProcess.WriteFile(receiver_buffer, path, "temp");
                    }
                }
            }
        }


        public void SendTCPData()
        {
            if(m_client.Connected)
            {
                //NetworkStream stream = m_client.GetStream();
                int size = trans_buffer.Length;
                byte[] btDatas = new byte[4];
                btDatas[0] = (byte)(size & 0xff);
                btDatas[1] = (byte)((size & 0xff00) >> 8);
                btDatas[2] = (byte)((size & 0xff0000) >> 16);
                btDatas[3] = (byte)((size & 0xff000000) >> 24);
                Console.WriteLine("Send File Size " + size.ToString("X"));
                m_stream.Write(btDatas, 0, btDatas.Length);


                //btDatas = Encoding.ASCII.GetBytes(file_name);
                //size = btDatas.Length;
                //btDatas[0] = (byte)(size & 0xff);
                //btDatas[1] = (byte)((size & 0xff00) >> 8);
                //btDatas[2] = (byte)((size & 0xff0000) >> 16);
                //btDatas[3] = (byte)((size & 0xff000000) >> 24);
                //Console.WriteLine("Send File Name Size " + size.ToString("X"));
                //m_stream.Write(btDatas, 0, btDatas.Length);

                //Console.WriteLine("Send File Name " + file_name);
                //btDatas = Encoding.ASCII.GetBytes(file_name);
                //m_stream.Write(btDatas, 0, btDatas.Length);

                Console.WriteLine("Send File ");
                m_stream.Write(trans_buffer, 0, trans_buffer.Length);
            }
        }




        public void CopyFileBuffer(byte[] buf)
        {
            Array.Resize(ref trans_buffer, buf.Length);
            Array.Copy(buf, trans_buffer, buf.Length);
        }


        public void CloseServer()
        {
            m_tcpListener.Stop();
        }

    }
}
