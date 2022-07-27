using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace EthnetTool
{
    public partial class main : Form
    {
        private Server tcpServer;
        private Client tcpClient;
        

        CancellationTokenSource tokenSource2;
        CancellationToken ct;
        byte[] buf;

        public main()
        {
            InitializeComponent();
            CK_Server.Checked = true;
        }

        private void bt_connect_Click(object sender, EventArgs e)
        {
            int port = Convert.ToInt32(tb_port.Text);
            if (CK_Server.Checked)
            {
                // seriver
                tokenSource2 = new CancellationTokenSource();
                ct = tokenSource2.Token;
                tcpServer = new Server(tb_IPaddr.Text, port, tb_save.Text);
                //tcpServer.CopyFileBuffer(buf);
                tcpServer.msg = "Serve message";
                tcpServer.ListenToConnection();
                Task.Run(() => tcpServer.Listening(), tokenSource2.Token);
            }
            else
            {
                // client
                tokenSource2 = new CancellationTokenSource();
                ct = tokenSource2.Token;
                tcpClient = new Client();
                //tcpClient.CopyFileBuffer(buf);
                tcpClient.msg = "client data";
                tcpClient.ConnectToServer(tb_IPaddr.Text, port, tb_save.Text);
                Task.Run(() => tcpClient.WaitTCPData(tb_IPaddr.Text, port), tokenSource2.Token);
            }
        }

        private void bt_close_Click(object sender, EventArgs e)
        {
            try
            {
                if (CK_Server.Checked)
                {
                    tcpServer.CloseServer();
                    tokenSource2.Cancel();
                }
                else
                {
                    tcpClient.CloseClient();
                }
            }
            catch
            {

            }

        }

        private void CK_Server_CheckedChanged(object sender, EventArgs e)
        {
            if (CK_Server.Checked)
                bt_connect.Text = "Create Host Serve";
            else
                bt_connect.Text = "Connect Host";

        }


        private void bt_getfile_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            if(open.ShowDialog() == DialogResult.OK)
            {
                // read file
                buf = FileProcess.ReadFile(open.FileName);

                if(CK_Server.Checked)
                {
                    tcpServer.file_name = Path.GetFileName(open.FileName);
                }
                else
                {
                    tcpClient.file_name = Path.GetFileName(open.FileName);
                }
            }
        }

        private void main_Load(object sender, EventArgs e)
        {
            string hostName = Dns.GetHostName();
            Console.WriteLine("host name:" + hostName);
            IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress[] ip = ipEntry.AddressList;
            for (int i = 0; i < ip.Length; i++)
            {
                //Console.WriteLine("IP Address {0}: {1}", i, ip[i].ToString());
                textBox2.Text += string.Format("IP Address {0}: {1}\r\n", i, ip[i].ToString());
            }
        }

        private void bt_send_Click(object sender, EventArgs e)
        {
            try
            {
                if (CK_Server.Checked)
                {
                    tcpServer.CopyFileBuffer(buf);
                    tcpServer.SendTCPData();
                }
                else
                {
                    tcpClient.CopyFileBuffer(buf);
                    tcpClient.SendTCPData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(CK_Server.Checked)
            {

            }
            else
            {

            }
        }

        private void tb_save_TextChanged(object sender, EventArgs e)
        {
            
            if(CK_Server.Checked)
            {
                if(tcpServer == null) bt_connect_Click(null, null);
                tcpServer.g_path = tb_save.Text;
            }
            else
            {
                if(tcpClient == null) bt_connect_Click(null, null);
                tcpClient.g_path = tb_save.Text;
            }
        }
    }
}
