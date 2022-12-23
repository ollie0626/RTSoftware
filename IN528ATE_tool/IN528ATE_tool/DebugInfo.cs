using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Windows.Forms;

namespace IN528ATE_tool
{
    public static class DebugInfo
    {
        private static int idx;
        private static FileStream stream_wr;

        public static void CreateFile()
        {
            string app_path = Application.StartupPath + "\\IN528Tool_DebugInfo.txt";
            string info = string.Format("[{0}] {1:MM/dd hh:mm:ss} Create Debug File \r\n", idx++, DateTime.Now);
            byte[] str_buf = Encoding.ASCII.GetBytes(info);
            stream_wr = File.Open(app_path, FileMode.OpenOrCreate);
            stream_wr.Write(str_buf, 0, str_buf.Length);
        }

        public static void Write(string info)
        {
            string cmd = string.Format("[{0}] {1:MM/dd hh:mm:ss} ", idx++, DateTime.Now) + info + "\r\n";
            byte[] str_buf = Encoding.ASCII.GetBytes(cmd);
            stream_wr.Write(str_buf, 0, str_buf.Length);
        }

    }
}
