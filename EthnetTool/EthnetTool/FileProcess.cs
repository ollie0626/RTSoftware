using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;


namespace EthnetTool
{
    public class FileProcess
    {
        private static FileStream swr;
        private static FileStream srd;

        public static void WriteFile(byte[] buffer, string path, string file_name)
        {
            swr = new FileStream(path + @"\" + file_name, FileMode.OpenOrCreate);
            swr.Write(buffer, 0, buffer.Length);
            swr.Close();
        }

        public static byte[] ReadFile(string path)
        {
            srd = new FileStream(path, FileMode.Open, FileAccess.Read);
            BinaryReader r = new BinaryReader(srd);
            byte[] buffer = r.ReadBytes((int)r.BaseStream.Length);
            srd.Close();
            return buffer;
        }

    }
}
