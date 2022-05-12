using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace InsLibDotNet
{
    public class TekTronix7Serise : VisaCommand
    {
        public TekTronix7Serise(string Addr)
        {
            LinkingIns(Addr);
        }

        public TekTronix7Serise()
        {

        }

        ~TekTronix7Serise()
        {
            InsClose();
        }

        public void ConnectOscilloscope(string Addr)
        {
            LinkingIns(Addr);
        }

        public void TekTronix7Serise_RST()
        {
            doCommand("*RST");
        }

        public void DoCommand(string cmd)
        {
            doCommand(cmd);
        }
        public string doQuery(string cmd)
        {
            return doQueryString(cmd);
        }

        public string doRead()
        {
            return doReadString();
        }

        public void SaveWaveform(string path, string filename)
        {
            string buf = path.Substring(path.Length - 1, 1) == @"/" ? path.Substring(0, path.Length - 1) : path;
            buf = buf + @"/" + filename + ".png";

            string waveFmt = "EXP:FORM PNG";
            doCommand(waveFmt);

            string portFile = "HARDCopy:PORT FILE";
            doCommand(portFile);

            string hard_Cp_FileName = "HARDCopy:FILEName " + @"""C:\scope.png"""; /* scope can't save C:\ directly */
            doCommand(hard_Cp_FileName);

            string hard_Cp_Start = "HARDCopy STARt";
            doCommand(hard_Cp_Start);

            string FileSystem_ReadFile = "FILESystem:READFile " + @"""C:\scope.png""";
            doCommand(FileSystem_ReadFile);

#if DEBUG
            Console.WriteLine(buf);
            Console.WriteLine(waveFmt);
            Console.WriteLine(portFile);
            Console.WriteLine(hard_Cp_FileName);
            Console.WriteLine(hard_Cp_Start);
            Console.WriteLine(FileSystem_ReadFile);
#endif

            System.Threading.Thread.Sleep(1000);
            int count_out = 0;
            int len = 500000;
            byte[] bytRead = new byte[len];
            visa32.viBufRead(device, bytRead, len, out count_out);
            FileStream fStream = File.Open(buf, FileMode.Create);
            fStream.Write(bytRead, 0, bytRead.Length);
            System.Threading.Thread.Sleep(500);
            fStream.Close();
            fStream.Dispose();

            visa32.viFlush(device, visa32.VI_READ_BUF);
            visa32.viFlush(device, visa32.VI_WRITE_BUF);
        }




    }
}
