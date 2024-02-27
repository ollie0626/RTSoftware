using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class TekOSC : VisaCommand
    {
        public void ConnectOscilloscope(string Address)
        {
            LinkingIns(Address);
        }
        public void GetCHx_Waveform(int channel)
        {
           /* TekScope.Write("WFMOutpre:PT_Off?");
            temp = TekScope.ReadString().Trim();
            pt_off = Int32.Parse(temp);
            TekScope.Write("WFMOutpre:XINcr?");
            temp = TekScope.ReadString().Trim();
            xinc = Single.Parse(temp);
            TekScope.Write("WFMOutpre:XZEro?");
            temp = TekScope.ReadString().Trim();
            xzero = Single.Parse(temp);
            TekScope.Write("WFMOutpre:YMUlt?");
            temp = TekScope.ReadString().Trim();
            ymult = Single.Parse(temp);
            TekScope.Write("WFMOutpre:YOFf?");
            temp = TekScope.ReadString().Trim().TrimEnd('0').TrimEnd('.');
            yoff = Single.Parse(temp);
            TekScope.Write("WFMOutpre:YZEro?");
            temp = TekScope.ReadString().Trim();
            yzero = Single.Parse(temp);
            */
            // Turn on curve streaming
           // TekScope.Write("CURVEStream?");


            /* mbSession.Write("DATA:SOURCE CH1");
             mbSession.Write("DATa:ENCdg RPB");
             mbSession.Write("WFMOutpre:BYT_Nr 1");
             mbSession.Write("DATA:START 1");
             mbSession.Write("HEADER OFF");
             mbSession.Write("curve?");
             result = mbSession.Query("HORizontal:RECOrdlength?");
             RecLength = Int32.Parse(result);
             mbSession.Write("DATA:STOP " + RecLength);
             RecLength += 25;
             mbSession.DefaultBufferSize = RecLength;
             string gogoCMD = ":MEASure:VAVE?";
             doCommand(CHx);
             buf = doQueryNumber(gogoCMD);
             return buf;*/
        }

        public void GetPicture(int channel)
        {
            //Scope.Write("ACQuire:STATE STOP");
            //Scope.Write("HardCopy:FormatException bmp");
            //Scope.Write("Hardcopy:Layout Portrait");
            //Scope.Write("HARDCopy:INKSaver OFF");
            //Scope.Write("Hardcopy:port USB");
            //Scope.Write("Hardcopy Start");
        }

        public void SaveWaveform(string path, string filename)
        {
            if (device == 0) return;
            Console.WriteLine("Path " + path);
            string buf = path.Substring(path.Length - 1, 1) == @"/" ? path.Substring(0, path.Length - 1) : path;
            buf = buf + @"/" + filename + ".png";

            string waveFmt = "EXP:FORM PNG";
            doCommand(waveFmt);

            string portFile = "HARDCopy:PORT FILE";
            doCommand(portFile);

            string hard_Cp_FileName = "HARDCopy:FILEName " + @"""C:\temp\scope.png""";
            doCommand(hard_Cp_FileName);

            string hard_Cp_Start = "HARDCopy STARt";
            doCommand(hard_Cp_Start);

            string FileSystem_ReadFile = "FILESystem:READFile " + @"""C:\temp\scope.png""";
            doCommand(FileSystem_ReadFile);

#if DEBUG
            Console.WriteLine(buf);
            Console.WriteLine(waveFmt);
            Console.WriteLine(portFile);
            Console.WriteLine(hard_Cp_FileName);
            Console.WriteLine(hard_Cp_Start);
            Console.WriteLine(FileSystem_ReadFile);
#endif

            int count_out = 0;
            int len = 500000;
            byte[] bytRead = new byte[len];
            bool flag = false;
            for(int i = 0; i < 5; ++i)
            {
                System.Threading.Thread.Sleep(2000);
                Console.WriteLine(i);
                visa32.viBufRead(device, bytRead, len, out count_out);
                for(int j = 0; j < len; ++j)
                {
                    if(bytRead[j] > 0)
                    {
                        flag = true;
                        break;
                    }
                }
                if(flag) break;
            }
            Console.WriteLine("GetFin");
            if (flag)
            {
                FileStream fStream = File.Open(buf, FileMode.Create);
                fStream.Write(bytRead, 0, bytRead.Length);
                //System.Threading.Thread.Sleep(500);
                fStream.Close();
                fStream.Dispose();
            }
            Console.WriteLine("WriteFin");
            visa32.viFlush(device, visa32.VI_READ_BUF);
            visa32.viFlush(device, visa32.VI_WRITE_BUF);
        }

        public void ExeRun(bool IsRun)
        {
            if (device == 0) return;
            if (!IsRun)
            {
                doCommand("ACQuire:STATE OFF");
            }
            else
            {
                doCommand("ACQuire:STATE ON");
            }
        }

        public void ExeType(bool IsSingle)
        {
            if (device == 0) return;
            if (!IsSingle)
            {
                doCommand("ACQuire:STOPAfter RUNSTop");
            }
            else
            {
                doCommand("ACQuire:STOPAfter SEQuence");
            }
        }

    }
}
