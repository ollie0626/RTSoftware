using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace InsLibDotNet
{
    public class TekTronixOSC : VisaCommand
    {
        public TekTronixOSC(string Addr)
        {
            LinkingIns(Addr);
        }

        public TekTronixOSC()
        {

        }

        ~TekTronixOSC()
        {
            InsClose();
        }


        public void ConnectOscilloscope(string Addr)
        {
            LinkingIns(Addr);
        }

        public void TekTronixOSC_RST()
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
            

            string scope_cwd = "FILESystem:CWD " + (char)34 + @"C:" + (char)34;
            doCommand(scope_cwd);

            string scope_dest = "SAVEON:FILE:DEST " + (char)34 + @"C:" + (char)34;
            doCommand(scope_dest);

            string scope_imag_on = "SAVEON:IMAG ON";
            doCommand(scope_imag_on);

            string scope_save_image = "SAVE:IMAGE " + (char)34 + @"C:\tmp.png" + (char)34;
            doCommand(scope_save_image);

            string scope_readfile = "FILESYSTEM:READFILE " + (char)34 + @"C:\tmp.png" + (char)34;
            doCommand(scope_readfile);

#if DEBUG
            Console.WriteLine("scope save path " + buf);
            Console.WriteLine(scope_cwd);
            Console.WriteLine(scope_dest);
            Console.WriteLine(scope_imag_on);
            Console.WriteLine(scope_save_image);
            Console.WriteLine(scope_readfile);
#endif

            System.Threading.Thread.Sleep(2000);
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


        public void MeasPeriod(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS PERIOD";
            doCommand(cmd);
        }


        public void MeasHigh(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS HIGH";
            doCommand(cmd);
        }

        public void MeasLOW(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS LOW";
            doCommand(cmd);
        }

        public void MeasBase(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS BASE";
            doCommand(cmd);
        }

        public void MeasAmp(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS AMPlITUDE";
            doCommand(cmd);
        }

        public void MeasMax(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS MAXIMUM";
            doCommand(cmd);
        }

        public void MeasMean(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS MEAN";
            doCommand(cmd);
        }

        public void MeasMin(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS MINIMUM";
            doCommand(cmd);
        }

        public void MeasFreq(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS FREQUENCY";
            doCommand(cmd);
        }

        public void MeasFallTime(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS FALLTIME";
            doCommand(cmd);
        }

        public void MeasNDuty(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS NDUty";
            doCommand(cmd);
        }

        public void MeasPDuty(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS PDUTY";
            doCommand(cmd);
        }

        public void MeasPK2PK(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS PK2Pk";
            doCommand(cmd);
        }


        public void MeasPOverShoot(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS POVERSHOOT";
            doCommand(cmd);
        }

        public void MeasTop(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS TOP";
            doCommand(cmd);
        }

        public void MeasNOverShoot(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS NOVERSHOOT";
            doCommand(cmd);
        }

        public void MeasRiseTime(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS RISETIME";
            doCommand(cmd);
        }


        public void MeasRiseSlewRate(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS RISESLEWRATE";
            doCommand(cmd);
        }

        public void MeasFallSlewRate(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS FALLSLEWRATE";
            doCommand(cmd);
        }

    }
}





//visa32.viOpenDefaultRM(out Rm);
//visa32.viParseRsrc(Rm, "USB0::0x0699::0x0522::C013665::INSTR", ref intfType, ref intfNum);
//visa32.viOpen(Rm, "USB0::0x0699::0x0522::C013665::INSTR", 0, 0, out vi);
//visa32.viPrintf(vi, "*IDN?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//visa32.viPrintf(vi, "*IDN?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//visa32.viPrintf(vi, "FILESystem:CWD " + "\"C:\"\r\n");
//visa32.viPrintf(vi, "FILESystem:CWD?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//str = "SAVEON:FILE:DEST " + "\"C:/temp/\"\r\n";
//visa32.viPrintf(vi, str);

//str = "SAVEON:IMAG ON\r\n";
//visa32.viPrintf(vi, str);
//System.Threading.Thread.Sleep(500);
//str = "SAVE:IMAGE " + "\"C:/temp/Tek_wave.png\"\r\n";
//visa32.viPrintf(vi, str);
//System.Threading.Thread.Sleep(500);
//string cmd = "FILESYSTEM:READFILE " + "\"C:/temp/Tek_wave.png\"\r\n";
//visa32.viPrintf(vi, cmd);
//System.Threading.Thread.Sleep(1000);
//int dataLen = 500000;
//byte[] bytRead = new byte[dataLen];
//visa32.viBufRead(vi, bytRead, dataLen, out count_out);
//FileStream fStream = File.Open(@"D:\\Tek000.png", FileMode.Create);
//fStream.Write(bytRead, 0, bytRead.Length);
//System.Threading.Thread.Sleep(1000);
//fStream.Close();
//fStream.Dispose();

