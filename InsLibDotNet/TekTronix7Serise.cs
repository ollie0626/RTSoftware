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

        string CH1 = "CH1";
        string CH2 = "CH2";
        string CH3 = "CH3";
        string CH4 = "CH4";

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

        public void SetRun()
        {
            string cmd = "ACQuire:STATE RUN";
            DoCommand(cmd);
        }

        public void SetStop()
        {
            string cmd = "ACQuire:STATE STOP";
            DoCommand(cmd); 
        }

        public void SetSingle()
        {
            string cmd = "ACQuire:STOPAfter SEQuence";
            DoCommand(cmd);
        }

        public void SetClear()
        {
            string cmd = "DISplay:PERSistence:RESET";
            DoCommand(cmd);
        }

        public void SetTriggerRise()
        {
            string cmd = "TRIGger:A:EDGE:SLOpe RISE";
            DoCommand(cmd);
        }

        public void SetTriggerFall()
        {
            string cmd = "TRIGger:A:EDGE:SLOpe FALL";
            DoCommand(cmd);
        }

        public void SetTriggerMode(bool _isAuto = true)
        {
            string cmd = "TRIGger:A:MODe " + (_isAuto ? "AUTO" : "NORMal");
            DoCommand(cmd);
        }

        public void SetTrigger_50Percent()
        {
            string cmd = "TRIGger:A SETLevel";
            DoCommand(cmd);
        }


        public void SetTriggerSource(int ch)
        {
            string cmd = "TRIGger:A:EDGE:SOUrce CH" + ch.ToString();
            DoCommand(cmd);
        }

        public void SetTriggerLevel(double level)
        {
            string cmd = string.Format("TRIGger:A:LEVel {0}", level);
            DoCommand(cmd);
        }

        public void CHx_On(int i)
        {
            string cmd = string.Format("SELect:CH{0} ON", i);
            DoCommand(cmd);
        }

        public void CHx_Off(int i)
        {
            string cmd = string.Format("SELect:CH{0} OFF", i);
            DoCommand(cmd);
        }

        public void PersistenceEnable()
        {
            string cmd = "DISplay:PERSistence INFPersist";
            DoCommand(cmd);
        }

        public void PersistenceDisable()
        {
            string cmd = "DISplay:PERSistence OFF";
            DoCommand(cmd);
        }

        public void CHx_BWlimitOn(int num)
        {
            string cmd = string.Format("CH{0}:BANdwidth TWEnty", num);
            DoCommand(cmd);
        }

        public void CHx_BWlimitOff(int num)
        {
            string cmd = string.Format("CH{0}:BANdwidth FULl", num);
            DoCommand(cmd);
        }

        public void CHx_ACEnable(int num)
        {
            string cmd = string.Format("CH{0}:COUPling AC", num);
            DoCommand(cmd);
        }

        public void CHx_DCEnable(int num)
        {
            string cmd = string.Format("CH{0}:COUPling DC", num);
            DoCommand(cmd);
        }

        public void CHx_Level(int num, double level)
        {
            string cmd = string.Format("CH{0}:SCAle {1}", num, level);
            DoCommand(cmd);
        }

        public void CHx_Position(int num, double pos)
        {
            string cmd = string.Format("CH{0}:POSition {1}", num, pos);
            DoCommand(cmd);
        }

        public void SetTimeScale(double time)
        {
            string cmd = string.Format("HORizontal:SCAle {0}", time);
            DoCommand(cmd);
        }

        public void SetTimeBasePosition(double pos)
        {
            string cmd = string.Format("HORizontal:POSition {0}", pos);
            DoCommand(cmd);
        }

        public void SetZoomFunc(bool en)
        {
            string cmd = "ZOOM:STATE " + (en ? "ON" : "OFF");
            DoCommand(cmd);
        }

        /// <summary>
        /// SetZoomSize
        /// Tektronix 7Series Zoom In function
        /// </summary>
        /// <param name="size"></param>
        /// ZOOm:GRAticule:SIZE {50|80|100}
        public void SetZoomSize(int size)
        {
            string cmd = string.Format("ZOOm:GRAticule:SIZE {0}", size);
            DoCommand(cmd);
        }

        public void SetZoomInPos(int pos)
        {
            string cmd = string.Format("ZOOm:HORizontal:POSition {0}", pos);
            DoCommand(cmd);
        }



        /*
            MEASUrement:MEAS<x>:TYPe {AMPlitude|AREa|
            BURst|CARea|CMEan|CRMs|DELay|DISTDUty|
            EXTINCTDB|EXTINCTPCT|EXTINCTRATIO|EYEHeight|
            EYEWIdth|FALL|FREQuency|HIGH|HITs|LOW|
            MAXimum|MEAN|MEDian|MINImum|NCROss|NDUty|
            NOVershoot|NWIdth|PBASe|PCROss|PCTCROss|PDUty|
            PEAKHits|PERIod|PHAse|PK2Pk|PKPKJitter|
            PKPKNoise|POVershoot|PTOP|PWIdth|QFACtor|
            RISe|RMS|RMSJitter|RMSNoise|SIGMA1|SIGMA2|
            SIGMA3|SIXSigmajit|SNRatio|STDdev|UNDEFINED| WAVEFORMS}         
         */
        public void SetMeasureSource(int ch, int meas, string type)
        {
            string cmd = "";
            cmd = string.Format("MEASUrement:MEAS{1}:SOUrce{0}", meas, ch);
            DoCommand(cmd);

            cmd = string.Format("MEASUrement:MEAS{0}:TYPe {1}", meas, type);
            DoCommand(cmd);
        }

        public double MeasureMean(int num)
        {
            string cmd = "";
            cmd = string.Format("MEASUrement:MEAS{0}:MEAN?", num);
            double res = doQueryNumber(cmd);
            return res;
        }

        public double MeasureMin(int num)
        {
            string cmd = "";
            cmd = string.Format("MEASUrement:MEAS{0}:MIN?", num);
            double res = doQueryNumber(cmd);
            return res;
        }

        public double MeasureMax(int num)
        {
            string cmd = "";
            cmd = string.Format("MEASUrement:MEAS{0}:MAX?", num);
            double res = doQueryNumber(cmd);
            return res;
        }

    }
}
