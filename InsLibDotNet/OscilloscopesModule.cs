using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace InsLibDotNet
{
    public class OscilloscopesModule : VisaCommand
    {
        /*
         * 
         * 0: Tektronix 7 series
         * 1: Agilent 9 series
         * 
         */
        public int osc_sel;

        public OscilloscopesModule()
        {

        }

        public OscilloscopesModule(string Addr)
        {
            LinkingIns(Addr);
            if (doQueryIDN().Split(',')[0].IndexOf("TEKTRONIX") != -1)
            {
                osc_sel = 0;
            }
            else
            {
                osc_sel = 1;
            }
        }

        ~OscilloscopesModule()
        {
            InsClose();
        }

        public void SetLevelAboutGnd()
        {
            switch (osc_sel)
            {
                case 0: break;
                case 1:
                    doCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
                    break;
            }
        }

        public void SetLevelAboutCenter()
        {
            switch (osc_sel)
            {
                case 0: break;
                case 1:
                    doCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandCenter\"");
                    break;
            }
        }

        public void SetRST()
        {
            doCommand("*RST");
        }

        public void SetStop()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("ACQuire:STATE STOP");
                    break;
                case 1:
                    doCommand(":STOP");
                    break;
            }
        }

        public void SetPERSistence()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("DISplay:PERSistence INFPersist");
                    break;
                case 1:
                    break;
            }
        }

        public void SetDPXOn()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("FASTAcq:STATE ON");
                    break;
                case 1:
                    break;
            }
        }

        public void SetDPXOff()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("FASTAcq:STATE OFF");
                    break;
                case 1:
                    break;
            }
        }

        public void SetPERSistenceOff()
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand("DISplay:PERSistence OFF");
                    break;
                case 1:
                    break;
            }
        }


        public void SetClear()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("DISplay:PERSistence:RESET");
                    break;
                case 1:
                    doCommand(":CDISplay");
                    break;
            }
        }

        public void SetSingle()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand(":SINGle");
                    break;
                case 1:
                    doCommand("ACQuire:STOPAfter SEQuence");
                    break;
            }
        }

        public void SetAutoTrigger()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("TRIGger:A:MODe AUTO");
                    break;
                case 1:
                    doCommand(":TRIGger:SWEep AUTO");
                    break;
            }
        }

        public void SetNormalTrigger()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("TRIGger:A:MODe NORMal");
                    break;
                case 1:
                    doCommand(":TRIGger:SWEep TRIGgered");
                    break;
            }
        }

        public void SetTriggerRise()
        {

            switch (osc_sel)
            {
                case 0:
                    doCommand("TRIGger:A:EDGE:SLOpe RISE");
                    break;
                case 1:
                    doCommand(":TRIGger:EDGE:SLOPe POSitive");
                    break;
            }

        }

        public void SetTriggerFall()
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand("TRIGger:A:EDGE:SLOpe FALL");
                    break;
                case 1:
                    doCommand(":TRIGger:EDGE:SLOPe NEGative");
                    break;
            }
        }

        public void CHx_BWLimitOn(int ch)
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand(string.Format(":CH{0}:BWLimit 20e6", ch));
                    break;
                case 1:
                    doCommand(string.Format("CH{0}:BANdwidth TWEnty", ch));
                    break;
            }
        }

        public void SetTriggerLevel(int ch, double level)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("TRIGger: A: EDGE: SOUrce CH{0}", ch));
                    doCommand(string.Format("TRIGger:A:LEVel {0}", level));
                    break;
                case 1:
                    doCommand(string.Format(":TRIGger:LEVel CHANnel{0}, {1}", ch, level));
                    break;
            }
        }

        public void CHx_On(int ch)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("SELect:CH{0} ON", ch));
                    break;
                case 1:
                    doCommand(string.Format(":CHANnel{0}:DISPLAY ON", ch));
                    break;
            }
        }

        public void CHx_Off(int ch)
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand(string.Format("SELect:CH{0} OFF", ch));
                    break;
                case 1:
                    doCommand(string.Format(":CHANnel{0}:DISPLAY OFF", ch));
                    break;
            }
        }

        public void CHx_ACoupling(int ch)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("CH{0}:COUPling AC", ch));
                    break;
                case 1:
                    doCommand(string.Format(":CHANnel{0}:INPut AC", ch));
                    break;
            }
        }

        public void CHx_DCoupling(int ch)
        {
            switch (osc_sel)
            {
                case 0:
                    doCommand(string.Format("CH{0}:COUPling DC", ch));
                    break;
                case 1:
                    doCommand(string.Format(":CHANnel{0}:INPut DC", ch));
                    break;
            }
        }

        public void SetTimeScale(double time)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("HORizontal:SCAle {0}", time));
                    break;
                case 1:
                    doCommand(string.Format(":TIMEBASE:SCALE {0}", time));
                    break;
            }
        }

        public void SetTimeBasePosition(int pos)
        {
            double timeScale = 0;
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("HORizontal:POSition {0}", pos));
                    break;
                case 1:
                    timeScale = doQueryNumber(":TIMEBASE:SCALE?");
                    doCommand(string.Format(":TIMEBASE:POSITION {0}", pos * timeScale));
                    break;  
            }
        }

        public void SetMeasurePercent(double hi, double mid, double lo)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("MEASUrement:IMMed:REFLevel:METHod PERCent"));
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:HIGH {0}", hi));
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:MID {0}", mid));
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:LOW {0}", lo));
                    break;
                case 1:
                    doCommand(string.Format(":MEASure:THResholds:METHod ALL,PERCent"));
                    doCommand(string.Format(":MEASure:THResholds:GEN:PERCent ALL,{0},{1},{2}", hi, mid, lo));
                    break;
            }
        }

        public void SetMeasureAbsolute(double hi, double mid, double lo)
        {
            switch(osc_sel)
            {
                case 0:
                    // need to test
                    doCommand("MEASUrement:IMMed:REFLevel:METHod ABSolute");
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:HIGH {0}", hi));
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:MID {0}", mid));
                    doCommand(string.Format("MEASUrement:REFLevel:PERCent:LOW {0}", lo));
                    break;
                case 1:
                    doCommand(":MEASure:THResholds:GENeral:METHod ALL,ABSolute");
                    doCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute ALL,{0},{1},{2}", hi, mid, lo));
                    break;
            }
        }

        public void SetDelayTime(int meas, int ch1, int ch2, bool _first_edge_rising = true, bool _second_edge_rising = true)
        {
            switch(osc_sel)
            {
                case 0:
                    doCommand(string.Format("MEASUrement:MEAS{0}:TYPe DELay", meas));
                    doCommand(string.Format("MEASUrement:MEAS{0}:SOUrce1 CH{1}", meas, ch1));
                    doCommand(string.Format("MEASUrement:MEAS{0}:SOUrce2 CH{1}", meas, ch2));
                    doCommand(string.Format("MEASUrement:MEAS{0}:DELay:EDGE1 {1};EDGE2 {2}",
                            meas, _first_edge_rising ? "RISe" : "FALL", _second_edge_rising ? "RISe" : "FALL"));
                    doCommand(string.Format("MEASUrement:MEAS{0}:STATE ON", meas));
                    break;
                case 1:
                    // measure first edge
                    doCommand(string.Format(":MEASure:DELTatime CHANnel{0}, CHANnel{1}", ch1, ch2));
                    doCommand(string.Format(":MEASure:DELTatime:DEFine {0}, 1, MID, {1}, 1, LOWer",
                        _first_edge_rising ? "RISing" : "Falling",
                        _second_edge_rising ? "RISing" : "Falling"
                        ));
                    break;
            }
        }

        public void SaveWaveform(string path, string file)
        {
            string buf = path.Substring(path.Length - 1, 1) == @"\" ? path.Substring(0, path.Length - 1) : path;
            buf = buf + @"\" + file + ".png";
            FileStream fStream;
            switch (osc_sel)
            {
                case 0:
                    string waveFmt = "EXP:FORM PNG";
                    doCommand(waveFmt);
                    string portFile = "HARDCopy:PORT FILE";
                    doCommand(portFile);
                    string hard_Cp_FileName = "HARDCopy:FILEName " + @"""C:\\TekScope\\scope.png"""; /* scope can't save C:\ directly */
                    doCommand(hard_Cp_FileName);
                    string hard_Cp_Start = "HARDCopy STARt";
                    doCommand(hard_Cp_Start);
                    string FileSystem_ReadFile = "FILESystem:READFile " + @"""C:\\TekScope\\scope.png""";
                    doCommand(FileSystem_ReadFile);
#if DEBUG
                    //Console.WriteLine(buf);
                    //Console.WriteLine(waveFmt);
                    //Console.WriteLine(portFile);
                    //Console.WriteLine(hard_Cp_FileName);
                    //Console.WriteLine(hard_Cp_Start);
                    //Console.WriteLine(FileSystem_ReadFile);
#endif
                    System.Threading.Thread.Sleep(1000);
                    int count_out = 0;
                    int len = 500000;
                    byte[] bytRead = new byte[len];
                    visa32.viBufRead(device, bytRead, len, out count_out);
                    fStream = File.Open(buf, FileMode.Create);
                    fStream.Write(bytRead, 0, bytRead.Length);
                    System.Threading.Thread.Sleep(500);
                    fStream.Close();
                    fStream.Dispose();

                    visa32.viFlush(device, visa32.VI_READ_BUF);
                    visa32.viFlush(device, visa32.VI_WRITE_BUF);
                    break;
                case 1:
                    fStream = File.Open(buf, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                    string gogoCMD = ":DISPlay:DATA? PNG";
                    doCommand(gogoCMD); System.Threading.Thread.Sleep(2000);
                    byte[] ResultsArray = new byte[300000];
                    IEEEBlock_Bytes(out ResultsArray);
                    fStream.Write(ResultsArray, 0, ResultsArray.Length);
                    fStream.Close();
                    fStream.Dispose();
                    break;
            }
        }

        public void SetMeasureSource(int ch, int meas, string type)
        {
            string cmd = "";
            switch(osc_sel)
            {
                case 0:
                    cmd = string.Format("MEASUrement:MEAS{1}:SOUrce1 CH{0}", ch, meas);
                    doCommand(cmd);

                    cmd = string.Format("MEASUrement:MEAS{0}:TYPe {1}", meas, type);
                    doCommand(cmd);

                    cmd = string.Format("MEASUrement:MEAS{0}:STATE ON", meas);
                    doCommand(cmd);
                    break;
                case 1:
                    cmd = string.Format(":MEASure:{0} CHANnel{1}", type, ch);
                    doCommand(cmd);
                    break;
            }
        }


        public void CHx_Level(int CHx, double level)
        {
            string cmd = "";
            switch(osc_sel)
            {
                case 0:
                    cmd = string.Format("CH{0}:SCAle {1}", CHx, level);
                    break;
                case 1:
                    cmd = ":CHANNEL" + CHx.ToString() + ":SCALe " + level.ToString();
                    break;
            }
            
            doCommand(cmd);
        }

        public void CHx_Offset(int CHx, double offset)
        {
            string cmd = "";
            switch(osc_sel)
            {
                case 0:
                    cmd = string.Format("CH{0}:OFFSet {1}", CHx, offset);
                    break;
                case 1:
                    break;
            }
            doCommand(cmd);
        }

        public void CHx_Position(int CHx, double pos)
        {
            string cmd = "";
            switch(osc_sel)
            { 
                case 0:
                    cmd = string.Format("CH{0}:POSition {1}", CHx, pos);
                    break;
                case 1:
                    cmd = string.Format(":CHANnel{0}:OFFSet {1}", CHx, pos);
                    break;
            }
            doCommand(cmd);
        }


        public string GetStatistics(int sel)
        {
            string res = "";
            string cmd = "";

            switch(sel)
            {
                case 0:
                    cmd = ":MEASure:STATistics CURRent";
                    break;
                case 1:
                    cmd = ":MEASure:STATistics Max";
                    break;
                case 2:
                    cmd = ":MEASure:STATistics Min";
                    break;
                case 3:
                    cmd = ":MEASure:STATistics MEAN";
                    break;
            }
            switch(osc_sel)
            {
                case 0:
                    break;
                case 1:
                    doCommand(cmd);
                    res = doQueryString(":MEASure:RESults?");
                    break;
            }
            return res;
        }

        public void Measure_Clear()
        {
            switch(osc_sel)
            {
                case 0:
                    for(int i = 0; i < 10; i++)
                        doCommand(string.Format("MEASUrement:MEAS{0}:STATE OFF", i));
                    break;
                case 1:
                    doCommand(":MEASure:CLEar");
                    break;
            }
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

        /*  
         *  Tektronix
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

        public void SetMeasureOff(int meas)
        {
            string cmd = string.Format("MEASUrement:MEAS{0}:STATE OFF", meas);
            doCommand(cmd); System.Threading.Thread.Sleep(100);
        }

        public double CHx_Meas_Overshoot(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "POVershoot");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:OVERshoot? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Undershoot(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "NOVershoot");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:OVERshoot? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Max(int ch, int meas = 1)
        {
            double res = 0;
            switch(osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "MAXimum");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VMAX? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Min(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "MINImum");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VMIN? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Mean(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "MEAN");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VAVerage? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_AMP(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "AMPlitude");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VAMPlitude? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Top(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "HIGH");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VTOP? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Base(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "LOW");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VBASE? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Rise(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "RISe");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:RISetime? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Fall(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "FALL");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:FALLtime? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Freq(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "FREQuency");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:FREQuency? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_Period(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "PERIod");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:PERiod? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        public double CHx_Meas_VPP(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "PK2Pk");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    res = doQueryNumber(string.Format(":MEASure:VPP? CHANnel{0}", ch));
                    break;
            }
            return res;
        }

        //PKPKJitter

        public double CHx_Meas_Jitter(int ch, int meas = 1)
        {
            double res = 0;
            switch (osc_sel)
            {
                case 0:
                    SetMeasureSource(ch, meas, "PKPKJitter");
                    res = MeasureMean(meas);
                    break;
                case 1:
                    //res = doQueryNumber(string.Format(":MEASure:VPP? CHANnel{0}", ch));
                    break;
            }
            return res;
        }


    }
}
