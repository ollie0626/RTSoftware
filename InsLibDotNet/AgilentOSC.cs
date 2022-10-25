using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace InsLibDotNet
{
    public class AgilentOSC : VisaCommand
    {
        // 1. override docommand
        // 2. :PDER?
        // 3. :ADER?
        string MEAS_CH1 = "CHANnel1";
        string MEAS_CH2 = "CHANnel2";
        string MEAS_CH3 = "CHANnel3";
        string MEAS_CH4 = "CHANnel4";
        string CH1 = "CHANnel1";
        string CH2 = "CHANnel2";
        string CH3 = "CHANnel3";
        string CH4 = "CHANnel4";

        public AgilentOSC(string Addr)
        {
            LinkingIns(Addr);
        }

        public AgilentOSC()
        {
        }

        ~AgilentOSC()
        {
            InsClose();
        }


        public void ConnectOscilloscope(string Addr)
        {
            LinkingIns(Addr);
        }

        public void AgilentOSC_RST()
        {
            doCommand("*RST");
        }

        public void Measure_Clear()
        {
            doCommand(":MEASure:CLEar");
        }

        public void Root_STOP()
        {
            doCommand(":STOP");
        }

        public void Root_RUN()
        {
            doCommand(":RUN");
        }

        public void Root_Single()
        {
            doCommand(":SINGle");
        }

        public void Root_Clear()
        {
            doCommand(":CDISplay");
        }

        public void AutoTrigger()
        {
            doCommand(":TRIGger:SWEep AUTO");
        }

        public void NormalTrigger()
        {
            doCommand(":TRIGger:SWEep TRIGgered");
        }

        public void SingleTrigger()
        {
            doCommand(":TRIGger:SWEep SINGle");
        }

        public void TimeBasePositionUs(double position)
        {
            doCommand(":TIMEBASE:POSITION " + position.ToString() + "us");
        }

        public void TimeBasePosition(double position)
        {
            doCommand(":TIMEBASE:POSITION " + position.ToString());
        }

        public void TimeBasePositionMs(double position)
        {
            doCommand(":TIMEBASE:POSITION " + position.ToString() + "ms");
        }

        public void TimeScale(double Scale)
        {
            doCommand(":TIMEBASE:SCALE " + Scale.ToString());
        }

        public void TimeScaleMs(double Scale)
        {
            TimeScale(Scale / 1000);
        }

        public void TimeScaleUs(double Scale)
        {
            TimeScaleMs(Scale / 1000);
        }

        public void SaveWaveform(string Path, string FileName)
        {
#if true
            string buf = Path.Substring(Path.Length - 1, 1) == @"\" ? Path.Substring(0, Path.Length - 1) : Path;
            buf = buf + @"\" + FileName + ".png";

            FileStream fStream = File.Open(buf, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);

            string gogoCMD = ":DISPlay:DATA? PNG";
            doCommand(gogoCMD); System.Threading.Thread.Sleep(2000);
            byte[] ResultsArray = new byte[300000];
            IEEEBlock_Bytes(out ResultsArray);
            fStream.Write(ResultsArray, 0, ResultsArray.Length);
            //System.Threading.Thread.Sleep(1000);
            fStream.Close();
            fStream.Dispose();
#endif
        }

#if false
        public void SlewRate_90Range()
        {
            string gogoCMD = ":MEASure:THResholds:GENeral:METHod ALL,PERCent";
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:PERCent ALL,90,50,10";
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:RFALl:METHod ALL,PERCent";
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:RFALl:PERCent ALL,90,50,10";
            doCommand(gogoCMD);
        }
#else

        public void SlewRate_90Range(string CHx)
        {
            //doCommand(CHx);
            double amp = doQueryNumber(":MEASure:VAMPlitude? " + CHx);
            double hi = amp * 0.9;
            double mid = amp * 0.5;
            double lo = amp * 0.1;
            string gogoCMD = ":MEASure:THResholds:GENeral:METHod ALL,ABSolute";
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:GENeral:ABSolute ALL," + hi.ToString() + "," + mid.ToString() + "," + lo.ToString();
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:GENeral:METHod ALL,ABSolute";
            doCommand(gogoCMD);
            gogoCMD = ":MEASure:THResholds:GENeral:ABSolute ALL," + hi.ToString() + "," + mid.ToString() + "," + lo.ToString();
            doCommand(gogoCMD);
        }


        public void SlewRate20_80Range()
        {
            string cmd = ":MEASure:THResholds:METHod ALL,PERCent";
            doCommand(cmd);
            cmd = ":MEASure:THResholds:GEN:PERCent ALL,80,50,20";
            doCommand(cmd);
        }


        public void Meas_ThresholdMethod(int ch, bool isPercent)
        {
            string method = isPercent ? "PERCent" : "ABSolute";
            string cmd = string.Format(":MEASure:THResholds:METHod CHANnel{0}, {1}", ch, method);
            doCommand(cmd);
        }

        public void Meas_Percent(int ch, double hi, double mid, double low)
        {
            string cmd = string.Format(":MEASure:THResholds:PERCent CHANnel{0}, {1}, {2}, {3}", ch, hi, mid, low);
            doCommand(cmd);
        }

        public void Meas_Absolute(int ch, double hi, double mid, double low)
        {
            string cmd = string.Format(":MEASure:THResholds:ABSolute CHANnel{0}, {1}, {2}, {3}", ch, hi, mid, low);
            doCommand(cmd);
        }


#endif


        private void CHx_Input(string CHx, bool is50ohm)
        {
            string cmd = ":" + CHx + ":INPut ";
            if (is50ohm)
                cmd += "DC50";
            else
                cmd += "DC";
            doCommand(cmd);
        }

        private void CHx_50ohm(string CHx)
        {
            CHx_Input(CHx, true);
        }

        private void CHx_1Mohm(string CHx)
        {
            CHx_Input(CHx, false);
        }

        public void CH1_50ohm()
        {
            CHx_50ohm(CH1);
        }

        public void CH2_50ohm()
        {
            CHx_50ohm(CH2);
        }
        public void CH3_50ohm()
        {
            CHx_50ohm(CH3);
        }

        public void CH4_50ohm()
        {
            CHx_50ohm(CH4);
        }

        public void CH1_1Mohm()
        {
            CHx_1Mohm(CH1);
        }
        public void CH2_1Mohm()
        {
            CHx_1Mohm(CH2);
        }
        public void CH3_1Mohm()
        {
            CHx_1Mohm(CH3);
        }
        public void CH4_1Mohm()
        {
            CHx_1Mohm(CH4);
        }

        private void CHx_On(string CHx)
        {
            string cmd = ":" + CHx + ":DISPLAY ON";
            doCommand(cmd);
        }

        public void CH1_On()
        {
            CHx_On(CH1);
        }

        public void CH2_On()
        {
            CHx_On(CH2);
        }

        public void CH3_On()
        {
            CHx_On(CH3);
        }

        public void CH4_On()
        {
            CHx_On(CH4);
        }

        private void CHx_Off(string CHx)
        {
            string cmd = ":" + CHx + ":DISPLAY OFF";
            doCommand(cmd);
        }

        public void CH1_Off()
        {
            CHx_Off(CH1);
        }

        public void CH2_Off()
        {
            CHx_Off(CH2);
        }

        public void CH3_Off()
        {
            CHx_Off(CH3);
        }

        public void CH4_Off()
        {
            CHx_Off(CH4);
        }

        private void CHx_Offset(string CHx, double offset)
        {
            string gogoCMD = CHx + ":OFFSet " + offset.ToString();
            doCommand(gogoCMD);
        }

        public void CH1_Offset(double offset)
        {
            CHx_Offset(CH1, offset);
        }

        public void CH2_Offset(double offset)
        {
            CHx_Offset(CH2, offset);
        }

        public void CH3_Offset(double offset)
        {
            CHx_Offset(CH3, offset);
        }

        public void CH4_Offset(double offset)
        {
            CHx_Offset(CH4, offset);
        }

        private void TriggerSource(string CHx)
        {
            string gogoCMD = ":TRIGger:EDGE:SOURce " + CHx;
            doCommand(gogoCMD);
        }

        public void Trigger_CH1()
        {
            TriggerSource(CH1);
        }

        public void Trigger_CH2()
        {
            TriggerSource(CH2);
        }

        public void Trigger_CH3()
        {
            TriggerSource(CH3);
        }

        public void Trigger_CH4()
        {
            TriggerSource(CH4);
        }

        private void CHx_DCoupling(string CHx)
        {
            string gogoCMD = ":" + CHx + "INPut DC";
            doCommand(gogoCMD);
        }

        public void CH1_DCoupling()
        {
            CHx_DCoupling(CH1);
        }

        public void CH2_DCoupling()
        {
            CHx_DCoupling(CH2);
        }

        public void CH3_DCoupling()
        {
            CHx_DCoupling(CH3);
        }

        public void CH4_DCoupling()
        {
            CHx_DCoupling(CH4);
        }

        private void CHx_ACoupling(string CHx)
        {
            string gogoCMD = ":" + CHx + ":INPut AC";
            doCommand(gogoCMD);
        }

        public void CH1_ACoupling()
        {
            CHx_ACoupling(CH1);
        }

        public void CH2_ACoupling()
        {
            CHx_ACoupling(CH2);
        }

        public void CH3_ACoupling()
        {
            CHx_ACoupling(CH3);
        }

        public void CH4_ACoupling()
        {
            CHx_ACoupling(CH4);
        }

        private void TriggerLevel(string CHx, double level)
        {
            string gogoCMD = ":TRIGger:LEVel " + CHx + "," + level.ToString();
            doCommand(gogoCMD);
        }
        public void TriggerLevel_CH1(double level)
        {
            TriggerLevel(CH1, level);
        }
        public void TriggerLevel_CH2(double level)
        {
            TriggerLevel(CH2, level);
        }
        public void TriggerLevel_CH3(double level)
        {
            TriggerLevel(CH3, level);
        }
        public void TriggerLevel_CH4(double level)
        {
            TriggerLevel(CH4, level);
        }

        public void CHx_BWLimitOn(string CHx)
        {
            //:CHANnel<N>:BWLimit
            string gogoCMD = ":" + CHx + ":BWLimit 20e6";
            doCommand(gogoCMD);
            gogoCMD = ":" + CHx + ":BWLimit ON";
            doCommand(gogoCMD);

        }
        
        public void CH1_BWLimitOn()
        {
            CHx_BWLimitOn(CH1);
        }

        public void CH2_BWLimitOn()
        {
            CHx_BWLimitOn(CH2);
        }

        public void CH3_BWLimitOn()
        {
            CHx_BWLimitOn(CH3);
        }

        public void CH4_BWLimitOn()
        {
            CHx_BWLimitOn(CH4);
        }

        public void CHx_BWLimitOff(string CHx)
        {
            string gogoCMD = ":" + CHx + ":BWLIMIT OFF";
            doCommand(gogoCMD);
        }

        public void CH1_BWLimitOff()
        {
            CHx_BWLimitOff(CH1);
        }

        public void CH2_BWLimitOff()
        {
            CHx_BWLimitOff(CH2);
        }

        public void CH3_BWLimitOff()
        {
            CHx_BWLimitOff(CH3);
        }

        public void CH4_BWLimitOff()
        {
            CHx_BWLimitOff(CH4);
        }


        public void CHx_Level(int CHx, double level)
        {
            string cmd = ":CHANNEL" + CHx.ToString() + ":SCALe " + level.ToString();
            doCommand(cmd);
        }

        public void CH1_Level(double level)
        {
            CHx_Level(1, level);
        }

        public void CH2_Level(double level)
        {
            CHx_Level(2, level);
        }

        public void CH3_Level(double level)
        {
            CHx_Level(3, level);
        }

        public void CH4_Level(double level)
        {
            CHx_Level(4, level);
        }

        // ----------------------------------------------------------------------
        // Measure command
        // ----------------------------------------------------------------------

        public void SetTrigModeTimeOut(bool isNegative, double delay_s)
        {
            doCommand(":TRIGger:MODE Timeout");
            string slope = isNegative ? "Low" : "High";
            doCommand(":TRIGger:TIMeout:CONDition " + slope);
            doCommand(":TRIGger:TIMeout:TIME " + delay_s.ToString());
        }

        public void SetTrigModeTrans(double time_ns, bool isGthan = true, int source = 1, bool isRising = true)
        {
            string cmd = isGthan ? "GTHan" : "LTHan";
            doCommand(":TRIGger:TRANsition1:DIRection " + cmd);
            doCommand(":TRIGger:TRANsition1:SOURce " + "CHANnel" + source.ToString());
            doCommand(":TRIGger:TRANsition:TIME " + time_ns * Math.Pow(10, -9));
            doCommand(":TRIGger:TRANsition1:TYPE " + (isRising ? "RISetime" : "FALLtime"));
        }


        public double Meas_DeltaTime(int CH1, int CH2)
        {
            string cmd = string.Format(":MEASure:DELTatime? CHANnel{0}, CHANnel{1}", CH1, CH2);
            double num = doQueryNumber(cmd);
            return num;
        }

        public void SetDeltaTime_Rising_to_Rising(int start_pos, int stop_pos)
        {
            string cmd = string.Format(
                ":MEASure:DELTatime:DEFine {0}, {1}, {2}, {3}, {4}, {5}",
                "RISing", start_pos, "MIDDle", "RISing", stop_pos, "LOWer"
                );
            doCommand(cmd);
        }

        public void SetDeltaTime(bool isRising1, int start, int level1, bool isRising2, int stop, int level2)
        {
            string rising1 = isRising1 ? "RISing" : "Falling";
            string rising2 = isRising2 ? "RISing" : "Falling";
            string threshold1 = "LOWer";
            string threshold2 = "LOWer";

            switch (level1)
            {
                case 0:
                    threshold1 = "LOWer";
                    break;
                case 1:
                    threshold1 = "MIDDle";
                    break;
                case 2:
                    threshold1 = "UPPer";
                    break;
            }

            switch (level2)
            {
                case 0:
                    threshold2 = "LOWer";
                    break;
                case 1:
                    threshold2 = "MIDDle";
                    break;
                case 2:
                    threshold2 = "UPPer";
                    break;
            }




            string cmd = string.Format(
                            ":MEASure:DELTatime:DEFine {0}, {1}, {2}, {3}, {4}, {5}",
                            rising1, start, threshold1, rising2, stop, threshold2
                            );
            doCommand(cmd);
        }

        public void SetDeltaTime_Rising_to_Falling(int start_pos, int stop_pos)
        {
            string cmd = string.Format(
                ":MEASure:DELTatime:DEFine {0}, {1}, {2}, {3}, {4}, {5}",
                "RISing", start_pos, "MIDDle", "Falling", stop_pos, "LOWer"
                );
            doCommand(cmd);
        }

        private double Meas_Rise(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:RISetime? " + CHx;
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        private double Meas_Fall(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:FALLtime? " + CHx;
            doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1Rise()
        {
            return Meas_Rise(MEAS_CH1);
        }
        public double Meas_CH2Rise()
        {
            return Meas_Rise(MEAS_CH2);
        }
        public double Meas_CH3Rise()
        {
            return Meas_Rise(MEAS_CH3);
        }
        public double Meas_CH4Rise()
        {
            return Meas_Rise(MEAS_CH4);
        }

        public double Meas_CH1Fall()
        {
            return Meas_Fall(MEAS_CH1);
        }

        public double Meas_CH2Fall()
        {
            return Meas_Fall(MEAS_CH2);
        }

        public double Meas_CH3Fall()
        {
            return Meas_Fall(MEAS_CH3);
        }
        public double Meas_CH4Fall()
        {
            return Meas_Fall(MEAS_CH4);
        }
        private double Meas_Top(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:VTOP? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1Top()
        {
            return Meas_Top(MEAS_CH1);
        }
        public double Meas_CH2Top()
        {
            return Meas_Top(MEAS_CH2);
        }
        public double Meas_CH3Top()
        {
            return Meas_Top(MEAS_CH3);
        }
        public double Meas_CH4Top()
        {
            return Meas_Top(MEAS_CH4);
        }
        private double Meas_Base(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:VBASE? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }
        public double Meas_CH1Base()
        {
            return Meas_Base(MEAS_CH1);
        }

        public double Meas_CH2Base()
        {
            return Meas_Base(MEAS_CH2);
        }

        public double Meas_CH3Base()
        {
            return Meas_Base(MEAS_CH3);
        }

        public double Meas_CH4Base()
        {
            return Meas_Base(MEAS_CH4);
        }

        private double Meas_Freq(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:FREQuency? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1Freq()
        {
            return Meas_Freq(MEAS_CH1);
        }

        public double Meas_CH2Freq()
        {
            return Meas_Freq(MEAS_CH2);
        }

        public double Meas_CH3Freq()
        {
            return Meas_Freq(MEAS_CH3);
        }

        public double Meas_CH4Freq()
        {
            return Meas_Freq(MEAS_CH4);
        }

        private double Meas_Period(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:PERiod? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }
        
        public double Meas_CH1Period()
        {
            return Meas_Period(MEAS_CH1);
        }

        public double Meas_CH2Period()
        {
            return Meas_Period(MEAS_CH2);
        }

        public double Meas_CH3Period()
        {
            return Meas_Period(MEAS_CH3);
        }

        public double Meas_CH4Period()
        {
            return Meas_Period(MEAS_CH4);
        }

        private double Meas_MAX(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:VMAX? " + CHx;
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1MAX()
        {
            return Meas_MAX(CH1);
        }

        public double Meas_CH2MAX()
        {
            return Meas_MAX(CH2);
        }

        public double Meas_CH3MAX()
        {
            return Meas_MAX(CH3);
        }

        public double Meas_CH4MAX()
        {
            return Meas_MAX(CH4);
        }

        private double Meas_MIN(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:VMIN? " + CHx;
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1MIN()
        {
            return Meas_MIN(CH1);
        }

        public double Meas_CH2MIN()
        {
            return Meas_MIN(CH2);
        }

        public double Meas_CH3MIN()
        {
            return Meas_MIN(CH3);
        }

        public double Meas_CH4MIN()
        {
            return Meas_MIN(CH4);
        }

        private double Meas_XDelta(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MARKer:XDELta? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }
        public double Meas_CH1XDelta()
        {
            return Meas_XDelta(MEAS_CH1);
        }
        public double Meas_CH2XDelta()
        {
            return Meas_XDelta(MEAS_CH2);
        }
        public double Meas_CH3XDelta()
        {
            return Meas_XDelta(MEAS_CH3);
        }
        public double Meas_CH4XDelta()
        {
            return Meas_XDelta(MEAS_CH4);
        }

        private double Meas_Duty(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:DUTYcycle? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1Duty()
        {
            return Meas_Duty(MEAS_CH1);
        }

        public double Meas_CH2Duty()
        {
            return Meas_Duty(MEAS_CH2);
        }

        public double Meas_CH3Duty()
        {
            return Meas_Duty(MEAS_CH3);
        }

        public double Meas_CH4Duty()
        {
            return Meas_Duty(MEAS_CH4);
        }


        private double Meas_VPP(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:VPP? " + CHx;
            //doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1VPP()
        {
            return Meas_VPP(MEAS_CH1);
        }

        public double Meas_CH2VPP()
        {
            return Meas_VPP(MEAS_CH2);
        }

        public double Meas_CH3VPP()
        {
            return Meas_VPP(MEAS_CH3);
        }

        public double Meas_CH4VPP()
        {
            return Meas_VPP(MEAS_CH4);
        }

        //New Add
        public void CHx_Display(int channel, bool isdisplay)
        {
            string _isdisplay = isdisplay ? " 1" : " 0";
            string gogoCMD = ":CHANnel" + channel.ToString() + ":DISPlay" + _isdisplay;
            doCommand(gogoCMD);
        }
        public void CHx_Scale(int channel, double scale)
        {
            string gogoCMD = ":CHANnel" + channel.ToString() + ":SCALe " + scale.ToString();
            doCommand(gogoCMD);
        }

        public void SystemPresetDefault()
        {
            doCommand(":SYSTem:PRESet DEFault");
        }
        public void SetTrigModeEdge(bool isNegative)
        {
            doCommand(":TRIGger:MODE EDGE");
            string slope = isNegative ? "NEGative" : "POSitive";
            doCommand(":TRIGger:EDGE:SLOPe " + slope);
        }

        public void SweepModeTrig()
        {
            doCommand(":TRIGger:SWEep TRIGgered");
        }
        public void SweepModeAuto()
        {
            doCommand(":TRIGger:SWEep Auto");
        }
        public void MeasureStatisticsMean()
        {
            doCommand(":MEASure:STATistics MEAN");
            doCommand(":MEASure:SENDvalid ON");
        }
        public void MeasureStatisticsMin()
        {
            doCommand(":MEASure:STATistics Min");
            doCommand(":MEASure:SENDvalid ON");
        }
        public void MeasureStatisticsMax()
        {
            doCommand(":MEASure:STATistics Max");
            doCommand(":MEASure:SENDvalid ON");
        }
        public void MeasureStatisticsCurrent()
        {
            doCommand(":MEASure:STATistics CURRent");
            doCommand(":MEASure:SENDvalid ON");
        }
        public string GetMeasureStatistics()
        {
            return doQueryString(":MEASure:RESults?");
        }
        public void SetCHx_MeasureVmin(int channel)
        {
            string gogoCMD = ":MEASure:VMIN CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasureVmax(int channel)
        {
            string gogoCMD = ":MEASure:VMAX CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasureVbase(int channel)
        {
            string gogoCMD = ":MEASure:VBASe CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasureVtop(int channel)
        {
            string gogoCMD = ":MEASure:VTOP CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasureVave(int channel)
        {
            string gogoCMD = ":MEASure:VAVerage DISPLAY, CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasureVpp(int channel)
        {
            string gogoCMD = ":MEASure:VPP CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetCHx_MeasurePPulses(int channel)
        {
            string gogoCMD = ":MEASure:PPULses CHANnel" + channel.ToString();
            doCommand(gogoCMD);
        }
        public void SetMeasureHisTogram()
        {
            string gogoCMD = ":MEASure:HISTogram:PP HISTogram";
            doCommand(gogoCMD);
        }

        public double Meas_AVG(string CHx)
        {
            string cmd = ":MEASure:VAVerage? DISPlay, " + CHx;
            return doQueryNumber(cmd);
        }

        public double Meas_CH1AVG()
        {
            return Meas_AVG(CH1);
        }

        public double Meas_CH2AVG()
        {
            return Meas_AVG(CH2);
        }

        public double Meas_CH3AVG()
        {
            return Meas_AVG(CH3);
        }

        public double Meas_CH4AVG()
        {
            return Meas_AVG(CH4);
        }

        //public new void doCommand(string cmd)
        //{
        //    doCommand(cmd);
        //}

        public void DoCommand(string cmd)
        {
            doCommand(cmd);
            //doCommandViWrite(cmd + "\r\n");
        }

        public string doQeury(string cmd)
        {
            return doQueryString(cmd);
        }

        public string doRead()
        {
            return doReadString();
        }

        public double Meas_Result()
        {
            string cmd = ":MEASure:RESults?";
            return doQueryNumber(cmd);
        }


        private double Meas_PWidth(string CHx)
        {
            string cmd = ":MEASure:PWIDth?";
            DoCommand(CHx);
            return doQueryNumber(cmd);
        }


        public double Meas_CH1PWidth()
        {
            return Meas_PWidth(MEAS_CH1);
        }

        public double Meas_CH2PWidth()
        {
            return Meas_PWidth(MEAS_CH2);
        }
        public double Meas_CH3PWidth()
        {
            return Meas_PWidth(MEAS_CH3);
        }
        public double Meas_CH4PWidth()
        {
            return Meas_PWidth(MEAS_CH4);
        }

        /******* for lx, edited by Mo ******/

        public void Measurement_Thresholds_Absolute(double upper, double middle, double lower, int channel)
        {
            string cmd = ":MEASure:THResholds:GENeral:METHod CHANnel" + channel.ToString() + ",ABSolute";
            doCommand(cmd);
            cmd = ":MEASure:THResholds:GENeral:ABSolute CHANnel" + channel.ToString() + ","
                  + upper.ToString() + "," + middle.ToString() + "," + lower.ToString();
            doCommand(cmd);
        }

        public void Measurement_Threshold_Percent_Mode(int channel)
        {
            string cmd = ":MEASure:THResholds:GENeral:METHod CHANnel" + channel.ToString() + ",PERCent";
            doCommand(cmd);
        }

        public double Measure_SlewRate_Rising(int channel)
        {
            return doQueryNumber(":MEASure:SLEWrate? CHANnel"+ channel +",RISing");
        }
        public double Measure_SlewRate_Falling(int channel)
        {
            return doQueryNumber(":MEASure:SLEWrate? CHANnel" + channel + ",Falling");
        }

        public double Measure_Fall_Time(int channel)
        {
            return Meas_Fall(":MEASure:SOURce CHANnel" + channel);
        }

        public double Measure_Rise(int channel)
        {
            return Meas_Rise(":MEASure:SOURce CHANnel" + channel);
        }

        public void Bandwidth_Limit_On(int channel)
        {
            CHx_BWLimitOn("CHANNEL" + channel);
        }

        public void Bandwidth_Limit_Off(int channel)
        {
            CHx_BWLimitOff("CHANNEL" + channel);
        }

        public void Ch_On(int channel)
        {
            CHx_On("CHANNEL" + channel);
        }

        public void Ch_Off(int channel)
        {
            CHx_Off("CHANNEL" + channel);
        }

        public double Measure_Freq(int channel)
        {
            return Meas_Freq(":MEASure:SOURce CHANnel" + channel);
        }

        public void Ch_Offset(int channel, double offset)
        {
            CHx_Offset("CHANNEL"+channel, offset);
        }

        public double Measure_Ch_Max(int channel)
        {
            return Meas_MAX("CHANnel" + channel);
        }

        public double Measure_Ch_Min(int channel)
        {
            return Meas_MIN("CHANnel" + channel);
        }

        public double Measure_Ch_Vpp(int channel)
        {
            return Meas_VPP("CHANnel" + channel);
        }

        public double Measure_Top(int channel)
        {
            return Meas_Top(":MEASure:SOURce CHANnel" + channel);
        }

        public double Measure_Base(int channel)
        {
            return Meas_Base(":MEASure:SOURce CHANnel" + channel);
        }

        public void Trigger_Level(int channel, double level)
        {
            TriggerLevel("CHANNEL" + channel, level);
        }

        public void Trigger(int channel)
        {
            TriggerSource("CHANNEL" + channel);
        }

        public double Measure_Ch_XDelta(int channel)
        {
            return Meas_XDelta(":MEASure:SOURce CHANnel" + channel);
        }

        public double Measure_Count()
        {
            double result;
            doCommand(":MEASure:STATistics ON");
            doCommand(":MEASure:STATistics COUNt");

            result = doQueryNumber(":MEASure:RESults?");
            result *= 1000;

            return result;
        }

        public void TriggerRunt_POLarity(bool IsPositive)
        {
            DoCommand(string.Format(":TRIGger:RUNT:POLarity {0}", IsPositive ? "Pos" : "NEG"));
        }

        public void TriggerRunt_QUALified(bool IsOn)
        {
            DoCommand(string.Format(":TRIGger:RUNT:QUALified {0}", IsOn ? 1 : 0));
        }

        public void TriggerRunt_Source(int ch)
        {
            DoCommand(string.Format(":TRIGger:RUNT:SOURce CHANnel{0}", ch));
        }

        public void SetTriggerMode(string mode)
        {
            DoCommand(string.Format(":TRIGger:MODE {0}", mode));
        }

        public void SetTriggerLevel_HL(int ch, double level, bool IsHigh)
        {
            if (IsHigh) DoCommand(string.Format(":TRIGger:HTHReshold Channel{0}, {1}", ch, level));
            else DoCommand(string.Format(":TRIGger:LTHReshold Channel{0}, {1}", ch, level));
        }

        public void SetTimeoutCondition(bool IsHigh)
        {
            if (IsHigh) DoCommand(":TRIGger:TIMeout1:CONDition HIGH");
            else DoCommand(":TRIGger:TIMeout1:CONDition LOW");
        }

        public void SetTimeoutSource(int ch)
        {
            DoCommand(string.Format(":TRIGger:TIMeout1:SOURce CHANnel{0}", ch));
        }

        public void SetTimeoutTime(double ns)
        {
            DoCommand(string.Format(":TRIGger:TIMeout:TIME {0}ns", ns));
        }




        // save waveform data for csv
        public void SaveWaveformData(int Ch, ref double SSR, out double[] OutData, bool IsFunc = false)
        {
            if (device == 0)
            {
                OutData = null;
                return;
            }
            DoCommand(":SYSTEM:HEADER OFF");
            DoCommand(":WAVeform:FORMat Byte");
            if (IsFunc)
                DoCommand(":WAVeform:SOURce FUNC" + Ch.ToString());
            else
                DoCommand(":WAVeform:SOURce CHAN" + Ch.ToString());
            DoCommand(":WAVeform:STReaming ON");
            double X_INS = doQueryNumber(":WAVEFORM:XINCrement?");
            double X_ORG = doQueryNumber(":WAVEFORM:XORigin?");
            double Y_INS = doQueryNumber(":WAVEFORM:YINCrement?");
            double Y_ORG = doQueryNumber(":WAVEFORM:YORigin?");
            int Point = (int)doQueryNumber(":WAVEFORM:POINTS?");
            int NewLen = Point;
            int step = 1;
            if (X_INS > SSR) SSR = X_INS;
            else
            {
                step = (int)Math.Round(SSR / X_INS, 0);
                NewLen /= step;
            }
            byte[] Arr = new byte[Point];
            DoCommand(":WAVeform:DATA? 1," + Point.ToString());
            int cnt = IEEEBlock_Bytes2(ref Arr);
            Console.WriteLine("N:{0},Xin:{1},Xor:{2},Yin:{3},Yor{4},FB:{5}", Point, X_INS, X_ORG, Y_INS, Y_ORG, cnt);
            double data;
            OutData = new double[NewLen];
            int idx = 0;
            for (int i = 2; i < Arr.Length && idx < NewLen; i += step)
            {
                if (Arr[i] > 127)
                    data = (Arr[i] - 256) * Y_INS + Y_ORG;
                else
                    data = (Arr[i]) * Y_INS + Y_ORG;
                OutData[idx++] = data;
            }
            /*using (StreamWriter sw = new StreamWriter(path))   //小寫TXT     
            {
                for (int i = 0; i < Arr.Length; i += step)
                {
                    if (Arr[i] > 127)
                        data = (Arr[i] - 256) * Y_INS + Y_ORG;
                    else
                        data = (Arr[i]) * Y_INS + Y_ORG;
                    sw.WriteLine(data);
                }
            }*/
            Arr = null;
        }
    }
}


