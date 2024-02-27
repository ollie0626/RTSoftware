using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

namespace InsLibDotNet
{
    public class AgilentOSC : VisaCommand
    {
        string MEAS_CH1 = ":MEASure:SOURce CHANnel1";
        string MEAS_CH2 = ":MEASure:SOURce CHANnel2";
        string MEAS_CH3 = ":MEASure:SOURce CHANnel3";
        string MEAS_CH4 = ":MEASure:SOURce CHANnel4";
        string CH1 = "CHANNEL1";
        string CH2 = "CHANNEL2";
        string CH3 = "CHANNEL3";
        string CH4 = "CHANNEL4";

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

        public void TimeBasePosition(double position)
        {
            doCommand(":TIMEBASE:POSITION " + position.ToString() + "us");
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
            string buf = Path.Substring(Path.Length - 1, 1) == @"/" ? Path.Substring(0, Path.Length - 1) : Path;
            string buf_tmp = Path.Substring(Path.Length - 1, 1) == @"/" ? Path.Substring(0, Path.Length - 1) : Path;
            buf = buf +  FileName + ".png";

            FileStream fStream = File.Open(buf_tmp + @"__temp.png", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);

            string gogoCMD = ":DISPlay:DATA? PNG";
            doCommand(gogoCMD); System.Threading.Thread.Sleep(500);
            byte[] ResultsArray = new byte[300000];
            IEEEBlock_Bytes(out ResultsArray);
            fStream.Write(ResultsArray, 0, ResultsArray.Length);
            System.Threading.Thread.Sleep(1000);
            fStream.Close();
            fStream.Dispose();

            CompressImage(buf_tmp + @"__temp.png", buf);
#endif
        }


        public void CompressImage(string sFile, string dFile, int size = 300, bool sfsc = true)
        {
            Image iSource = Image.FromFile(sFile);
            ImageFormat tFormat = iSource.RawFormat;
            FileInfo firstFileInfo = new FileInfo(sFile);

            if (sfsc == true && firstFileInfo.Length < size * 1024)
            {
                firstFileInfo.CopyTo(dFile);
            }

            int dHeight = Convert.ToInt16(Convert.ToDouble(iSource.Height) * 0.8);
            int dWidth = Convert.ToInt16(Convert.ToDouble(iSource.Width) * 0.8);
            int sW = 0, sH = 0;
            Size tem_size = new Size(iSource.Width, iSource.Height);
            if (tem_size.Width > dHeight || tem_size.Width > dWidth)
            {
                if ((tem_size.Width * dHeight) > (tem_size.Width * dWidth))
                {
                    /* high > width */
                    sW = dWidth;
                    sH = (dWidth * tem_size.Height) / tem_size.Width;
                }
                else
                {
                    /* width > high */
                    sH = dHeight;
                    sW = (tem_size.Width * dHeight) / tem_size.Height;
                }
            }
            else
            {
                sW = tem_size.Width;
                sH = tem_size.Height;
            }


            Bitmap ob = new Bitmap(dWidth, dHeight);
            Graphics g = Graphics.FromImage(ob);

            g.Clear(Color.WhiteSmoke);
            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.DrawImage(iSource, new Rectangle((dWidth - sW) / 2, (dHeight - sH) / 2, sW, sH), 0, 0, iSource.Width, iSource.Height, GraphicsUnit.Pixel);

            g.Dispose();
            ob.Save(dFile, tFormat);
            iSource.Dispose();
            ob.Dispose();

            File.Delete(sFile);
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
            doCommand(CHx);
            double amp = doQueryNumber(":MEASure:VAMPlitude?");
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

        public void CH1_10_90Range()
        {
            SlewRate_90Range(CH1);
        }

        public void CH2_10_90Range()
        {
            SlewRate_90Range(CH2);
        }

        public void CH3_10_90Range()
        {
            SlewRate_90Range(CH3);
        }

        public void CH4_10_90Range()
        {
            SlewRate_90Range(CH4);
        }



        public void SlewRate20_80Range()
        {
            string cmd = ":MEASure:THResholds:METHod ALL,PERCent";
            doCommand(cmd);
            cmd = ":MEASure:THResholds:GEN:PERCent ALL,80,50,20";
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
            string gogoCMD = ":" + CHx + ":BWLIMIT ON";
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

        /* Measure command */

        public double MeasDelta(int src1, int src2)
        {
            double delta = 0;
            string cmd = ":MEASure:DELTatime? CHANnel" + src1 + ", CHANnel" + src2;
            delta = doQueryNumber(cmd);
            return delta;
        }




        private double Meas_Rise(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:RISetime?";
            doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        private double Meas_Fall(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:FALLtime?";
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
            return Meas_Rise(MEAS_CH4);
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
            string gogoCMD = ":MEASure:VTOP?";
            doCommand(CHx);
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
            string gogoCMD = ":MEASure:VBASE?";
            doCommand(CHx);
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
            string gogoCMD = ":MEASure:FREQuency?";
            doCommand(CHx);
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
            string gogoCMD = ":MEASure:PERiod?";
            doCommand(CHx);
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
            string gogoCMD = ":MEASure:VMAX?";
            doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1MAX()
        {
            return Meas_MAX(MEAS_CH1);
        }

        public double Meas_CH2MAX()
        {
            return Meas_MAX(MEAS_CH2);
        }

        public double Meas_CH3MAX()
        {
            return Meas_MAX(MEAS_CH3);
        }

        public double Meas_CH4MAX()
        {
            return Meas_MAX(MEAS_CH4);
        }

        private double Meas_MIN(string CHx)
        {
            
            double buf;
            string gogoCMD = ":MEASure:VMIN?";
            doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }

        public double Meas_CH1MIN()
        {
            return Meas_MIN(MEAS_CH1);
        }

        public double Meas_CH2MIN()
        {
            return Meas_MIN(MEAS_CH2);
        }

        public double Meas_CH3MIN()
        {
            return Meas_MIN(MEAS_CH3);
        }

        public double Meas_CH4MIN()
        {
            return Meas_MIN(MEAS_CH4);
        }

        private double Meas_XDelta(string CHx)
        {
            double buf;
            string gogoCMD = ":MARKer:XDELta?";
            doCommand(CHx);
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


        private double Meas_VPP(string CHx)
        {

            double buf;
            string gogoCMD = ":MEASure:VPP?";
            doCommand(CHx);
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
        private double Meas_AVE(string CHx)
        {
            double buf;
            string gogoCMD = ":MEASure:VAVE?";
            doCommand(CHx);
            buf = doQueryNumber(gogoCMD);
            return buf;
        }
        public double Meas_CH1AVE()
        {
            return Meas_MIN(MEAS_CH1);
        }
        public double Meas_CH2AVE()
        {
            return Meas_MIN(MEAS_CH2);
        }
        public double Meas_CH3AVE()
        {
            return Meas_MIN(MEAS_CH3);
        }
        public double Meas_CH4AVE()
        {
            return Meas_MIN(MEAS_CH4);
        }

        //public new void doCommand(string cmd)
        //{
        //    doCommand(cmd);
        //}

        public void DoCommand(string cmd)
        {
            doCommand(cmd);
        }

        public double DoQueryNumber(string cmd)
        {
            return doQueryNumber(cmd);
        }

        public string DoQueryString(string cmd)
        {
            return doQueryString(cmd);
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



        private double Meas_NWidth(string CHx)
        {
            string cmd = ":MEASure:NWIDth?";
            DoCommand(CHx);
            return doQueryNumber(cmd);
        }


        public double Meas_CH1NWidth()
        {
            return Meas_NWidth(MEAS_CH1);
        }

        public double Meas_CH2NWidth()
        {
            return Meas_NWidth(MEAS_CH2);
        }

        public double Meas_CH3NWidth()
        {
            return Meas_NWidth(MEAS_CH3);
        }

        public double Meas_CH4NWidth()
        {
            return Meas_NWidth(MEAS_CH4);
        }


    }
}
