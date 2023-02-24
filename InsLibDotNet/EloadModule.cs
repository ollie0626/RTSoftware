using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class EloadModule : VisaCommand
    {
        public string CH1 = "CHAN 1";
        public string CH2 = "CHAN 2";
        public string CH3 = "CHAN 3";
        public string CH4 = "CHAN 4";

        ~EloadModule()
        {
            InsClose();
        }

        public EloadModule()
        {
            LinkingIns("GPIB0::7::INSTR");

        }

        public EloadModule(int Addr)
        {
            //GPIB0::7::INSTR
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public EloadModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public void ConnectELoad(string Addr)
        {
            LinkingIns(Addr);
        }

        public void DoCommand(string cmd)
        {
            doCommand(cmd);
        }

        public void ChannelSel(int ch)
        {
            string cmd = "CHAN " + ch.ToString();
            DoCommand(cmd);
        }

        public void ChannelActive(int ch)
        {
            string cmd = "CHANnel:ACTive ON";
            DoCommand(cmd);
        }

        public void ConnectELoad(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }
        public void CCL_Mode()
        {
            string mode = "MODE CCL";
            doCommand(mode);
        }

        public void CCDL_Mode()
        {
            string mode = "MODE CCDL";
            doCommand(mode);
        }

        public void CCDM_Mode()
        {
            string mode = "MODE CCDM";
            doCommand(mode);
        }

        public void CCM_Mode()
        {
            string mode = "MODE CCM";
            doCommand(mode);
        }

        public void CCDH_Mode()
        {
            string mode = "MODE CCDH";
            doCommand(mode);
        }

        public void CCH_Mode()
        {
            string mode = "MODE CCH";
            doCommand(mode);
        }

        public void CV_Mode()
        {
            string mode = "MODE CVH";
            doCommand(mode);
        }

        public void SWDL_Mode()
        {
            string mode = "MODE SWDL";
            doCommand(mode);
        }

        public void SWDM_Mode()
        {
            string mode = "MODE SWDM";
            doCommand(mode);
        }

        public void SWDH_Mode()
        {
            string mode = "MODE SWDH";
            doCommand(mode);
        }

        public void SetCV_Vol(double vol)
        {
            string cmd = "VOLT:STAT:ILIM MAX";
            doCommand(cmd);

            cmd = "VOLTage:STATic:L1 " + vol;
            doCommand(cmd);

            cmd = "LOAD ON";
            doCommand(cmd);
        }

        public void SetCV_Current(double current)
        {
            string cmd = "VOLT:STAT:ILIM " + current;
            doCommand(cmd);
        }

        public void SetCV_VolMode(bool isFast)
        {
            string cmd = "VOLT:STAT:RES " + (isFast ? "FAST" : "SLOW");
            doCommand(cmd);
        }

        public double Meas_Vol(int ch)
        {
            string cmd = "MEASure:VOLTage?";
            ChannelSel(ch);
            return doQueryNumber(cmd);
        }

        public double Meas_Curr(int ch)
        {
            string cmd = "MEASure:CURRent?";
            ChannelSel(ch);
            return doQueryNumber(cmd);
        }


        private void DymanicLoadClear(string CHx)
        {
            doCommand(CHx);
            
            string cmd;
            CCDL_Mode();
            cmd = "CURRent:DYNamic:L1 0A";
            doCommandViWrite(cmd);
            cmd = "CURRent:DYNamic:L2 0A";
            doCommandViWrite(cmd);

            CCDM_Mode();
            cmd = "CURRent:DYNamic:L1 0A";
            doCommandViWrite(cmd);
            cmd = "CURRent:DYNamic:L2 0A";
            doCommandViWrite(cmd);

            CCDH_Mode();
            cmd = "CURRent:DYNamic:L1 0A";
            doCommandViWrite(cmd);
            cmd = "CURRent:DYNamic:L2 0A";
            doCommandViWrite(cmd);
        }


        public void CH1_DymanicClear()
        {
            DymanicLoadClear(CH1);
        }

        public void CH2_DymanicClear()
        {
            DymanicLoadClear(CH2);
        }

        public void CH3_DymanicClear()
        {
            DymanicLoadClear(CH3);
        }

        public void CH4_DymanicClear()
        {
            DymanicLoadClear(CH4);
        }


        public void DymanicLoad(string CHx, 
                                double L1, double L2,
                                double T1, double T2)
        {
            doCommand(CHx);
            double MaxCurr = L2;
            if(L1 > L2)
                MaxCurr = L1;
            else
                MaxCurr = L2;

            if (MaxCurr <= 0.1)
                CCDL_Mode();
            else if (MaxCurr >= 0.1 && MaxCurr <= 1)
                CCDM_Mode();
            else if (MaxCurr >= 1)
                CCDH_Mode();

            string cmd = "CURR:DYN:RISE MAX";
            doCommand(cmd);
            cmd = "CURR:DYN:FALL MAX";
            doCommand(cmd);

            cmd = "CURRent:DYNamic:L1 " + L1 + "A";
            doCommandViWrite(cmd);
            //doCommand(cmd);
            cmd = "CURRent:DYNamic:L2 " + L2 + "A";
            doCommandViWrite(cmd);

            cmd = "CURR:DYN:T1 " + T1 + "mS";
            doCommand(cmd);
            cmd = "CURR:DYN:T2 " + T2 + "mS";
            doCommand(cmd);

            cmd = "LOAD ON";
            doCommand(cmd);
        }


        public void DymanicCH1(double L1, double L2,
                               double T1, double T2)
        {
            DymanicLoad(CH1, L1, L2, T1, T2);
        }

        public void DymanicCH2(double L1, double L2,
                               double T1, double T2)
        {
            DymanicLoad(CH2, L1, L2, T1, T2);
        }

        public void DymanicCH3(double L1, double L2,
                               double T1, double T2)
        {
            DymanicLoad(CH3, L1, L2, T1, T2);
        }

        public void DymanicCH4(double L1, double L2,
                               double T1, double T2)
        {
            DymanicLoad(CH4, L1, L2, T1, T2);
        }

        public void AllChannel_LoadOff()
        {
            string CMD_Off = "LOAD OFF";
            doCommand(CH4);
            doCommand(CMD_Off);
            doCommand(CH3);
            doCommand(CMD_Off);
            doCommand(CH2);
            doCommand(CMD_Off);
            doCommand(CH1);
            doCommand(CMD_Off);
        }

        public void SyncOn()
        {
            string cmd = ":CHAN:SYNC ON";
            doCommand(cmd);
        }

        public void Loading(string CH, double iout)
        {
            string gogoCMD;
            doCommand(CH);
            System.Threading.Thread.Sleep(50);
            //if (iout < 0.2)
            //    CCL_Mode();
            //else if (iout >= 0.2 && iout < 1)
            //    CCM_Mode();
            //else if(iout >= 1 && iout < 20)
            //    CCH_Mode();
            System.Threading.Thread.Sleep(150);


            gogoCMD = "CURR:STAT:L2 " + string.Format("{0:0.####}", iout);
            doCommand(gogoCMD);
            gogoCMD = "LOAD ON";
            doCommand(gogoCMD);
        }

        public void Loading(int CH, double iout)
        {
            string gogoCMD;
            doCommand("CHAN " + CH.ToString());
            System.Threading.Thread.Sleep(50);
            //if (iout < 0.2)
            //    CCL_Mode();
            //else if (iout >= 0.2 && iout < 1)
            //    CCM_Mode();
            //else if(iout >= 1 && iout < 20)
            //    CCH_Mode();
            System.Threading.Thread.Sleep(150);


            gogoCMD = "CURR:STAT:L2 " + string.Format("{0:0.####}", iout);
            doCommand(gogoCMD);
            gogoCMD = "LOAD ON";
            doCommand(gogoCMD);
        }


        private void CHx_ClearSetting(string CH)
        {
            string cmd;
            doCommand(CH);
            cmd = "CURR:STAT:L2 0";
            CCL_Mode();
            doCommand(cmd);
            CCM_Mode();
            doCommand(cmd);
            CCM_Mode();
            doCommand(cmd);
        }

        public void CH1_ClearSetting()
        {
            CHx_ClearSetting(CH1);
        }

        public void CH2_ClearSetting()
        {
            CHx_ClearSetting(CH2);
        }
        public void CH3_ClearSetting()
        {
            CHx_ClearSetting(CH3);
        }
        public void CH4_ClearSetting()
        {
            CHx_ClearSetting(CH4);
        }

        public void CH1_Loading(double iout)
        {
            Loading(CH1, iout);
        }

        public void CH2_Loading(double iout)
        {
            Loading(CH2, iout);
        }

        public void CH3_Loading(double iout)
        {
            Loading(CH3, iout);
        }

        public void CH4_Loading(double iout)
        {
            Loading(CH4, iout);
        }

        public double GetVol()
        {
            string buffer = "MEAS:VOLT?";
            return doQueryNumber(buffer);
        }

        public double GetIout()
        {
            string buffer = "MEAS:CURR?";
            return doQueryNumber(buffer);
        }

        public double[] GetAllChannel_Vol()
        {
            double[] vol = new double[4];
            string buffer = "MEAS:ALLV?";
            string[] temp;
            buffer = doQueryString(buffer);
            temp = buffer.Split(',');
            for (int i = 0; i < 4; i++)
            {
                vol[i] = Convert.ToDouble(string.Format("{0:0.###}", temp[i]));
            }
            return vol;
        }

        public double[] GetAllChannel_Iout()
        {
            double[] vol = new double[4];
            string buffer = "MEAS:ALLC?";
            string[] temp;
            buffer = doQueryString(buffer);
            temp = buffer.Split(',');
            for (int i = 0; i < 4; i++)
            {
                vol[i] = Convert.ToDouble(string.Format("{0:0.###}", temp[i]));
            }
            return vol;
        }

        /* advance function */
        public void SetAdvanceDwell(double dwell)
        {
            string cmd = "ADV:OCP:DEWL " + dwell + "us";
            doCommand(cmd);
        }


        public void SetAdvanceSpecHL(double Lo, double Hi)
        {
            string cmd = "ADV:OCP:SPEC:H " + Hi;
            doCommand(cmd);
            cmd = "ADV:OCP:SPEC:H " + Lo;
            doCommand(cmd);
        }

        public void SetAdvanceCurrentRange(double ISTA, double IEND)
        {
            string cmd = "ADV:OCP:ISTA " + ISTA;
            doCommand(cmd);
            cmd = "ADV:OCP:ISTA " + IEND;
            doCommand(cmd);
        }

        public void SetAdvanceVol(double vol)
        {
            string cmd = "ADV:OCP:TRIG:VOL " + vol;
            doCommand(cmd);
        }

        public void SetAdvanceLatchOn()
        {
            string cmd = "ADV:OCP:LATC ON";
            doCommand(cmd);
        }

        public void SetAdvanceLatchOff()
        {
            string cmd = "ADV:OCP:LATC OFF";
            doCommand(cmd);
        }

        public void SetAdvanceOCPStep(int step)
        {
            string cmd = "ADV:OCP:STEP " + step;
            doCommand(cmd);
        }

        public double GetAdvanceOCPResult()
        {
            string cmd;
            string res;
            string[] tmp;
            double doures = 0;

            try
            {
                cmd = "ADVance:OCP:RESult?";
                res = doQueryString(cmd);
                tmp = res.Split(',');
                doures = Convert.ToDouble(tmp[1]) * 1000;
            }
            catch
            {
                Console.WriteLine("Advance res: " + doures);
            }
            return doures;
        }

        // support UVP delay time

        public void ShortOn()
        {
            string cmd = "LOAD:SHOR ON";
            doCommand(cmd);
        }

        public void ShortOff()
        {
            string cmd = "LOAD:SHOR OFF";
            doCommand(cmd);
        }

        // support sine wave Eload

        public void SetAdvanceSine_Freq(double freq_hz)
        {
            string cmd = string.Format("ADV:SINE:FREQ {0}", freq_hz);
            doCommand(cmd);
        }

        public void SetAdvanceSine_IAC(double IAC)
        {
            string cmd = string.Format("ADV:SINE:IAC {0}", IAC);
            doCommand(cmd);
        }

        public void SetAdvanceSine_IDC(double IDC)
        {
            string cmd = string.Format("ADV:SINE:IDC {0}", IDC);
            doCommand(cmd);
        }
    }
}
