using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace InsLibDotNet
{
    public class PowerN6705 : VisaCommand
    {
        public PowerN6705()
        {
            LinkingIns("GPIB0::4::INSTR");
        }

        public PowerN6705(string Addr)
        {
            LinkingIns(Addr);
        }

        public PowerN6705(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void PowerOff(int ch)
        {
            doCommand(string.Format("OUTP OFF,(@{0})",ch));
        }

        public void PowerOff_All()
        {
            doCommand("OUTP OFF,(@1:4)");
        }

        void SetLimitCurr(int ch, double range)
        {
            if(range > 20.39) doCommand(string.Format("CURR:LIM {1},(@{0})", ch, 4.08));
            else if (range > 15.29) doCommand(string.Format("CURR:LIM {1},(@{0})", ch, 5.1));
            else if (range > 10.19) doCommand(string.Format("CURR:LIM {1},(@{0})", ch, 6.83));
            else  doCommand(string.Format("CURR:LIM {1},(@{0})", ch, 8.16));
        }

        public void VoltageSetting(int ch, double vol, double range)
        {
            doCommand(string.Format("VOLT:RANG {1},(@{0})",ch, range));
            doCommand(string.Format("VOLT {1},(@{0})", ch, vol));
            doCommand(string.Format("VOLT:RANG {1},(@{0})", ch, range));
            SetLimitCurr(ch, range);
        }

        public void AutoPowerOn(int ch)
        {
            doCommand(string.Format("OUTP On,(@{0})", ch));
        }

        public double GetVoltage(int ch)
        {
            return doQueryNumber(string.Format("MEAS:VOLT? (@{0})", ch));
        }

        public void SetVoltageSR(int ch,int VperS, bool IsMax = true)
        {
            if(IsMax) doCommand(string.Format("VOLT:SLEW INF,(@{0})", ch));
            else doCommand(string.Format("VOLT:SLEW {1},(@{0})", ch, VperS));
        }

        public void SetOCPOn()
        {
            doCommand(string.Format("OUTP:PROT:COUP ON"));
        }


        public void SetArbWave_Count(bool LastDCvalue, int channel, int Count = 99999)
        {
            string command = "";
            if (LastDCvalue) command = "ARB:TERM:LAST OFF, " + String.Format("(@{0})", channel);
            else command = "ARB:TERM:LAST ON, " + String.Format("(@{0})", channel);
            doCommand(command);

            //set continuous as default
            if(Count >= 99999) doCommand(string.Format("ARB:COUN INF, (@{0})", channel));
            else doCommand(string.Format("ARB:COUN {1}, (@{0})", channel,Count));
        }

        public void SetArbWave_UserDef(List<double> points, int ch, double FixTimeStep = 0.001)
        {
            string command = "ARB:VOLT:UDEF:LEV ";
            string Time_command = "ARB:VOLT:UDEF:DWEL ";
            for (int i = 0; (i < points.Count) && (i < 512); ++i)
            {
                command += points[i].ToString("F3") + ",";
                Time_command += FixTimeStep.ToString() + ",";
            }
            command += string.Format("(@{0})", ch);
            Time_command += string.Format("(@{0})", ch);
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", ch));
            doCommand(string.Format("ARB:FUNC:SHAP UDEF,(@{0})", ch));
            doCommand(command);
            doCommand(Time_command);
        }

        public void SetArbWave_UserDef(List<double> timestep, List<double> points, int ch)
        {
            string command = "ARB:VOLT:UDEF:LEV ";
            string Time_command = "ARB:VOLT:UDEF:DWEL ";
            for (int i = 0; (i < timestep.Count) && (i < points.Count) && (i < 512); ++i)
            {
                command += points[i].ToString("F3") + ",";
                Time_command += timestep[i].ToString("F3") + ",";
            }
            command += string.Format("(@{0})", ch);
            Time_command += string.Format("(@{0})", ch);
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", ch));
            doCommand(string.Format("ARB:FUNC:SHAP UDEF,(@{0})", ch));
            doCommand(command);
            doCommand(Time_command);
        }

        public void SetArbWave_Trapezoid(double v0,
                               double v1,
                               double t0_s,
                               double t1_s,
                               double t2_s,
                               double t3_s,
                               double t4_s,
                               int channel)
        {
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", channel));
            doCommand(string.Format("ARB:FUNC:SHAP TRAP,(@{0})", channel));
            string command = "";
            command = "ARB:VOLT:TRAP:STAR " + String.Format("{0:0.###}, ", v0) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:TOP " + String.Format("{0:0.###}, ", v1) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:STAR:TIM " + String.Format("{0:0.###}, ", t0_s) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:RTIM " + String.Format("{0:0.###}, ", t1_s) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:TOP:TIM " + String.Format("{0:0.###}, ", t2_s) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:FTIM " + String.Format("{0:0.###}, ", t3_s) + String.Format("(@{0})", channel); doCommand(command);
            command = "ARB:VOLT:TRAP:END:TIM " + String.Format("{0:0.###}, ", t4_s) + String.Format("(@{0})", channel); doCommand(command);
        }

        public void SetArbWave_Ramp(double v0,
                       double v1,
                       double t0_s,
                       double t1_s,
                       double t2_s,
                       int channel)
        {
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", channel));
            doCommand(string.Format("ARB:FUNC:SHAP RAMP,(@{0})", channel));
            //string command = "";
            doCommand(string.Format("ARB:VOLT:RAMP:STAR {1},(@{0})", channel, v0));
            doCommand(string.Format("ARB:VOLT:RAMP:END {1},(@{0})", channel, v1));
            doCommand(string.Format("ARB:VOLT:RAMP:STAR:TIM {1},(@{0})", channel, t0_s));
            doCommand(string.Format("ARB:VOLT:RAMP:RTIM {1},(@{0})", channel, t1_s));
            doCommand(string.Format("ARB:VOLT:RAMP:END:TIM {1},(@{0})", channel, t2_s));
        }

        public void SetArbWave_Staricase(double v0,
                double v1,
                double t0_s,
                double t1_s,
                double t2_s,
                int step,
                int channel)
        {
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", channel));
            doCommand(string.Format("ARB:FUNC:SHAP STA,(@{0})", channel));

            doCommand(string.Format("ARB:VOLT:STA:STAR {1},(@{0})", channel, v0));
            doCommand(string.Format("ARB:VOLT:STA:END {1},(@{0})", channel, v1));
            doCommand(string.Format("ARB:VOLT:STA:STAR:TIM {1},(@{0})", channel, t0_s));
            doCommand(string.Format("ARB:VOLT:STA:TIM {1},(@{0})", channel, t1_s));
            doCommand(string.Format("ARB:VOLT:STA:END:TIM {1},(@{0})", channel, t2_s));
            doCommand(string.Format("ARB:VOLT:STA:NST {1},(@{0})", channel, step));
        }


        public void SetArbWave_Pulse(double v0,
               double v1,
               double t0_s,
               double t1_s,
               double t2_s,
               int channel)
        {
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", channel));
            doCommand(string.Format("ARB:FUNC:SHAP PULS,(@{0})", channel));
            //string command = "";
            doCommand(string.Format("ARB:VOLT:PULS:STAR {1},(@{0})", channel, v0));
            doCommand(string.Format("ARB:VOLT:PULS:TOP {1},(@{0})", channel, v1));
            doCommand(string.Format("ARB:VOLT:PULS:STAR:TIM {1},(@{0})", channel, t0_s));
            doCommand(string.Format("ARB:VOLT:PULS:TOP:TIM {1},(@{0})", channel, t1_s));
            doCommand(string.Format("ARB:VOLT:PULS:END:TIM {1},(@{0})", channel, t2_s));
        }

        public void RunArbWave(int ch, int type = 0)
        {
            doCommand(string.Format("VOLT:MODE ARB, (@{0})", ch));
            SetTrigger(type);
            doCommand(string.Format("OUTP ON, (@{0})", ch));
            doCommand(string.Format("INIT:TRAN (@{0})", ch));
        }


        public void RunArbWave(int start, int end, int type = 0)
        {
            doCommand(string.Format("VOLT:MODE ARB, (@{0}:{1})", start, end));
            SetTrigger(type);
            doCommand(string.Format("OUTP ON, (@{0}:{1})", start, end));
            doCommand(string.Format("INIT:TRAN (@{0}:{1})", start, end));
        }

        public void SetTrigger(int type)
        {
            if(type == 0) doCommand(string.Format("TRIG:ARB:SOUR IMM"));
            else doCommand(string.Format("TRIG:ARB:SOUR EXT"));
        }

        public void StopArbWave(int ch)
        {
            doCommand(string.Format("ABOR:TRAN (@{0})", ch));
        }

        public void SetSeqArbWave(int ch)
        {
            doCommand(string.Format("ARB:FUNC:TYPE VOLT,(@{0})", ch));
            doCommand(string.Format("ARB:FUNC:SHAP SEQ,(@{0})", ch));
            doCommand(string.Format("ARB:SEQ:RESet,(@{0})", ch));
        }

        public void SetSeqArbWave_Pulse(int seq, double v0,
              double v1,
              double t0_s,
              double t1_s,
              double t2_s,
              int channel)
        {
            doCommand(string.Format("ARB:SEQ:STEP:FUNC:SHAP PULS, {1},(@{0})", channel, seq));
            doCommand(string.Format("ARB:SEQ:STEP:VOLT:PULS:STAR {1},{2},(@{0})", channel, v0, seq));
            doCommand(string.Format("ARB:SEQ:STEP:VOLT:PULS:TOP {1},{2},(@{0})", channel, v1, seq));
            doCommand(string.Format("ARB:SEQ:STEP:VOLT:PULS:STAR:TIM {1},{2},(@{0})", channel, t0_s, seq));
            doCommand(string.Format("ARB:SEQ:STEP:VOLT:PULS:TOP:TIM {1},{2},(@{0})", channel, t1_s, seq));
            doCommand(string.Format("ARB:SEQ:STEP:VOLT:PULS:END:TIM {1},{2},(@{0})", channel, t2_s, seq));
        }

        public void SetSeqArbWave_Ramp(int seq, double v0,
               double v1,
               double t0_s,
               double t1_s,
               double t2_s,
               int channel)
        {
            doCommand(string.Format("ARB:SEQ:FUNC:SHAP RAMP, {1},(@{0})", channel, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:RAMP:STAR {1},{2},(@{0})", channel, v0, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:RAMP:END {1},{2},(@{0})", channel, v1, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:RAMP:STAR:TIM {1},{2},(@{0})", channel, t0_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:RAMP:RTIM {1},{2},(@{0})", channel, t1_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:RAMP:END:TIM {1},{2},(@{0})", channel, t2_s, seq));
        }

        public void SetSeqWave_Trapezoid(int seq, double v0,
                       double v1,
                       double t0_s,
                       double t1_s,
                       double t2_s,
                       double t3_s,
                       double t4_s,
                       int channel)
        {
            doCommand(string.Format("ARB:SEQ:FUNC:SHAP TRAP, {1},(@{0})", channel, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:STAR {1},{2},(@{0})", channel, v0, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:TOP {1},{2},(@{0})", channel, v1, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:STAR:TIM {1},{2},(@{0})", channel, t0_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:RTIM {1},{2},(@{0})", channel, t1_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:TOP:TIM {1},{2},(@{0})", channel, t2_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:FTIM {1},{2},(@{0})", channel, t3_s, seq));
            doCommand(string.Format("ARB:SEQ:VOLT:TRAP:END:TIM {1},{2},(@{0})", channel, t4_s, seq));
        }

        //EMUL PS1Q,(@1)
        //VOLT:RANG 5,(@1)

        public void SetPowerSupplyRange(int channel, double vol)
        {
            doCommand(string.Format("EMUL PS1Q,(@{0})", channel));
            doCommand(string.Format("VOLT:RANG {1},(@{0})", channel, vol));
        }





        /*public void CoupleSource(bool Ch1, bool Ch2 = false, bool Ch3 = false, bool Ch4 = false)
        {
            
            return doQueryNumber(string.Format("OUTP:COUP:CHAN 1", ch));
        }*/

        public double GetCurrent(int ch)
        {
            return doQueryNumber(string.Format("MEAS:CURR? (@{0})", ch));
        }

        public void AutoPowerOn_All()
        {
            doCommand("OUTP On,(@1:4)");
        }

        public void SaveWaveform(string Path, string FileName)
        {
            int datalen = 180000;
            Console.WriteLine("Save Pic Path : {0}", Path + "\\" + FileName + ".gif");
            //byte[] bytRead = new byte[datalen];
            string gogoCMD = "HCOPy:SDUMp:DATA?";
            doCommand(gogoCMD); System.Threading.Thread.Sleep(3000);
            byte[] ResultsArray = new byte[datalen];
            int nViStatus;
            nViStatus = visa32.viScanf(device, "%#b", ref datalen, ResultsArray);

            FileStream fStream = File.Open(Path + "\\" + FileName + ".gif", FileMode.Create);
            fStream.Write(ResultsArray, 0, datalen);
            //System.Threading.Thread.Sleep(500);
            fStream.Close();
            fStream.Dispose();

            ResultsArray = null;
            //bytRead = null;
        }
    }
}
