using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class FuncGenModule : VisaCommand
    {
        ~FuncGenModule()
        {
            InsClose();
        }

        public FuncGenModule()
        {
            LinkingIns("GPIB0::11::INSTR");
        }

        public FuncGenModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public FuncGenModule(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void ConnectFuncGen(string Addr)
        {
            LinkingIns(Addr);
        }

        public void ConnectFuncGen(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void CH1_ContinuousMode()
        {
            string CMD = "SOURce1:BURSt:STATe 0";
            doCommand(CMD);
        }

        public void CH2_ContinuousMode()
        {
            string CMD = "SOURce2:BURSt:STATe 0";
            doCommand(CMD);
        }

        public void CHl1_HiLevel(double vol)
        {
            string HiLevelcmd = "SOURce1:VOLTage:LEVel:IMMediate:HIGH " + vol.ToString() + "V";
            doCommand(HiLevelcmd);
        }

        public void CH1_LoLevel(double vol)
        {
            string LoLevelCmd = "SOURce1:VOLTage:LEVel:IMMediate:LOW " + vol.ToString() + "V";
            doCommand(LoLevelCmd);
        }

        public void CHl2_HiLevel(double vol)
        {
            string HiLevelcmd = "SOURce2:VOLTage:LEVel:IMMediate:HIGH " + vol.ToString() + "V";
            doCommand(HiLevelcmd);
        }

        public void CH2_LoLevel(double vol)
        {
            string LoLevelCmd = "SOURce2:VOLTage:LEVel:IMMediate:LOW " + vol.ToString() + "V";
            doCommand(LoLevelCmd);
        }

        public void CH1_Frequency(double freq)
        {
            string Cmd = "SOURCE1:FREQ " + freq.ToString();
            doCommand(Cmd);
        }

        public void CH2_Frequency(double freq)
        {
            string Cmd = "SOURCE2:FREQ " + freq.ToString();
            doCommand(Cmd);
        }

        public void CH1_Normal()
        {
            string Cmd = "OUTPUT1:POLARITY NORMAL";
            doCommand(Cmd);
        }

        public void CH2_Normal()
        {
            string Cmd = "OUTPUT2:POLARITY NORMAL";
            doCommand(Cmd);
        }

        public void CH1_Inverte()
        {
            string Cmd = "OUTPUT1:POLARITY INVERTED";
            doCommand(Cmd);
        }

        public void CH2_Inverte()
        {
            string Cmd = "OUTPUT2:POLARITY INVERTED";
            doCommand(Cmd);
        }

        public void CH1_On()
        {
            string Cmd = "OUTPut1:STATe ON";
            doCommand(Cmd);
        }

        public void CH1_Off()
        {
            string Cmd = "OUTPut1:STATe OFF";
            doCommand(Cmd);
        }

        public void CH2_On()
        {
            string Cmd = "OUTPut2:STATe ON";
            doCommand(Cmd);
        }

        public void CH2_Off()
        {
            string Cmd = "OUTPut2:STATe OFF";
            doCommand(Cmd);
        }

        public void CH1_DutyCycle(double duty)
        {
            string Cmd = "SOURCE1:PULS:DCYC " + String.Format("{0:0.####}", duty);
            doCommand(Cmd);
            CH1_On();
        }

        public void CH2_DutyCycle(double duty)
        {
            string Cmd = "SOURCE2:PULS:DCYC " + String.Format("{0:0.####}", duty);
            doCommand(Cmd);
            CH2_On();
        }

        public void SetTimerTrigMode()
        {
            string Cmd = "TRIGGER:SEQUENCE:SOURCE TIMER";
            doCommand(Cmd);
        }

        public void SetExternalTrigMode()
        {
            string Cmd = "TRIGGER:SEQUENCE:SOURCE EXTERNAL";
            doCommand(Cmd);
        }

        public void SetTrigForce()
        {
            string Cmd = "TRIGGER:SEQUENCE:IMMEDIATE";
            doCommand(Cmd);
        }

        public void CH1_LoadImpedance(double val)
        {
            string Cmd = "OUTPut1:IMPedance " + val.ToString();
            doCommand(Cmd);
        }

        public void CH2_LoadImpedance(double val)
        {
            string Cmd = "OUTPut2:IMPedance " + val.ToString();
            doCommand(Cmd);
        }

        public void CH1_LoadImpedanceHiz()
        {
 
            string Cmd = "OUTPut1:IMPedance INFinity";
            doCommand(Cmd);
        }

        public void CH2_LoadImpedanceHiz()
        {
            string Cmd = "OUTPut2:IMPedance INFinity";
            doCommand(Cmd);
        }

        public void CH1_PulseMode()
        {
            string Cmd = "SOURCE1:FUNC PULS";
            doCommand(Cmd);
        }

        public void CH2_PulseMode()
        {
            string Cmd = "SOURCE2:FUNC PULS";
            doCommand(Cmd);
        }

        public void CH1_BurstMode(bool IsOneCycle = true)
        {
            string Cmd = "SOURce1:BURSt:STATe 1";
            doCommand(Cmd);
            if(IsOneCycle)
            {
                doCommand("SOURce1:BURSt:NCYCles 1");
            }
        }
        public void CH2_BurstMode()
        {
            string Cmd = "OUTPut2:FUNC BURST";
            doCommand(Cmd);
        }
        public void SetCH1_TrTfFunc(double tr, double tf)
        {
            string Cmd = "SOURCE1:PULS:TRAN:LEAD " + (tr * 0.000001).ToString();
            doCommand(Cmd);
            Cmd = "SOURCE1:PULS:TRAN:TRA " + (tf * 0.000001).ToString();
            doCommand(Cmd);
        }
        public double GetCHl1_HiLevel()
        {
            string HiLevelcmd = "SOURce1:VOLTage:LEVel:IMMediate:HIGH?";
            return doQueryNumber(HiLevelcmd);
        }

        public double GetCH1_LoLevel()
        {
            string LoLevelCmd = "SOURce1:VOLTage:LEVel:IMMediate:LOW?";
            return doQueryNumber(LoLevelCmd);
        }
        public void DoCommand(string CMD)
        {
            doCommand(CMD);
        }
    }
}
