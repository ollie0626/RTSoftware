using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class PowerModule : VisaCommand
    {
        private bool E3632Sel;
        private int E3631Sel;
        private bool E3633Sel;

        ~PowerModule()
        {
            InsClose();
        }

        public PowerModule()
        {
            LinkingIns("GPIB0::5::INSTR");
        }

        public PowerModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public PowerModule(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void ConnectPowerSupply(string Addr)
        {
            LinkingIns(Addr);
        }

        public void ConnectPowerSupply(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void PowerOff()
        {
            string poweroff = "OUTP OFF";
            doCommand(poweroff);
        }

        public void PowerOn()
        {
            string cmd = "OUT ON";
            doCommand(cmd);
        }


        public void E3633_Sel(bool Sel)
        {
            E3633Sel = Sel;
            string range = "";
            if (Sel)
            {
                range = "VOLTage:RANGe P8V";
                doCommand(range);
            }
            else if (!Sel)
            {
                range = "VOLTage:RANGe P20V";
                doCommand(range);
            }
        }

        public void E3633_Vol(double vol)
        {
            if (E3633Sel) if (vol > 8) vol = 8;
            else if (vol > 20) vol = 20;
            string Voltage = "APPLy " + String.Format("{0:0.###}", vol);
            string poweron = "OUTPut:STATE ON";
            doCommand(Voltage);
            doCommand(poweron);
        }

        public void ChromaVinVoltage(double vol)
        {
            string gogoCMD = "SOURce:VOLTage " + vol.ToString();
            doCommand(gogoCMD);
            gogoCMD = "CONFigure:OUTPut ON";
            doCommand(gogoCMD);
        }

        public void ChromaPowerOff()
        {
            string gogoCMD = "CONFigure:OUTPut OFF";
            doCommand(gogoCMD);
        }

        public void ChromaCurrentLimit(double curr)
        {
            string gogoCMD = "SOURce:CURRent " + curr.ToString();
            doCommand(gogoCMD);
        }

        public double GetCurrentP25()
        {
            string measureCur = "MEAS:CURR:DC? P25V";
            return doQueryNumber(measureCur);
        }
        public double GetCurrentN25()
        {
            string measureCur = "MEAS:CURR:DC? N25V";
            return doQueryNumber(measureCur);
        }

        public double GetVoltageN25()
        {
            string measureVol = "MEAS:VOLT:DC? N25V";
            return doQueryNumber(measureVol);
        }

        public double GetCurrentP6()
        {          
            string measureVol = "MEAS:CURR:DC? P6V";
            return doQueryNumber(measureVol);
        }

        public double GetCurrent()
        {      
            string measureCur = "MEAS:CURR?";
            return doQueryNumber(measureCur);
        }

        public double GetVoltageP6()
        {
            
            string measureVol = "MEAS:VOLT:DC? P6V";
            return doQueryNumber(measureVol);
        }

        public double GetVoltageP25()
        {
            
            string measureVol = "MEAS:VOLT:DC? P25V";
            return doQueryNumber(measureVol);
        }


        public void E3631_Sel(int RangeSel)
        {
            E3631Sel = RangeSel;
            if (RangeSel == 0)
            {
                string vinrange = "INST:SEL P6V";
                doCommand(vinrange);
            }
            else if (RangeSel == 1)
            {
                string vinrange = "INST:SEL P25V";
                doCommand(vinrange);
            }
            else if (RangeSel == 2)
            {
                string vinrange = "INST:SEL N25V";
                doCommand(vinrange);
            }
        }

        public void E3631_Vol(double vol)
        {
            if (E3631Sel == 0)
            {
                if (vol > 6) vol = 6;
            }
            else if (E3631Sel == 1)
            {
                if (vol > 25) vol = 25;
            }
            else if (E3631Sel == 2)
            {
                if (vol < -25) vol = -25;
            }
            string Voltage = "VOLT " + String.Format("{0:0.###}", vol);
            string poweron = "OUTP ON";
            doCommand(Voltage);
            doCommand(poweron);
        }


        public double GetVoltage()
        {
            string measureVol = "MEAS:VOLT?";
            return doQueryNumber(measureVol);
        }

        public void E3633PowerOff()
        {
            string gogoCmd = "OUTPut:STATE OFF";
            doCommand(gogoCmd);
        }

        public void E3632_Sel(bool p15orp30 = true)
        {
            E3632Sel = p15orp30;
            if (p15orp30)
            {
                string P15VRange = "VOLT:RANG P15V";
                string P15VSel = "INST:SEL P15V";
                doCommand(P15VRange);
                doCommand(P15VSel);
            }
            else
            {
                string P30VRange = "VOLT:RANG P30V";
                string P30VSel = "INST:SEL P30V";
                doCommand(P30VRange);
                doCommand(P30VSel);
            }
        }

        public void E3632_Vol(double vol)
        {
            if (E3632Sel)
            {
                if (vol > 15) vol = 15;
            }
            else if (!E3632Sel)
            {
                if (vol > 30) vol = 30;
            }

            string Voltage = "VOLT " + String.Format("{0:0.###}", vol);
            string poweron = "OUTP ON";
            doCommand(Voltage);
            doCommand(poweron);
        }

        public void AutoPowerOff()
        {
            string IDN = doQueryIDN();
            if (IDN.IndexOf("E3632") != -1 || IDN.IndexOf("E3631") != -1)
                PowerOff();
            else if (IDN.IndexOf("E3633") != -1)
                E3633PowerOff();
            else if (IDN.IndexOf("620") != -1)
                ChromaPowerOff();
        }


        public void AutoSelPowerOn(double vol)
        {
            string IDN = doQueryIDN();
            if (IDN.IndexOf("E3632") != -1)
            {
                bool sel = vol < 15 ? true : false;
                E3632_Sel(sel);
                E3632_Vol(vol);
            }
            else if (IDN.IndexOf("E3633") != -1)
            {
                bool sel = vol < 8 ? true : false;
                E3633_Sel(sel);
                E3633_Vol(vol);
            }
            else if (IDN.IndexOf("E3631") != -1)
            {
                int sel = 0;
                if (vol < 6 && vol > 0)
                    sel = 0;
                else if (vol > 6 && vol < 25)
                    sel = 1;
                else if (vol < 0)
                    sel = 2;

                E3631_Sel(sel);
                E3631_Vol(vol);
            }
            else if (IDN.IndexOf("620") != -1)
            {
                ChromaVinVoltage(vol);
            }
        }


        public void AutoSetOCP(double ocp)
        {
            string IDN = doQueryIDN();
            if (IDN.IndexOf("620") != -1)
                ChromaCurrentLimit(ocp);
            else
            {
                string cmd = "Curr MAX";// + string.Format("{0:##.##}", ocp);
                doCommand(cmd);
            }
        }


    }
}
