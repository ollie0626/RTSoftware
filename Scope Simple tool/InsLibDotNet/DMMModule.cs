using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class DMMModule : VisaCommand
    {
        ~DMMModule()
        {
            InsClose();
        }

        /// <summary>
        /// Tektronix DMM Meter Addr 
        /// </summary>
        public DMMModule()
        {
            LinkingIns("GPIB0::10::INSTR");
        }

        public DMMModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public DMMModule(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }
        public void ConnectDMM(string Addr)
        {
            LinkingIns(Addr);
        }
        public void ConnectDMM(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }
        public void AFilterOn()
        {
            string filter = "FILT:DC:STAT ON";
            doCommand(filter);
        }
        public void DFilterOn()
        {
            string filter = "FILT:DC:DIG ON";
            doCommand(filter);
        }
        public void AFilterOff()
        {
            string filter = "FILT:DC:STAT OFF";
            doCommand( filter);
        }
        public void DFilterOff()
        {
            string filter = "FILT:DC:DIG OFF";
            doCommand(filter);
        }

        public double GetVoltage(int level = 4)
        {
            string MeasVol = "";
            switch (level)
            {
                case 0:
                    MeasVol = "MEAS:VOLT:DC? 1e-1";
                    break;
                case 1:
                    MeasVol = "MEAS:VOLT:DC? 1";
                    break;
                case 2:
                    MeasVol = "MEAS:VOLT:DC? 10";
                    break;
                case 3:
                    MeasVol = "MEAS:VOLT:DC? 100";
                    break;
                default:
                    MeasVol = "MEAS:VOLT:DC?";
                    break;
            }
            return doQueryNumber(MeasVol);
        }

        public double GetCurrent(int level, double customer = 0.4)
        {
            string MeasCur = "";
            switch (level)
            {
                case 0:
                    MeasCur = "MEAS:CURR:DC? 1e-4";
                    break;
                case 1:
                    MeasCur = "MEAS:CURR:DC? 4e-1";
                    break;
                case 2:
                    MeasCur = "MEAS:CURR:DC? " + String.Format("{0:0.####}", customer);
                    break;
                case 3:
                    MeasCur = "MEAS:CURR:DC? 10";
                    break;
                default:
                    MeasCur = "MEAS:CURR:DC?";
                    break;
            }
            return doQueryNumber(MeasCur) * 1000;
        }
    }
}
