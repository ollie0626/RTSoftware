using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class MultiChannelModule : VisaCommand
    {
        ~MultiChannelModule()
        {
            InsClose();
        }

        public MultiChannelModule()
        {
            LinkingIns("GPIB0::6::INSTR");
        }

        public MultiChannelModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public MultiChannelModule(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void ConnectMultiChannel(string Addr)
        {
            LinkingIns(Addr);
        }

        public void ConnectMultiChannel(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void Init34970A()
        {
            ChannelDelaySet(0.5, 1, 19);
        }

        public void ChannelDelaySet(double tick, int form, int to)
        {
            string buf = string.Format("ROUT:CHAN:DELAY " + tick.ToString() + ",(@1{0:00}:" + "1{1:00})", form, to);
            doCommand(buf);
        }

        public void ScanList(bool[] ck)
        {
            string buf = "ROUT:SCAN (@";
            for (int i = 0; i < 20; i++)
            {
                if (ck[i])
                {
                    buf += string.Format("1{0:00},", i + 1);
                }
            }
            buf = buf.Substring(0, buf.Length - 1);
            buf += ")";
            doCommand(buf);
        }

        public double[] QuickMeasure(double rang, bool[] ck)
        {
            string buf = "MEASure:VOLTage:DC? " + rang.ToString() + ",(@";
            for (int i = 0; i < 20; i++)
            {
                if (ck[i])
                {
                    buf += string.Format("1{0:00},", i + 1);
                }
            }
            buf = buf.Substring(0, buf.Length - 1);
            buf += ")";
            double[] arr = new double[20];
            doQueryNumbers(buf, ref arr);
            return arr;
        }

        public double[] QuickMEasureDefine(double level, List<string> CHx_num)
        {
            string buf = "MEASure:VOLTage:DC? " + level.ToString() + ",1E-3,(@";
            for(int i = 0; i < CHx_num.Count; i++)
            {
                buf += CHx_num[i] + ",";
            }
            buf = buf.Substring(0, buf.Length - 1);
            buf += ")";

            double[] arr = new double[20];
            doQueryNumbers(buf, ref arr);
            return arr;
        }


        public double Get_100mVol(int channel)
        {
            string MeaVol = "";
            if (channel < 10)
                MeaVol = "MEAS:VOLT:DC? 0.1,1E-6,(@10" + channel.ToString() + ")";
            else if (channel >= 10 && channel < 20)
                MeaVol = "MEAS:VOLT:DC? 0.1,1E-6,(@1" + channel.ToString() + ")";
            else if (channel >= 20)
                MeaVol = "MEAS:VOLT:DC? 0.1,1E-6,(@2" + channel.ToString() + ")";
            return doQueryNumber(MeaVol);
        }

        public double Get_1Vol(int channel)
        {
            string MeaVol = "";
            if (channel < 10)
                MeaVol = "MEAS:VOLT:DC? 1,1E-6,(@10" + channel.ToString() + ")";
            else if (channel >= 10 && channel < 20)
                MeaVol = "MEAS:VOLT:DC? 1,1E-6,(@1" + channel.ToString() + ")";
            else if (channel >= 20)
                MeaVol = "MEAS:VOLT:DC? 1,1E-6,(@2" + channel.ToString() + ")";
            return doQueryNumber(MeaVol);
        }


        public double Get_10Vol(int channel)
        {
            string MeaVol = "";
            if (channel < 10)
                MeaVol = "MEAS:VOLT:DC? 10,1E-5,(@10" + channel.ToString() + ")";
            else if (channel >= 10 && channel < 20)
                MeaVol = "MEAS:VOLT:DC? 10,1E-5,(@1" + channel.ToString() + ")";
            else if (channel >= 20)
                MeaVol = "MEAS:VOLT:DC? 10,1E-5,(@2" + channel.ToString() + ")";
            return doQueryNumber(MeaVol);
        }

        public double Get_100Vol(int channel)
        {
            string MeaVol = "";
            if (channel < 10)
                MeaVol = "MEAS:VOLT:DC? 100,1E-3,(@10" + channel.ToString() + ")";
            else if (channel >= 10 && channel < 20)
                MeaVol = "MEAS:VOLT:DC? 100,1E-3,(@1" + channel.ToString() + ")";
            else if (channel >= 20)
                MeaVol = "MEAS:VOLT:DC? 100,1E-3,(@2" + channel.ToString() + ")";
            return doQueryNumber(MeaVol);
        }
    }
}
