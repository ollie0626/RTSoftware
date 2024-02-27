using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    public class ChamberModule : VisaCommand
    {
        /// <summary>
        /// Chamber initial GPIB address is 3.
        /// </summary>
        public ChamberModule()
        {
            LinkingIns("GPIB0::3::INSTR");
        }

        ~ChamberModule()
        {
            InsClose();
        }

        public override void LinkingIns(string Addr)
        {
            if (_IsDebug == true || Addr == "")
            {
                device = 0;
                return;
            }
            if (Rm == 0) visa32.viOpenDefaultRM(out Rm);
            visa32.viOpen(Rm, Addr, 0, 0, out device);
            Console.WriteLine(Addr + "   " + device);
            visa32.viSetAttribute(device, visa32.VI_ATTR_TMO_VALUE, 100);
        }

        public override bool InsState()
        {
            string cmd = "AT";
            //string buf;
            try
            {
                string gogoCMD = "AT";
                byte[] buffer = new byte[1024];
                string str;
                int count_out;
                int count_in = 1024;
                //Reapt:;

                Array.Clear(buffer, 0, buffer.Length);
                visa32.viPrintf(device, gogoCMD + "\r\n");
                visa32.viRead(device, buffer, count_in, out count_out);
                str = Encoding.ASCII.GetString(buffer, 0, count_out);

                //if (str == "")
                //    goto Reapt;

                str = str.Replace(">", ".");
                int idx = str.IndexOf('.') + 3;
                str = str.Substring(0, idx);
                str = str.Replace("=", "-");
                double data = Convert.ToDouble(str);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("stackTrace: " + ex.StackTrace);
                Console.WriteLine("Message: " + ex.Message);
                return false;
            }
        }


        public ChamberModule(string Addr)
        {
            LinkingIns(Addr);
        }

        public ChamberModule(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void ConnectChamber(string Addr)
        {
            LinkingIns(Addr);
        }

        public void ConnectChamber(int Addr)
        {
            LinkingIns("GPIB0::" + Addr.ToString() + "::INSTR");
        }

        public void ChamberOn(double temp)
        {
            string gogoCMD;
            gogoCMD = "T " + temp.ToString() + ", 0";
            doCommand(gogoCMD);
            doCommand(gogoCMD);
        }

        public void ChamberOff()
        {
            string gogoCMD = "KT\r\n";
            doCommand(gogoCMD);
            gogoCMD = "<command>,\r\n";
            doCommand(gogoCMD);
        }

        public double GetChamberTemperature()
        {
            string gogoCMD = "AT";
            byte[] buffer = new byte[1024];
            string str = "";
            int count_out;
            int count_in = 1024;
            double data = 0;
            Reapt:;
            try
            {
                Array.Clear(buffer, 0, buffer.Length);
                visa32.viPrintf(device, gogoCMD + "\r\n");
                visa32.viRead(device, buffer, count_in, out count_out);
                str = Encoding.ASCII.GetString(buffer, 0, count_out);
                if (str == "")
                    goto Reapt;

                str = str.Replace(">", ".");
                int idx = str.IndexOf('.') + 3;
                str = str.Substring(0, idx);
                str = str.Replace("=", "-");
                data = Convert.ToDouble(str);
            }
            catch
            {
                goto Reapt;
            }
            return data;
        }


        public void ChamberStableCheck(double temp)
        {
            double tempNow = GetChamberTemperature();
            double tempdown = temp - 1;
            double tempup = temp + 1;
            if (tempNow > tempup || tempNow < tempdown)
            {
                ChamberStableCheck(temp);
            }
        }
    }
}
