using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsLibDotNet
{
    class OscilloscopesModule : VisaCommand
    {
        public bool tektronix_en;

        public OscilloscopesModule()
        {

        }

        public OscilloscopesModule(string Addr)
        {
            LinkingIns(Addr);
            if(doQueryIDN().Split(',')[0].IndexOf("TEKTRONIX") != -1)
            {
                tektronix_en = false;
            }
            else
            {
                tektronix_en = true;
            }
        }

        ~OscilloscopesModule()
        {
            InsClose();
        }

        public void SetRST()
        {
            doCommand("*RST");
        }

        public void SetStop()
        {
            if (tektronix_en)
                doCommand("ACQuire:STATE STOP");
            else
                doCommand(":STOP");
        }

        public void SetClear()
        {
            if (tektronix_en)
                doCommand("DISplay:PERSistence:RESET");
            else
                doCommand(":CDISplay");
        }

        public void SetSingle()
        {
            if (tektronix_en)
                doCommand(":SINGle");
            else
                doCommand("ACQuire:STOPAfter SEQuence");
        }


        public void SetAutoTrigger()
        {
            if (tektronix_en)
                doCommand("TRIGger:A:MODe AUTO");
            else
                doCommand(":TRIGger:SWEep AUTO");
        }

        public void SetNormalTrigger()
        {
            if (tektronix_en)
                doCommand("TRIGger:A:MODe NORMal");
            else
                doCommand(":TRIGger:SWEep TRIGgered");
        }

        public void SetTriggerRise()
        {
            if (tektronix_en)
                doCommand("TRIGger:A:EDGE:SLOpe RISE");
            else
                doCommand(":TRIGger:EDGE:SLOPe POSitive");
        }

        public void SetTriggerFall()
        {
            if (tektronix_en)
                doCommand("TRIGger:A:EDGE:SLOpe FALL");
            else
                doCommand(":TRIGger:EDGE:SLOPe NEGative");
        }

        public void CHx_BWLimitOn(int Ch)
        {
            if (tektronix_en)
                doCommand(string.Format(":CH{0}:BWLimit 20e6", Ch));
            else
                doCommand(string.Format("CH{0}:BANdwidth TWEnty", Ch));
        }

    }
}
