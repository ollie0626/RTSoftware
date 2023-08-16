using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RT6971
{
    public partial class Form1 : Form
    {
        System.EventHandler[] eventHandlers;

        public Form1()
        {
            InitializeComponent();
            eventHandlers = new EventHandler[]
            {
                GAM1H_ValueChanged, GAM2H_ValueChanged, GAM3H_ValueChanged, GAM4H_ValueChanged, GAM5H_ValueChanged, GAM6H_ValueChanged,
                GAM7H_ValueChanged, GAM8H_ValueChanged, GAM9H_ValueChanged, GAM10H_ValueChanged, GAM11H_ValueChanged, GAM12H_ValueChanged,
                GAM13H_ValueChanged, GAM14H_ValueChanged, VCOM1H_ValueChanged, VCOM2H_ValueChanged, VCOM3H_ValueChanged
            };
        }

        private double GAMOut_Calculate(int code)
        {
            double res = (double)GLDOV.Value / 1024;
            double GAMout = res * code;
            return GAMout;
        }

        private double VCOMout_Calculate(int code)
        {
            double res = (double)GLDOV.Value / 1024;
            double VCOMout = res * code;
            return VCOMout;
        }

        private void GAM_assign(int code, NumericUpDown MSB, NumericUpDown LSB)
        {
            byte bit9_8 = (byte)((code & 0x300) >> 8);
            byte bit7_0 = (byte)(code & 0xff);

            MSB.Value = bit9_8 | ((int)MSB.Value & 0xFC);
            LSB.Value = bit7_0;
        }

        private void AVDDH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)AVDDH.Value;
            double vol = 13.8;
            if (code > 0)
            {
                vol = (((double)(code - 1) * 1) + 138) / 10;

            }
            AVDDSL.Value = (int)AVDDH.Value;
            AVDDV.Value = (decimal)vol;
            W00.Value = (int)AVDDH.Value | ((int)W00.Value & 0xC0);
        }

        private void AVDDSL_Scroll(object sender, ScrollEventArgs e)
        {
            AVDDH.Value = AVDDSL.Value;
        }

        private void VCC1H_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VCC1H.Value;
            double vol = 0.8;

            vol = (double)(code * 2 + 80) / 100;
            if (vol > 2.36) vol = 2.36;

            VCC1SL.Value = code;
            VCC1V.Value = (decimal)vol;
            W01.Value = (int)VCC1H.Value | ((int)W01.Value & 0x80);
        }

        private void VCC1SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCC1H.Value = VCC1SL.Value;
        }

        private void VCC2H_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VCC2H.Value;
            double vol = 2.2;
            vol = (double)(code * 1 + 22) / 10;
            VCC2V.Value = (decimal)vol;

            VCC2SL.Value = (int)VCC2H.Value;
            W02.Value = code | ((int)W02.Value & 0x80);
        }

        private void VCC2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCC2H.Value = VCC2H.Value;
        }

        private void VGHLTH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VGHLTH.Value;
            double vol = 21;

            vol = (double)(code * 2 + 210) / 10;
            if (vol > 45) vol = 45;
            VGHLTV.Value = (decimal)vol;
            VGHLTSL.Value = (int)VGHLTH.Value;

            W03.Value = code | (int)W03.Value & 0x80;
        }

        private void VGHLTSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGHLTH.Value = (int)VGHLTSL.Value;
        }

        private void VGHHTH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VGHHTH.Value;
            double vol = 35.6;

            vol = (double)(code * 2 + 200) / 10;
            if (vol > 44) vol = 44;
            VGHHTV.Value = (decimal)vol;
            VGHHTSL.Value = code;

            W04.Value = code | ((int)W04.Value & 0x80);
        }

        private void VGHHTSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGHHTH.Value = (int)VGHHTSL.Value;
        }

        private void VGL1H_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VGL1H.Value;
            double vol = ((double)code * 1 + 18) / -10;
            if (vol < -15) vol = -15;

            VGL1V.Value = (decimal)vol;
            VGL1SL.Value = (int)VGL1H.Value;
            W05.Value = code | (int)W05.Value & 0x80;
        }

        private void VGL1SL_Scroll(object sender, ScrollEventArgs e)
        {
            VGL1H.Value = VGL1SL.Value;
        }

        private void VGL2LTH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VGL2LTH.Value;
            double vol = ((double)code * 2 + 125) / -10;

            if (vol < -20) vol = -20;

            VGL2LTV.Value = (decimal)vol;
            VGL2LTSL.Value = (int)VGL2LTH.Value;
            W06.Value = code | (int)W06.Value & 0x80;
        }

        private void VGL2HTH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)VGL2HTH.Value;
            double vol = ((double)code * 2 + 125) / -10;

            if (vol < -20) vol = -20;

            VGL2HTV.Value = (decimal)vol;
            VGL2HTSL.Value = (int)VGL2HTH.Value;
            W07.Value = code | (int)W07.Value & 0x80;
        }

        private void VGL2HTSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGL2LTH.Value = VGL2LTSL.Value;
        }

        private void GLDOH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)GLDOH.Value;
            double vol = ((double)code * 2 + 130) / 10;
            if (vol > 18) vol = 18;

            GLDOV.Value = (decimal)vol;
            GLDOSL.Value = code;
            W09.Value = code << 3 | (int)W09.Value & 0x07;

            for(int i = 0; i < eventHandlers.Length; i++) eventHandlers[i](null, null);
        }

        private void GLDOSL_Scroll(object sender, ScrollEventArgs e)
        {
            GLDOH.Value = GLDOSL.Value;
        }

        private void HAVDDH_ValueChanged(object sender, EventArgs e)
        {
            HAVDDSL.Value = (int)HAVDDH.Value;
            HAVDDV.Value = (decimal)VCOMout_Calculate((int)HAVDDH.Value);

            byte bit9_8 = (byte)(((int)HAVDDH.Value & 0x300) >> 8);
            byte bit7_0 = (byte)((int)HAVDDH.Value & 0xFF);

            W09.Value = bit9_8 | (int)W09.Value & 0xFC;
            W0A.Value = bit7_0;
        }

        private void HAVDDSL_Scroll(object sender, ScrollEventArgs e)
        {
            HAVDDH.Value = HAVDDSL.Value;
        }

        private void GAM1H_ValueChanged(object sender, EventArgs e)
        {
            GAM1SL.Value = (int)GAM1H.Value;
            GAM1V.Value = (decimal)GAMOut_Calculate((int)GAM1H.Value);

            GAM_assign((int)GAM1H.Value, W0B, W0C);
        }

        private void GAM2H_ValueChanged(object sender, EventArgs e)
        {
            GAM2SL.Value = (int)GAM2H.Value;
            GAM2V.Value = (decimal)GAMOut_Calculate((int)GAM2H.Value);
            GAM_assign((int)GAM2H.Value, W0D, W0E);
        }

        private void GAM3H_ValueChanged(object sender, EventArgs e)
        {
            GAM3SL.Value = (int)GAM3H.Value;
            GAM3V.Value = (decimal)GAMOut_Calculate((int)GAM3H.Value);
            GAM_assign((int)GAM3H.Value, W0F, W10);
        }

        private void GAM4H_ValueChanged(object sender, EventArgs e)
        {
            GAM4SL.Value = (int)GAM4H.Value;
            GAM4V.Value = (decimal)GAMOut_Calculate((int)GAM4H.Value);
            GAM_assign((int)GAM4H.Value, W11, W12);
        }

        private void GAM5H_ValueChanged(object sender, EventArgs e)
        {
            GAM5SL.Value = (int)GAM5H.Value;
            GAM5V.Value = (decimal)GAMOut_Calculate((int)GAM5H.Value);
            GAM_assign((int)GAM5H.Value, W13, W14);
        }

        private void GAM6H_ValueChanged(object sender, EventArgs e)
        {
            GAM6SL.Value = (int)GAM6H.Value;
            GAM6V.Value = (decimal)GAMOut_Calculate((int)GAM6H.Value);
            GAM_assign((int)GAM6H.Value, W15, W16);
        }

        private void GAM7H_ValueChanged(object sender, EventArgs e)
        {
            GAM7SL.Value = (int)GAM7H.Value;
            GAM7V.Value = (decimal)GAMOut_Calculate((int)GAM7H.Value);
            GAM_assign((int)GAM7H.Value, W17, W18);
        }

        private void GAM8H_ValueChanged(object sender, EventArgs e)
        {
            GAM8SL.Value = (int)GAM8H.Value;
            GAM8V.Value = (decimal)GAMOut_Calculate((int)GAM8H.Value);
            GAM_assign((int)GAM7H.Value, W19, W1A);
        }

        private void GAM9H_ValueChanged(object sender, EventArgs e)
        {
            GAM9SL.Value = (int)GAM9H.Value;
            GAM9V.Value = (decimal)GAMOut_Calculate((int)GAM9H.Value);
            GAM_assign((int)GAM7H.Value, W1B, W1C);
        }

        private void GAM10H_ValueChanged(object sender, EventArgs e)
        {
            GAM10SL.Value = (int)GAM10H.Value;
            GAM10V.Value = (decimal)GAMOut_Calculate((int)GAM10H.Value);
        }

        private void GAM11H_ValueChanged(object sender, EventArgs e)
        {
            GAM11SL.Value = (int)GAM11H.Value;
            GAM11V.Value = (decimal)GAMOut_Calculate((int)GAM11H.Value);
        }

        private void GAM12H_ValueChanged(object sender, EventArgs e)
        {
            GAM12SL.Value = (int)GAM12H.Value;
            GAM12V.Value = (decimal)GAMOut_Calculate((int)GAM12H.Value);
            GAM_assign((int)GAM12H.Value, W32, W33);
        }

        private void GAM13H_ValueChanged(object sender, EventArgs e)
        {
            GAM13SL.Value = (int)GAM13H.Value;
            GAM13V.Value = (decimal)GAMOut_Calculate((int)GAM13H.Value);
            GAM_assign((int)GAM13H.Value, W34, W35);
        }

        private void GAM14H_ValueChanged(object sender, EventArgs e)
        {
            GAM14SL.Value = (int)GAM14H.Value;
            GAM14V.Value = (decimal)GAMOut_Calculate((int)GAM14H.Value);
            GAM_assign((int)GAM14H.Value, W36, W37);
        }

        private void VCOM1H_ValueChanged(object sender, EventArgs e)
        {
            VCOM1SL.Value = (int)VCOM1H.Value;
            VCOM1V.Value = (decimal)VCOMout_Calculate((int)VCOM1H.Value);
            GAM_assign((int)VCOM1H.Value, W38, W39);
        }

        private void VCOM2H_ValueChanged(object sender, EventArgs e)
        {
            VCOM2SL.Value = (int)VCOM2H.Value;
            VCOM2V.Value = (decimal)VCOMout_Calculate((int)VCOM2H.Value);
            GAM_assign((int)VCOM2H.Value, W3A, W3B);
        }

        private void VCOM3H_ValueChanged(object sender, EventArgs e)
        {
            VCOM3SL.Value = (int)VCOM3H.Value;
            VCOM3V.Value = (decimal)VCOMout_Calculate((int)VCOM3H.Value);
            GAM_assign((int)VCOM3H.Value, W3C, W3D);
        }

        private void GAM1SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM1H.Value = GAM1SL.Value;
        }

        private void GAM2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM2H.Value = GAM2SL.Value;
        }

        private void GAM3SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM3H.Value = GAM3SL.Value;
        }

        private void GAM4SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM4H.Value = GAM4SL.Value;
        }

        private void GAM5SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM5H.Value = GAM5SL.Value;
        }

        private void GAM6SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM6H.Value = GAM6SL.Value;
        }

        private void GAM7SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM7H.Value = GAM7SL.Value;
        }

        private void GAM8SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM8H.Value = GAM7SL.Value;
        }

        private void GAM9SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM9H.Value = GAM9SL.Value;
        }

        private void GAM10SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM10H.Value = GAM10SL.Value;
        }

        private void GAM11SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM11H.Value = GAM11SL.Value;
        }

        private void GAM12SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM12H.Value = GAM12SL.Value;
        }

        private void GAM13SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM13H.Value = GAM13SL.Value;
        }

        private void GAM14SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM14H.Value = GAM14SL.Value;
        }

        private void VCOM1SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM1H.Value = VCOM1SL.Value;
        }

        private void VCOM2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM2H.Value = VCOM2SL.Value;
        }

        private void VCOM3SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM3H.Value = VCOM3SL.Value;
        }
    }
}
