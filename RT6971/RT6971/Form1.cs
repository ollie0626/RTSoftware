using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RT6971
{
    public partial class Form1 : Form
    {
        System.EventHandler[] eventHandlers;

        NumericUpDown[] WriteTable;
        NumericUpDown[] ReadTable;
        RTBBControl RTDev = new RTBBControl();



        public Form1()
        {
            InitializeComponent();
            eventHandlers = new EventHandler[]
            {
                GAM1H_ValueChanged, GAM2H_ValueChanged, GAM3H_ValueChanged, GAM4H_ValueChanged, GAM5H_ValueChanged, GAM6H_ValueChanged,
                GAM7H_ValueChanged, GAM8H_ValueChanged, GAM9H_ValueChanged, GAM10H_ValueChanged, GAM11H_ValueChanged, GAM12H_ValueChanged,
                GAM13H_ValueChanged, GAM14H_ValueChanged, VCOM1H_ValueChanged, VCOM2H_ValueChanged, VCOM3H_ValueChanged
            };

            WriteTable = new NumericUpDown[]
            {
                W00, W01, W02, W03, W04, W05, W06, W07, W08, W09, W0A, W0B, W0C, W0D, W0E, W0F,
                W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W1A, W1B, W1C, W1D, W1E, W1F,
                W20, W21, W22, W23, W24, W25, W26, W27, W28, W29, W2A, W2B, W2C, W2D, W2E, W2F,
                W30, W31, W32, W33, W34, W35, W36, W37, W38, W39, W3A, W3B, W3C, W3D, W3E, W3F,
                W40, W41, W42, W43, W44, W45, W46, W47, W48, W49, W4A, W4B, W4C, W4D, W4E, W4F,
            };

            ReadTable = new NumericUpDown[]
            {
                R00, R01, R02, R03, R04, R05, R06, R07, R08, R09, R0A, R0B, R0C, R0D, R0E, R0F,
                R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R1A, R1B, R1C, R1D, R1E, R1F,
                R20, R21, R22, R23, R24, R25, R26, R27, R28, R29, R2A, R2B, R2C, R2D, R2E, R2F,
                R30, R31, R32, R33, R34, R35, R36, R37, R38, R39, R3A, R3B, R3C, R3D, R3E, R3F,
                R40, R41, R42, R43, R44, R45, R46, R47, R48, R49, R4A, R4B, R4C, R4D, R4E, R4F,
            };


            RTDev.BoadInit();
            List<byte> list = RTDev.ScanSlaveID();
            if (list != null)
            {
                if (list.Count > 0)
                    nuSlave.Value = list[0];
            }

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
            int reserve = (int)MSB.Value;
            byte bit9_8 = (byte)((code & 0x300) >> 8);
            byte bit7_0 = (byte)(code & 0xff);

            MSB.Value = bit9_8 | (reserve & 0xFC);
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
            if (vol > 3.7) vol = 3.7;
            VCC2V.Value = (decimal)vol;

            VCC2SL.Value = (int)VCC2H.Value;
            W02.Value = code | ((int)W02.Value & 0x80);
        }

        private void VCC2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCC2H.Value = VCC2SL.Value;
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
            double vol = 20;

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
            VGL2HTH.Value = VGL2HTSL.Value;
        }

        private void GLDOH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)GLDOH.Value;
            double vol = ((double)code * 2 + 130) / 10;
            if (vol > 18) vol = 18;

            GLDOV.Value = (decimal)vol;
            GLDOSL.Value = code;
            W09.Value = code << 3 | (int)W09.Value & 0x07;

            for (int i = 0; i < eventHandlers.Length; i++) eventHandlers[i](null, null);
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
            GAM_assign((int)GAM8H.Value, W19, W1A);
        }

        private void GAM9H_ValueChanged(object sender, EventArgs e)
        {
            GAM9SL.Value = (int)GAM9H.Value;
            GAM9V.Value = (decimal)GAMOut_Calculate((int)GAM9H.Value);
            GAM_assign((int)GAM9H.Value, W1B, W1C);
        }

        private void GAM10H_ValueChanged(object sender, EventArgs e)
        {
            GAM10SL.Value = (int)GAM10H.Value;
            GAM10V.Value = (decimal)GAMOut_Calculate((int)GAM10H.Value);

            int code = (int)GAM10H.Value;
            int MSB = (code & 0x300) >> 8;
            int LSB = (code & 0xFF);
            W2F.Value = LSB;
            W30.Value = MSB << 6 | (int)W30.Value & 0x3F;

        }

        private void GAM11H_ValueChanged(object sender, EventArgs e)
        {
            GAM11SL.Value = (int)GAM11H.Value;
            GAM11V.Value = (decimal)GAMOut_Calculate((int)GAM11H.Value);

            int code = (int)GAM11H.Value;
            int MSB = (code & 0x300) >> 8;
            int LSB = (code & 0xFF);

            W30.Value = MSB | (int)W30.Value & 0xFC;
            W31.Value = LSB;
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
            GAM8H.Value = GAM8SL.Value;
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

        private void cb_protection_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                ComboBox[] cb_arr = new ComboBox[]
                {
                    cb_gam_en, cb_havdd_en, cb_vgh_en, cb_avdd_en, cb_vgl2_en, cb_vgl1_en, cb_vcc2_en, cb_protection
                };

                int data = 0x00;

                for (int i = 0; i < 8; i++) data |= (cb_arr[i].SelectedIndex << i);
                W29.Value = data;
            }
            catch
            {

            }

        }

        private void cb_avdd_dis_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox[] cb_arr = new ComboBox[]
                {
                    cb_vcom1_dis, cb_vcom2_dis, cb_vcom3_dis, cb_vcc_dis, cb_vgl1_dis, cb_vgh_dis, cb_havdd_dis, cb_havdd_dis
                };
                int data = 0x00;
                for (int i = 0; i < 8; i++) data |= (cb_arr[i].SelectedIndex << i);
                W24.Value = data;
            }
            catch
            {

            }




        }

        private void cb_vcom1_en_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ComboBox[] cb_arr = new ComboBox[]
                {
                    cb_vcom3_en, cb_vcom2_en, cb_vcom1_en
                };
                int data = 0x00;
                for (int i = 5; i < 8; i++) data |= cb_arr[i - 5].SelectedIndex << i;
                W26.Value = data | (int)W26.Value & 0x1F;
            }
            catch
            {

            }

        }

        private void cb_vcc1_ss_SelectedIndexChanged(object sender, EventArgs e)
        {
            W01.Value = cb_vcc1_ss.SelectedIndex << 7 | (int)W01.Value & 0x7F;
        }

        private void cb_vcc2_ss_SelectedIndexChanged(object sender, EventArgs e)
        {
            W02.Value = cb_vcc2_ss.SelectedIndex << 7 | (int)W02.Value & 0x7F;
        }

        private void cb_dly0_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_dly0.SelectedIndex == -1) return;
            if (cb_dly1.SelectedIndex == -1) return;
            if (cb_dly2.SelectedIndex == -1) return;
            if (cb_dly3.SelectedIndex == -1) return;

            W08.Value = (cb_dly0.SelectedIndex << 6) | cb_dly1.SelectedIndex << 4 | cb_dly2.SelectedIndex << 2 | cb_dly3.SelectedIndex << 0;

        }

        private void cb_vgh_tc_en_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vgh_tc_en.SelectedIndex == -1) return;
            if (cb_vcom_tc_en.SelectedIndex == -1) return;
            if (cb_tc_type.SelectedIndex == -1) return;

            W1D.Value = cb_vgh_tc_en.SelectedIndex << 7 | cb_vcom_tc_en.SelectedIndex << 6 | cb_tc_type.SelectedIndex << 5 | (int)W1D.Value & 0x1F;


            string[] TC_Mode1 = new string[]
            { 
                "0h : TCOMP_L = 2.94V, TCOMP_H = 2.09V",
                "1h : TCOMP_L = 3.35V, TCOMP_H = 2.95V",
                "2h : TCOMP_L = 2.94V, TCOMP_H = 2.44V",
                "3h : TCOMP_L = 2.94V, TCOMP_H = 2.64V"
            };

            string[] TC_Mode2 = new string[]
            {
                "0h : TCOMP_L1 = 3.79V, TCOMP_L = 2.94V, TCOMP_H = 2.09V, TCOMP_H1 = 1.24V",
                "1h : TCOMP_L1 = 3.75V, TCOMP_L = 3.35V, TCOMP_H = 2.95V, TCOMP_H1 = 2.55V",
                "2h : TCOMP_L1 = 3.44V, TCOMP_L = 2.94V, TCOMP_H = 2.44V, TCOMP_H1 = 1.94V",
                "3h : TCOMP_L1 = 3.24V, TCOMP_L = 2.94V, TCOMP_H = 2.64V, TCOMP_H1 = 2.34V"
            };
            cb_vgh_tc_mode.Items.Clear();
            switch (cb_tc_type.SelectedIndex)
            {
                case 0:
                    foreach(string item in TC_Mode1)
                        cb_vgh_tc_mode.Items.Add(item);
                    break;
                case 1:
                    foreach (string item in TC_Mode2)
                        cb_vgh_tc_mode.Items.Add(item);
                    break;
            }
        }

        private void cb_vgh_tc_mode_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combo = (ComboBox)sender;
            if (cb_vgh_tc_mode.SelectedIndex == -1) return;
            if (cb_vgx_prt_off.SelectedIndex == -1) return;
            if (cb_vcom_tc.SelectedIndex == -1) return;
            
            W1E.Value = cb_vgh_tc_mode.SelectedIndex << 6 | (int)W1E.Value & 0x3F;
            W1E.Value = cb_vgx_prt_off.SelectedIndex << 5 | (int)W1E.Value & 0xDF;
            W1E.Value = cb_vcom_tc.SelectedIndex | (int)W1E.Value & 0xF0;


            cb_vgh_tc_mode.Width = 121;
            groupBox31.Width = 165;
        }

        private void cb_eocp_time_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_eocp_time.SelectedIndex == -1) return;
            if (cb_gocp_time.SelectedIndex == -1) return;

            W1F.Value = cb_eocp_time.SelectedIndex << 4 | (int)W1F.Value & 0x0F;
            W1F.Value = cb_gocp_time.SelectedIndex | (int)W1F.Value & 0xF0;

        }

        private void cb_eocp_level_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_eocp_level.SelectedIndex == -1) return;
            if (cb_gocp_level.SelectedIndex == -1) return;

            W20.Value = cb_eocp_level.SelectedIndex << 4 | (int)W20.Value & 0x8F;
            W20.Value = cb_gocp_level.SelectedIndex << 0 | (int)W20.Value & 0xF0;
        }

        private void cb_socp_time_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_socp_time.SelectedIndex == -1) return;
            if (cb_socp_level.SelectedIndex == -1) return;

            W21.Value = cb_socp_time.SelectedIndex << 4 | (int)W21.Value & 0x0F;
            W21.Value = cb_socp_level.SelectedIndex << 0 | (int)W21.Value & 0xF0;
        }

        private void cb_sclk_psk_rst_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_sclk_psk_rst.SelectedIndex == -1) return;
            W22.Value = cb_sclk_psk_rst.SelectedIndex << 5 | (int)W22.Value & 0xDF;
        }

        private void cb_dummy_clk_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_dummy_clk.SelectedIndex == -1) return;
            if (cb_reverse.SelectedIndex == -1) return;
            if (cb_double.SelectedIndex == -1) return;

            W23.Value = cb_dummy_clk.SelectedIndex << 5 | (int)W23.Value & 0xDF;
            W23.Value = cb_reverse.SelectedIndex << 3 | (int)W23.Value & 0xF7;
            W23.Value = cb_double.SelectedIndex << 2 | (int)W23.Value & 0xFB;
        }

        private void cb_vcc2_dis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vcc2_dis.SelectedIndex == -1) return;
            if (cb_vgl2_dis.SelectedIndex == -1) return;
            if (cb_avdd_ext_drv.SelectedIndex == -1) return;
            if (cb_ext_int.SelectedIndex == -1) return;


            W25.Value = cb_vcc2_dis.SelectedIndex << 5 | (int)W25.Value & 0xDF;
            W25.Value = cb_vgl2_dis.SelectedIndex << 4 | (int)W25.Value & 0xEF;
            W25.Value = cb_avdd_ext_drv.SelectedIndex << 1 | (int)W25.Value & 0xF9;
            W25.Value = cb_ext_int.SelectedIndex << 0 | (int)W25.Value & 0xFE;
        }

        private void cb_vcc1_sync_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vcc1_sync.SelectedIndex == -1) return;
            if (cb_vcc2_sync.SelectedIndex == -1) return;
            if (cb_vcc2_en.SelectedIndex == -1) return;
            if (cb_fre_vcc1.SelectedIndex == -1) return;
            if (cb_ft_vcc2.SelectedIndex == -1) return;

            W27.Value = cb_vcc1_sync.SelectedIndex << 6 | (int)W27.Value & 0xBF;
            W27.Value = cb_vcc2_sync.SelectedIndex << 5 | (int)W27.Value & 0xDF;
            W27.Value = cb_vcc2_en.SelectedIndex << 4 | (int)W27.Value & 0xEF;
            W27.Value = cb_fre_vcc1.SelectedIndex << 3 | (int)W27.Value & 0xF7;
            W27.Value = cb_ft_vcc2.SelectedIndex << 2 | (int)W27.Value & 0xFB;
        }

        private void cb_vgh_sst_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vgh_sst.SelectedIndex == -1) return;
            if (cb_avdd_ss.SelectedIndex == -1) return;
            if (cb_fre_avdd.SelectedIndex == -1) return;
            if (cb_fre_havdd.SelectedIndex == -1) return;
            if (cb_fre_vgh.SelectedIndex == -1) return;
            if (cb_fre_vgl.SelectedIndex == -1) return;
            if (cb_pmic_en.SelectedIndex == -1) return;

            try
            {
                ComboBox[] cb_arr = new ComboBox[]
                {
                    cb_pmic_en, cb_fre_vgl, cb_fre_vgh, cb_fre_havdd, cb_fre_avdd, cb_avdd_ss, cb_vgh_sst
                };

                int data = 0x00;
                for (int i = 0; i < 8; i++) data |= cb_arr[i - 5].SelectedIndex << i;
                W2A.Value = data;
            }
            catch
            {

            }
        }

        private void cb_ocp_level_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_ocp_level.SelectedIndex == -1) return;
            if (cb_ocp_time.SelectedIndex == -1) return;

            W2B.Value = cb_ocp_level.SelectedIndex << 4 | (int)W2B.Value & 0x8F;
            W2B.Value = cb_ocp_time.SelectedIndex << 0 | (int)W2B.Value & 0xF8;
        }

        private void cb_avdd_protect_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cb_avdd_protect.SelectedIndex == -1) return;
            //if (cb_vcc1_protect.SelectedIndex == -1) return;
            //if (cb_havdd_protect.SelectedIndex == -1) return;
            //if (cb_vgh_protect.SelectedIndex == -1) return;
            //if (cb_vgh_protect.SelectedIndex == -1) return;
            //if (cb_vgl1_protect.SelectedIndex == -1) return;

            //try
            //{
            //    ComboBox[] cb_arr = new ComboBox[]
            //    {
            //        cb_vgl1_protect, cb_vgh_protect, cb_vgh_protect, cb_havdd_protect, cb_vcc1_protect, cb_avdd_protect
            //    };

            //    int data = 0x00;
            //    for (int i = 0; i < 6; i++) data |= cb_arr[i].SelectedIndex << i;
            //    W2C.Value = data | (int)W2C.Value & 0xC0;
            //}
            //catch
            //{

            //}
        }

        private void cb_ls7_protect_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cb_ls7_protect.SelectedIndex == -1) return;
            //if (cb_ls6_protect.SelectedIndex == -1) return;
            //if (cb_ls5_protect.SelectedIndex == -1) return;
            //if (cb_ls4_protect.SelectedIndex == -1) return;
            //if (cb_ls3_protect.SelectedIndex == -1) return;
            //if (cb_ls2_protect.SelectedIndex == -1) return;
            //if (cb_ls1_protect.SelectedIndex == -1) return;
            //if (cb_otp_protect.SelectedIndex == -1) return;
            //try
            //{
            //    ComboBox[] cb_arr = new ComboBox[]
            //    {
            //        cb_otp_protect, cb_ls1_protect, cb_ls2_protect, cb_ls3_protect, cb_ls4_protect, cb_ls5_protect, cb_ls6_protect, cb_ls7_protect
            //    };

            //    int data = 0x00;
            //    for (int i = 0; i < 8; i++) data |= cb_arr[i].SelectedIndex << i;
            //    W2D.Value = data;
            //}
            //catch
            //{

            //}


        }

        private void cb_ls_en_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_ls_en.SelectedIndex == -1) return;
            if (cb_hsr.SelectedIndex == -1) return;
            if (cb_clk_rising.SelectedIndex == -1) return;
            if (cb_clk_falling.SelectedIndex == -1) return;
            W40.Value = cb_ls_en.SelectedIndex << 7 | (int)W40.Value & 0x7F;
            W40.Value = cb_hsr.SelectedIndex << 4 | (int)W40.Value & 0x8F;
            W40.Value = cb_clk_rising.SelectedIndex << 2 | (int)W40.Value & 0xF3;
            W40.Value = cb_clk_falling.SelectedIndex << 0 | (int)W40.Value & 0xFC;
        }

        private void cb_stv1_dis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_stv1_dis.SelectedIndex == -1) return;
            if (cb_stv2_dis.SelectedIndex == -1) return;
            if (cb_stv3_dis.SelectedIndex == -1) return;
            if (cb_disch_dis.SelectedIndex == -1) return;

            W41.Value = cb_stv1_dis.SelectedIndex << 6 | cb_stv2_dis.SelectedIndex << 4 | cb_stv3_dis.SelectedIndex << 2 | cb_disch_dis.SelectedIndex;

        }

        private void cb_clk_dis_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_clk_dis.SelectedIndex == -1) return;
            if (cb_lc_dis.SelectedIndex == -1) return;
            if (cb_lc_init.SelectedIndex == -1) return;
            if (cb_auto_pulse.SelectedIndex == -1) return;
            W42.Value = cb_clk_dis.SelectedIndex << 6 | cb_lc_dis.SelectedIndex << 4 | cb_lc_init.SelectedIndex << 1 | cb_auto_pulse.SelectedIndex << 0;
        }

        private void cb_vcom_dly_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vcom_dly.SelectedIndex == -1) return;
            if (cb_xon_on_dly.SelectedIndex == -1) return;
            if (cb_xon_off_dly.SelectedIndex == -1) return;

            W43.Value = cb_vcom_dly.SelectedIndex << 5 | cb_xon_on_dly.SelectedIndex << 3 | cb_xon_off_dly.SelectedIndex << 1 | (int)W43.Value & 0x01;
        }

        private void cb_ilmt1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_ilmt1.SelectedIndex == -1) return;
            if (cb_ilmta.SelectedIndex == -1) return;
            if (cb_vin_uvlo.SelectedIndex == -1) return;
            if (cb_enE_Type.SelectedIndex == -1) return;
            W44.Value = cb_ilmt1.SelectedIndex << 7 | cb_ilmta.SelectedIndex << 5 | cb_vin_uvlo.SelectedIndex << 3 | cb_enE_Type.SelectedIndex << 0 | (int)W44.Value & 0x44;
        }

        private void cb_vgh_uvlo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_vgh_uvlo.SelectedIndex == -1) return;
            if (cb_stv_rest.SelectedIndex == -1) return;
            if (cb_ch_mode.SelectedIndex == -1) return;
            if (cb_power_off.SelectedIndex == -1) return;
            W45.Value = cb_vgh_uvlo.SelectedIndex << 7 | cb_stv_rest.SelectedIndex << 6 | cb_ch_mode.SelectedIndex << 3 | cb_power_off.SelectedIndex;
        }

        private int GetValue(int code, int strart_bit, int end_bit)
        {
            int res = 0x00;
            int[] bit_arr = new int[] { 0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80 };
            int mask_value = 0x00;
            for(int i = end_bit; i < strart_bit + 1; i++)
            {
                mask_value |= bit_arr[i];
            }
            res = code & mask_value;
            return res >> end_bit;
        }

        private void W00_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W00.Value;
            AVDDH.Value = code & 0x3F;
        }

        private void W01_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W01.Value;
            VCC1H.Value = code & 0x7F;
            cb_vcc1_ss.SelectedIndex = (code & 0x80) >> 7;
        }

        private void W02_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W02.Value;
            VCC2H.Value = code & 0x1F;
            cb_vcc2_ss.SelectedIndex = (code & 0x80) >> 7;
        }

        private void W03_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W03.Value;
            VGHLTH.Value = GetValue(code, 6, 0);
        }

        private void W04_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W04.Value;
            VGHHTH.Value = GetValue(code, 6, 0);
        }

        private void W05_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W05.Value;
            VGL1H.Value = GetValue(code, 6, 0);
        }

        private void W06_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W06.Value;
            VGL2LTH.Value = GetValue(code, 6, 0);
        }

        private void W07_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W07.Value;
            VGL2HTH.Value = GetValue(code, 6, 0);
        }

        private void W08_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W08.Value;
            cb_dly0.SelectedIndex = GetValue(code, 7, 6);
            cb_dly1.SelectedIndex = GetValue(code, 5, 4);
            cb_dly2.SelectedIndex = GetValue(code, 3, 2);
            cb_dly3.SelectedIndex = GetValue(code, 1, 0);
        }

        private void W09_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W09.Value;
            int MSB = 0x00;
            GLDOH.Value = GetValue(code, 7, 3);
            MSB = GetValue(code, 1, 0);
            HAVDDH.Value = MSB << 8 | (int)W0A.Value;
        }

        private void W0A_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W0A.Value;
            int reg09 = (int)W09.Value;
            int MSB = GetValue(reg09, 1, 0);
            HAVDDH.Value = MSB << 8 | code;
        }

        private void W0B_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W0B.Value;
            int LSB = (int)W0C.Value;
            int MSB = GetValue(code, 1, 0);
            GAM1H.Value = MSB << 8 | LSB;
        }

        private void W0D_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W0D.Value & 0x03;
            int LSB = (int)W0E.Value & 0xff;

            GAM2H.Value = MSB << 8 | LSB;
        }

        private void W0F_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W0F.Value & 0x03;
            int LSB = (int)W10.Value & 0xff;

            GAM3H.Value = MSB << 8 | LSB;
        }

        private void W11_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W11.Value & 0x03;
            int LSB = (int)W12.Value & 0xff;

            GAM4H.Value = MSB << 8 | LSB;
        }

        private void W13_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W13.Value & 0x03;
            int LSB = (int)W14.Value & 0xff;

            GAM5H.Value = MSB << 8 | LSB;
        }

        private void W15_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W15.Value & 0x03;
            int LSB = (int)W16.Value & 0xff;

            GAM6H.Value = MSB << 8 | LSB;
        }

        private void W17_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W17.Value & 0x03;
            int LSB = (int)W18.Value & 0xff;

            GAM7H.Value = MSB << 8 | LSB;
        }

        private void W19_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W19.Value & 0x03;
            int LSB = (int)W1A.Value & 0xff;

            GAM8H.Value = MSB << 8 | LSB;
        }

        private void W1B_ValueChanged(object sender, EventArgs e)
        {
            int MSB = (int)W1B.Value & 0x03;
            int LSB = (int)W1C.Value & 0xff;

            GAM9H.Value = MSB << 8 | LSB;
        }

        private void W1D_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W1D.Value;
            cb_vgh_tc_en.SelectedIndex = GetValue(code, 7, 7);
            cb_vcom_tc_en.SelectedIndex = GetValue(code, 6, 6);
            cb_tc_type.SelectedIndex = GetValue(code, 5, 5);
        }

        private void W1E_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W1E.Value;
            cb_vgh_tc_mode.SelectedIndex = GetValue(code, 7, 6);
            cb_vgx_prt_off.SelectedIndex = GetValue(code, 5, 5);
            cb_vcom_tc.SelectedIndex = GetValue(code, 3, 0);
        }

        private void W1F_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W1F.Value;
            cb_eocp_time.SelectedIndex = GetValue(code, 7, 4);
            cb_gocp_time.SelectedIndex = GetValue(code, 4, 0);
        }

        private void W20_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W20.Value;
            cb_eocp_level.SelectedIndex = GetValue(code, 6, 4);
            cb_gocp_level.SelectedIndex = GetValue(code, 3, 0);
        }

        private void W21_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W21.Value;
            cb_socp_time.SelectedIndex = GetValue(code, 7, 4);
            cb_socp_level.SelectedIndex = GetValue(code, 3, 0);
        }

        private void W22_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W22.Value;
            cb_sclk_psk_rst.SelectedIndex = GetValue(code, 5, 5);
        }

        private void W23_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W23.Value;
            cb_dummy_clk.SelectedIndex = GetValue(code, 5, 5);
            cb_reverse.SelectedIndex = GetValue(code, 3, 3);
            cb_double.SelectedIndex = GetValue(code, 2, 2);
        }

        private void W24_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W24.Value;
            cb_avdd_dis.SelectedIndex = GetValue(code, 7, 7);
            cb_havdd_dis.SelectedIndex = GetValue(code, 6, 6);
            cb_vgh_dis.SelectedIndex = GetValue(code, 5, 5);
            cb_vgl1_dis.SelectedIndex = GetValue(code, 4, 4);
            cb_vcc_dis.SelectedIndex = GetValue(code, 3, 3);
            cb_vcom3_dis.SelectedIndex = GetValue(code, 2, 2);
            cb_vcom2_dis.SelectedIndex = GetValue(code, 1, 1);
            cb_vcom1_dis.SelectedIndex = GetValue(code, 0, 0);
        }

        private void W25_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W25.Value;
            cb_vcc2_dis.SelectedIndex = GetValue(code, 5, 5);
            cb_vgl2_dis.SelectedIndex = GetValue(code, 4, 4);
            cb_avdd_ext_drv.SelectedIndex = GetValue(code, 2, 1);
            cb_ext_int.SelectedIndex = GetValue(code, 0, 0);
        }

        private void W26_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W26.Value;
            cb_vcom1_en.SelectedIndex = GetValue(code, 7, 7);
            cb_vcom2_en.SelectedIndex = GetValue(code, 6, 6);
            cb_vcom3_en.SelectedIndex = GetValue(code, 5, 5);
        }

        private void W27_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W27.Value;
            cb_vcc1_sync.SelectedIndex = GetValue(code, 6, 6);
            cb_vcc2_sync.SelectedIndex = GetValue(code, 5, 5);
            cb_vcc2_en.SelectedIndex = GetValue(code, 4, 4);
            cb_fre_vcc1.SelectedIndex = GetValue(code, 3, 3);
            cb_ft_vcc2.SelectedIndex = GetValue(code, 2, 2);
        }

        private void W29_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W29.Value;
            cb_protection.SelectedIndex = GetValue(code, 7, 7);
            cb_vcc_en.SelectedIndex = GetValue(code, 6, 6);
            cb_vgl1_en.SelectedIndex = GetValue(code, 5, 5);
            cb_vgl2_en.SelectedIndex = GetValue(code, 4, 4);
            cb_avdd_en.SelectedIndex = GetValue(code, 3, 3);
            cb_vgh_en.SelectedIndex = GetValue(code, 2, 2);
            cb_havdd_en.SelectedIndex = GetValue(code, 1, 1);
            cb_gam_en.SelectedIndex = GetValue(code, 0, 0);
        }

        private void W2A_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W2A.Value;
            cb_vgh_sst.SelectedIndex = GetValue(code, 7, 7);
            cb_avdd_ss.SelectedIndex = GetValue(code, 6, 5);
            cb_fre_avdd.SelectedIndex = GetValue(code, 4, 4);
            cb_fre_havdd.SelectedIndex = GetValue(code, 3, 3);
            cb_fre_vgh.SelectedIndex = GetValue(code, 2, 2);
            cb_fre_vgl.SelectedIndex = GetValue(code, 1, 1);
            cb_pmic_en.SelectedIndex = GetValue(code, 0, 0);
        }

        private void W2B_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W2B.Value;
            cb_ocp_level.SelectedIndex = GetValue(code, 7, 5);
            cb_ocp_time.SelectedIndex = GetValue(code, 2, 0);
        }

        private void W2C_ValueChanged(object sender, EventArgs e)
        {

        }

        private void W2D_ValueChanged(object sender, EventArgs e)
        {

        }

        private void W40_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W40.Value;
            cb_ls_en.SelectedIndex = GetValue(code, 7, 7);
            cb_hsr.SelectedIndex = GetValue(code, 6, 4);
            cb_clk_rising.SelectedIndex = GetValue(code, 3, 2);
            cb_clk_falling.SelectedIndex = GetValue(code, 1, 0);
        }

        private void W41_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W41.Value;
            cb_stv1_dis.SelectedIndex = GetValue(code, 7, 6);
            cb_stv2_dis.SelectedIndex = GetValue(code, 5, 4);
            cb_stv3_dis.SelectedIndex = GetValue(code, 3, 2);
            cb_disch_dis.SelectedIndex = GetValue(code, 1, 0);
        }

        private void W42_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W42.Value;
            cb_clk_dis.SelectedIndex = GetValue(code, 7, 6);
            cb_lc_dis.SelectedIndex = GetValue(code, 5, 4);
            cb_lc_init.SelectedIndex = GetValue(code, 3, 1);
            cb_auto_pulse.SelectedIndex = GetValue(code, 0, 0);
        }

        private void W43_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W43.Value;
            cb_vcom_dly.SelectedIndex = GetValue(code, 7, 5);
            cb_xon_on_dly.SelectedIndex = GetValue(code, 4, 3);
            cb_xon_off_dly.SelectedIndex = GetValue(code, 2, 1);

        }

        private void W44_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W44.Value;
            cb_ilmt1.SelectedIndex = GetValue(code, 7, 7);
            cb_ilmta.SelectedIndex = GetValue(code, 5, 5);
            cb_vin_uvlo.SelectedIndex = GetValue(code, 4, 3);
            cb_enE_Type.SelectedIndex = GetValue(code, 1, 0);
        }

        private void W45_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W45.Value;
            cb_vgh_uvlo.SelectedIndex = GetValue(code, 7, 7);
            cb_stv_rest.SelectedIndex = GetValue(code, 6, 6);
            cb_ch_mode.SelectedIndex = GetValue(code, 5, 3);
            cb_power_off.SelectedIndex = GetValue(code, 2, 0);
        }

        private void W2F_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W30.Value;
            int MSB = GetValue(code, 7, 6);
            GAM10H.Value = MSB << 8 | (int)W2F.Value;
        }

        private void W30_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W30.Value;
            int GAM10_MSB = GetValue(code, 7, 6);
            int GAM11_MSB = GetValue(code, 1, 0);

            GAM10H.Value = GAM10_MSB << 8 | (int)W2F.Value;
            GAM11H.Value = GAM11_MSB << 8 | (int)W31.Value;

        }

        private void W32_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W32.Value;
            int MSB = GetValue(code, 1, 0);
            GAM12H.Value = MSB << 8 | (int)W33.Value;
        }

        private void W34_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W34.Value;
            int MSB = GetValue(code, 1, 0);
            GAM13H.Value = MSB << 8 | (int)W35.Value;
        }

        private void W36_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W36.Value;
            int MSB = GetValue(code, 1, 0);
            GAM14H.Value = MSB << 8 | (int)W37.Value;
        }

        private void W38_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W38.Value;
            int MSB = GetValue(code, 1, 0);
            VCOM1H.Value = MSB << 8 | (int)W39.Value;
        }

        private void W3A_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W3A.Value;
            int MSB = GetValue(code, 1, 0);
            VCOM2H.Value = MSB << 8 | (int)W3B.Value;
        }

        private void W3C_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W3C.Value;
            int MSB = GetValue(code, 1, 0);
            VCOM3H.Value = MSB << 8 | (int)W3D.Value;
        }

        private void VGL2LTSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGL2LTH.Value = VGL2LTSL.Value;
        }

        private void R2D_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W2D.Value;
            cb_ls7_protect.SelectedIndex = GetValue(code, 7, 7);
            cb_ls6_protect.SelectedIndex = GetValue(code, 6, 6);
            cb_ls5_protect.SelectedIndex = GetValue(code, 5, 5);
            cb_otp_protect.SelectedIndex = GetValue(code, 4, 4);

            cb_ls4_protect.SelectedIndex = GetValue(code, 3, 3);
            cb_ls3_protect.SelectedIndex = GetValue(code, 2, 2);
            cb_ls2_protect.SelectedIndex = GetValue(code, 1, 1);
            cb_ls1_protect.SelectedIndex = GetValue(code, 0, 0);
        }

        private void R2C_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)W2C.Value;
            cb_avdd_protect.SelectedIndex = GetValue(code, 5, 5);
            cb_vcc1_protect.SelectedIndex = GetValue(code, 4, 4);
            cb_havdd_protect.SelectedIndex = GetValue(code, 3, 3);
            cb_vgh_protect.SelectedIndex = GetValue(code, 2, 2);
            cb_vgl2_protect.SelectedIndex = GetValue(code, 1, 1);
            cb_vgl1_protect.SelectedIndex = GetValue(code, 0, 0);
        }

        private void bt_ReadtoWrite_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < WriteTable.Length; i++)
            {
                WriteTable[i].Value = ReadTable[i].Value;
            }
        }

        private void bt_write_all_Click(object sender, EventArgs e)
        {
            byte slave = (byte)((byte)nuSlave.Value >> 1);
            byte[] WriteBuffer = new byte[WriteTable.Length];

            for(int i = 0; i < WriteTable.Length; i++)
            {
                WriteBuffer[i] = (byte)WriteTable[i].Value;
            }
            RTDev.I2C_Write(slave, 0x00, WriteBuffer);
        }

        private void bt_read_all_Click(object sender, EventArgs e)
        {
            byte slave = (byte)((byte)nuSlave.Value >> 1);
            byte[] ReadBuffer = new byte[ReadTable.Length];

            RTDev.I2C_Write(slave, 0xff, new byte[] { 0x00 });
            RTDev.I2C_Read(slave, 0x00, ref ReadBuffer);


            for(int i = 0; i < ReadTable.Length; i++)
            {
                ReadTable[i].Value = ReadBuffer[i];
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            byte slave = (byte)((byte)nuSlave.Value >> 1);
            byte[] WriteBuffer = new byte[WriteTable.Length];

            for (int i = 0; i < WriteTable.Length; i++)
            {
                WriteBuffer[i] = (byte)WriteTable[i].Value;
            }
            RTDev.I2C_Write(slave, 0x00, WriteBuffer);
            RTDev.I2C_Write(slave, 0xff, new byte[] { 0x80 });
        }

        private void button1_Click(object sender, EventArgs e)
        {
            byte slave = (byte)((byte)nuSlave.Value >> 1);
            byte[] ReadBuffer = new byte[ReadTable.Length];

            RTDev.I2C_Write(slave, 0xff, new byte[] { 0x01 });
            RTDev.I2C_Read(slave, 0x00, ref ReadBuffer);


            for (int i = 0; i < ReadTable.Length; i++)
            {
                ReadTable[i].Value = ReadBuffer[i];
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            byte slave = (byte)((byte)nuSlave.Value >> 1);
            RTDev.I2C_Write(slave, 0xff, new byte[] { 0x40 });
        }

        private void saveBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Bin File|*.bin";
            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = saveDlg.FileName;
                List<byte> bin_buf = new List<byte>();
                BinaryWriter bw = new BinaryWriter(new FileStream(file_name, FileMode.Create));

                for (int i = 0; i < 0x100; i++)
                {
                    if (i < WriteTable.Length)
                    {
                        bin_buf.Add(Convert.ToByte(WriteTable[i].Value));
                    }
                    else
                    {
                        bin_buf.Add(0);
                    }

                }
                bw.Write(bin_buf.ToArray());
                bw.Close();
            }
        }

        private void openBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Bin File|*.bin";
            if (openDlg.ShowDialog() == DialogResult.OK)
            {
                byte[] ReadBuf = new byte[255];
                string file_name = openDlg.FileName;
                BinaryReader br = new BinaryReader(new FileStream(file_name, FileMode.Open));

                br.Read(ReadBuf, 0, 0xff);

                for (int i = 0; i < 0x100; i++)
                {
                    if (i < WriteTable.Length)
                    {
                        WriteTable[i].Value = ReadBuf[i];
                    }
                }
                br.Close();
            }
        }

        private void linkRTBridgeBoardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(!RTDev.BoadInit())
            {
                MessageBox.Show("Linking bridge board fail !!!", this.Text);
            }
        }

        private void cb_vgh_tc_mode_MouseEnter(object sender, EventArgs e)
        {
            if(cb_tc_type.SelectedIndex == 0)
            {
                cb_vgh_tc_mode.Width = 300;
                groupBox31.Width = 350;
            }
            else
            {
                cb_vgh_tc_mode.Width = 450;
                groupBox31.Width = 550;
            }


        }

        private void cb_vgh_tc_mode_MouseLeave(object sender, EventArgs e)
        {
            cb_vgh_tc_mode.Width = 121;
            groupBox31.Width = 165;
        }
    }
}
