using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.IO;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using System.Runtime.InteropServices;

namespace CS601C
{
    public partial class Form1 : Form
    {
        string win_name = "CS601C_v1.0.0";

        public static NumericUpDown[] WriteTable;
        public static NumericUpDown[] ReadTable;

        public static NumericUpDown[] WriteTable2;
        public static NumericUpDown[] ReadTable2;

        public static NumericUpDown[] WriteTMTable;
        public static NumericUpDown[] ReadTMTable;

        public static NumericUpDown[] WriteTMTable2;
        public static NumericUpDown[] ReadTMTable2;

        public static TextBox[] StatusReg1;
        public static TextBox[] StatusReg2;
        public static TextBox[] StatusReg3;
        public static TextBox[] StatusReg4;

        private int bit0 = 0x01;
        private int bit1 = 0x02;
        private int bit2 = 0x04;
        private int bit3 = 0x08;
        private int bit4 = 0x10;
        private int bit5 = 0x20;
        private int bit6 = 0x40;
        private int bit7 = 0x80;
        private int[] bit_table;

        RTBBControl RTDev = new RTBBControl();

        Thread thread;


        public Form1()
        {
            InitializeComponent();
            RTDev.BoadInit();
            List<byte> list = new List<byte>();
            if(list != null)
            {
                if (list.Count > 0)
                    nuSlave.Value = list[0] << 1;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = win_name;
            WriteTable = new NumericUpDown[] {
                W00, W01, W02, W03, W04, W05, W06, W07, W08, W09, W0A, W0B, W0C, W0D, W0E, W0F,
                W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W1A, W1B, W1C, W1D, W1E, W1F,
                W20, W21, W22, W23, W24, W25, W26, W27, W28, W29, W2A, W2B, W2C, W2D, W2E, W2F,
            };

            ReadTable = new NumericUpDown[]
            {
                R00, R01, R02, R03, R04, R05, R06, R07, R08, R09, R0A, R0B, R0C, R0D, R0E, R0F,
                R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R1A, R1B, R1C, R1D, R1E, R1F,
                R20, R21, R22, R23, R24, R25, R26, R27, R28, R29, R2A, R2B, R2C, R2D, R2E, R2F,
            };

            WriteTable2 = new NumericUpDown[]
            {
                W30, W31, W32, W33, W34, W35, W36, W37, W38, W39, W3A, W3B, W3C, W3D, W3E, W3F,
                W40, W41, W42, W43, W44, W45, W46, W47, W48, W49, W4A, W4B, W4C, W4D, W4E, W4F,
                W50, W51, W52, W53, W54, W55, W56, W57, W58, W59, W5A, W5B, W5C, W5D, W5E, W5F,
            };

            ReadTable2 = new NumericUpDown[]
            {
                R30, R31, R32, R33, R34, R35, R36, R37, R38, R39, R3A, R3B, R3C, R3D, R3E, R3F,
                R40, R41, R42, R43, R44, R45, R46, R47, R48, R49, R4A, R4B, R4C, R4D, R4E, R4F,
                R50, R51, R52, R53, R54, R55, R56, R57, R58, R59, R5A, R5B, R5C, R5D, R5E, R5F,
            };

            WriteTMTable = new NumericUpDown[]
            {
                W70, W71, W72, W73, W74, W75, W76, W77, W78, W79, W7A, W7B, W7C, W7D, W7E, W7F,
                W80, W81, W82, W83, W84, W85, W86, W87, W88, W89, W8A, W8B, W8C, W8D, W8E, W8F,
                W90, W91, W92, W93, W94, W95, W96, W97, W98, W99, W9A, W9B, W9C, W9D, W9E, W9F,
                WA0, WA1, WA2, WA3, WA4, WA5, WA6, WA7, WA8, WA9, WAA, WAB, WAC, WAD, WAE, WAF,
                WB0, WB1, WB2, WB3, WB4, WB5, WB6, WB7, WB8, WB9, WBA, WBB, WBC, WBD, WBE, WBF,
                
            };

            ReadTMTable = new NumericUpDown[]
            {
                R70, R71, R72, R73, R74, R75, R76, R77, R78, R79, R7A, R7B, R7C, R7D, R7E, R7F,
                R80, R81, R82, R83, R84, R85, R86, R87, R88, R89, R8A, R8B, R8C, R8D, R8E, R8F,
                R90, R91, R92, R93, R94, R95, R96, R97, R98, R99, R9A, R9B, R9C, R9D, R9E, R9F,
                RA0, RA1, RA2, RA3, RA4, RA5, RA6, RA7, RA8, RA9, RAA, RAB, RAC, RAD, RAE, RAF,
                RB0, RB1, RB2, RB3, RB4, RB5, RB6, RB7, RB8, RB9, RBA, RBB, RBC, RBD, RBE, RBF,
                
            };

            WriteTMTable2 = new NumericUpDown[]
            {
                WE0, WE1, WE2, WE3, WE4, WE5, WE6, WE7, WE8, WE9, WEA, WEB, WEC, WED, WEE, WEF,
            };

            ReadTMTable2 = new NumericUpDown[]
            {
                RE0, RE1, RE2, RE3, RE4, RE5, RE6, RE7, RE8, RE9, REA, REB, REC, RED, REE, REF,
            };

            StatusReg1 = new TextBox[]
            {
                tbS1_7, tbS1_6, tbS1_5, tbS1_4, tbS1_3, tbS1_2, tbS1_1, tbS1_0
            };

            StatusReg2 = new TextBox[]
            {
                tbS2_7, tbS2_6, tbS2_5, tbS2_4, tbS2_3, tbS2_2, tbS2_1, tbS2_0
            };

            StatusReg3 = new TextBox[]
            {
                tbS3_7, tbS3_6, tbS3_5, tbS3_4, tbS3_3, tbS3_2, tbS3_1, tbS3_0
            };

            StatusReg4 = new TextBox[]
            {
                tbS4_7, tbS4_6, tbS4_5, tbS4_4, tbS4_3, tbS4_2, tbS4_1, tbS4_0
            };


            bit_table = new int[]
            {
                bit0, bit1, bit2, bit3, bit4, bit5, bit6, bit7
            };

            for (int i = 0; i < WriteTable.Length; i++)
            {
                WriteTable[i].Value = 0x01;
                WriteTable[i].Value = 0x00;

                WriteTable2[i].Value = 0x01;
                WriteTable2[i].Value = 0x00;
            }

            // GUI initial
            CBSEL_5V12V.SelectedIndex = 1;
            CBSEL_5V12V2.SelectedIndex = 1;

            comboBox1.SelectedIndex = 0;
            AVDDOCH.Value = 1;
            AVDDOCH.Value = 0;
        }

        private void CalculateGAM_VCOM()
        {
            double res = (double)OPLDOV.Value / 1024;
            double GAM1 = (double)GAM1H.Value * res;
            double GAM2 = (double)GAM2H.Value * res;
            double GAM3 = (double)GAM3H.Value * res;
            double GAM4 = (double)GAM4H.Value * res;
            double GAM5 = (double)GAM5H.Value * res;
            double GAM6 = (double)GAM6H.Value * res;
            double GAM7 = (double)GAM7H.Value * res;
            double GAM8 = (double)GAM8H.Value * res;
            double GAM9 = (double)GAM9H.Value * res;
            double GAM10 = (double)GAM10H.Value * res;
            double GAM11 = (double)GAM11H.Value * res;
            double GAM12 = (double)GAM12H.Value * res;
            double GAM13 = (double)GAM13H.Value * res;
            double GAM14 = (double)GAM14H.Value * res;
            int vcom1_code = 0;
            int vcom2_code = 0;
            if (VCOM1H.Value <= 0x186)
            {
                vcom1_code = (int)VCOM1H.Value + 250;
            }

            if(VCOM2H.Value <= 0x186)
            {
                vcom2_code = (int)VCOM2H.Value + 250;
            }
            

            double VCOM1 = vcom1_code * res;
            double VCOM2 = vcom2_code * res;

            GAM1V.Value = (decimal)GAM1;
            GAM2V.Value = (decimal)GAM2;
            GAM3V.Value = (decimal)GAM3;
            GAM4V.Value = (decimal)GAM4;
            GAM5V.Value = (decimal)GAM5;
            GAM6V.Value = (decimal)GAM6;
            GAM7V.Value = (decimal)GAM7;
            GAM8V.Value = (decimal)GAM8;
            GAM9V.Value = (decimal)GAM9;
            GAM10V.Value = (decimal)GAM10;
            GAM11V.Value = (decimal)GAM11;
            GAM12V.Value = (decimal)GAM12;
            GAM13V.Value = (decimal)GAM13;
            GAM14V.Value = (decimal)GAM14;
            VCOM1V.Value = (decimal)VCOM1;
            VCOM2V.Value = (decimal)VCOM2;

            res = (double)OPLDO2V.Value / 1024;
            GAM1 = (double)GAM1_2H.Value * res;
            GAM2 = (double)GAM2_2H.Value * res;
            GAM3 = (double)GAM3_2H.Value * res;
            GAM4 = (double)GAM4_2H.Value * res;
            GAM5 = (double)GAM5_2H.Value * res;
            GAM6 = (double)GAM6_2H.Value * res;
            GAM7 = (double)GAM7_2H.Value * res;
            GAM8 = (double)GAM8_2H.Value * res;
            GAM9 = (double)GAM9_2H.Value * res;
            GAM10 = (double)GAM10_2H.Value * res;
            GAM11 = (double)GAM11_2H.Value * res;
            GAM12 = (double)GAM12_2H.Value * res;
            GAM13 = (double)GAM13_2H.Value * res;
            GAM14 = (double)GAM14_2H.Value * res;

            if (VCOM1_2H.Value <= 0x186)
            {
                vcom1_code = (int)VCOM1_2H.Value + 250;
            }

            if (VCOM2_2H.Value <= 0x186)
            {
                vcom2_code = (int)VCOM2_2H.Value + 250;
            }

            VCOM1 = vcom1_code * res;
            VCOM2 = vcom2_code * res;

            GAM1_2V.Value = (decimal)GAM1;
            GAM2_2V.Value = (decimal)GAM2;
            GAM3_2V.Value = (decimal)GAM3;
            GAM4_2V.Value = (decimal)GAM4;
            GAM5_2V.Value = (decimal)GAM5;
            GAM6_2V.Value = (decimal)GAM6;
            GAM7_2V.Value = (decimal)GAM7;
            GAM8_2V.Value = (decimal)GAM8;
            GAM9_2V.Value = (decimal)GAM9;
            GAM10_2V.Value = (decimal)GAM10;
            GAM11_2V.Value = (decimal)GAM11;
            GAM12_2V.Value = (decimal)GAM12;
            GAM13_2V.Value = (decimal)GAM13;
            GAM14_2V.Value = (decimal)GAM14;
            VCOM1_2V.Value = (decimal)VCOM1;
            VCOM2_2V.Value = (decimal)VCOM2;
        }


        private void CKVDD_CheckedChanged(object sender, EventArgs e)
        {
            if(CBSEL_5V12V.SelectedIndex != -1)
            {
                int data = 0x00;
                CheckBox[] _00h_table = new CheckBox[]
                {
                CKVDD, CKVGL, CKVSS, CKControl, CKAVDD, CKVGH
                };

                for (int i = 0; i < _00h_table.Length; i++)
                {
                    data |= (int)((_00h_table[i].Checked ? 0x01 << i : 0x00));
                }

                W00.Value = data | CBSEL_5V12V.SelectedIndex << 7 | ((int)W00.Value & 0x40);
            }

        }

        private void CBSEL_5V12V_SelectedIndexChanged(object sender, EventArgs e)
        {
            int data = 0x00;
            CheckBox[] _00h_table = new CheckBox[]
            {
                CKVDD, CKVGL, CKVSS, CKControl, CKAVDD, CKVGH
            };

            for (int i = 0; i < _00h_table.Length; i++)
            {
                data |= (byte)((_00h_table[i].Checked ? 0x01 << i : 0x00));
            }

            W00.Value =  data | CBSEL_5V12V.SelectedIndex << 7 | ((int)W00.Value & 0x40);
        }

        private void W00_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W00.Value;
            CheckBox[] _00h_table = new CheckBox[]
            {
                CKVDD, CKVGL, CKVSS, CKControl, CKAVDD, CKVGH
            };

            for (int i = 0; i < _00h_table.Length; i++)
            {
                if ((data & (0x01 << i)) == bit_table[i])
                {
                    _00h_table[i].Checked = true;
                }
                else
                {
                    _00h_table[i].Checked = false;
                }
            }

            CBSEL_5V12V.SelectedIndex = (data & 0x80) >> 7;
        }

        private void CBDLY0_SelectedIndexChanged(object sender, EventArgs e)
        {
            byte data = 0;
            ComboBox[] _01h_table = new ComboBox[]
            {
                CBDLY0, CBDLY1, CBDLY2, CBAVDDSS
            };

            data |= (byte)(CBDLY0.SelectedIndex << 0);       // 1:0
            data |= (byte)(CBDLY1.SelectedIndex << 2);       // 3:2
            data |= (byte)(CBDLY2.SelectedIndex << 4);       // 5:4
            data |= (byte)(CBAVDDSS.SelectedIndex << 6);     // 6
            W01.Value = data | ((byte)W01.Value & 0x80);
        }

        private void W01_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W01.Value;

            CBDLY0.SelectedIndex = (data & 0x03);
            CBDLY1.SelectedIndex = (data & 0x0C) >> 2;
            CBDLY2.SelectedIndex = (data & 0x30) >> 4;
            CBAVDDSS.SelectedIndex = (data & 0x40) >> 6;
        }

        private void CBSWFreq_SelectedIndexChanged(object sender, EventArgs e)
        {
            W02.Value = (byte)(CBLSDis.SelectedIndex << 7) | (byte)(CBSWFreq.SelectedIndex << 6) | (byte)AVDDH.Value;
        }

        private void AVDDH_ValueChanged(object sender, EventArgs e)
        {
            byte code = (byte)AVDDH.Value; 
            if(code > 0)
            {
                double volt = (double)(((AVDDH.Value - 1) + 130) / 10);
                if (volt <= 19.2)
                    AVDDV.Value = (decimal)volt;
            }
            AVDDSL.Value = (int)AVDDH.Value;
            int other = ((int)W02.Value & 0xC0);
            W02.Value = ((int)AVDDH.Value | other);
        }

        private void AVDDSL_Scroll_1(object sender, ScrollEventArgs e)
        {
            AVDDH.Value = AVDDSL.Value;
        }

        private void W02_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W02.Value;
            AVDDH.Value = data & 0x3f;
            CBSWFreq.SelectedIndex = (data & 0x40) >> 6;
            CBLSDis.SelectedIndex = (data & 0x80) >> 7;

        }

        private void CBAVDDOCP_SelectedIndexChanged(object sender, EventArgs e)
        {
            W03.Value = (byte)(CBAVDDOCP.SelectedIndex << 5) | (int)VDDH.Value |(int)W03.Value & 0xC0;
        }

        private void VDDH_ValueChanged(object sender, EventArgs e)
        {
            double volt = ((((double)VDDH.Value * 5) / 100) + 2.2);
            VDDSL.Value = (int)VDDH.Value;
            if (volt <= 3.7)
                VDDV.Value = (decimal)volt;

            W03.Value = (int)VDDH.Value | ((int)W03.Value & 0xE0);
        }

        private void VDDSL_Scroll(object sender, ScrollEventArgs e)
        {
            VDDH.Value = VDDSL.Value;
        }

        private void W03_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W03.Value;
            VDDH.Value = (int)data & 0x1F;
            CBAVDDOCP.SelectedIndex = ((int)data & 0x20) >> 5;
        }

        private void VSSH_ValueChanged(object sender, EventArgs e)
        {
            double volt = (((double)VSSH.Value * 5) + 45) / (-10);
            if (volt >= -16)
                VSSV.Value = (decimal)volt;

            VSSSL.Value = (int)VSSH.Value;
            W04.Value = (int)VSSH.Value | ((int)W04.Value & 0xE0);
        }

        private void VSSSL_Scroll(object sender, ScrollEventArgs e)
        {
            VSSH.Value = VSSSL.Value;
        }

        private void W04_ValueChanged(object sender, EventArgs e)
        {
            VSSH.Value = (int)W04.Value & 0x1F;
        }

        private void VGLH_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = ((VGLH.Value * 5) + 45) / -10;
            VGLV.Value = volt;
            VGLSL.Value = (int)VGLH.Value;
            W06.Value = (int)VGLH.Value | ((int)W06.Value & 0xE0);
        }

        private void W06_ValueChanged(object sender, EventArgs e)
        {
            VGLH.Value = ((int)W06.Value & 0x1F);
        }

        private void VGHH_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = VGHH.Value * 1 + 20;
            if (volt <= 40)
                VGHV.Value = volt;
            VGHSL.Value = (int)VGHH.Value;
            W08.Value = (int)VGHH.Value | ((int)W08.Value & 0xE0);
        }

        private void VGHSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGHH.Value = VGHSL.Value;
        }

        private void W08_ValueChanged(object sender, EventArgs e)
        {
            VGHH.Value = ((int)W08.Value & 0x1F);
        }

        private void OPLDOH_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = (OPLDOH.Value * 2 + 130) / 10;
            if (volt <= 18)
                OPLDOV.Value = volt;
            OPLDOSL.Value = (int)OPLDOH.Value;
            W09.Value = (int)OPLDOH.Value | ((int)W09.Value & 0xE0);
        }

        private void OPLDOSL_Scroll(object sender, ScrollEventArgs e)
        {
            OPLDOH.Value = OPLDOSL.Value;
        }

        private void W09_ValueChanged(object sender, EventArgs e)
        {
            OPLDOH.Value = (int)W09.Value & 0x1F;
        }

        private void AVDDOCH_ValueChanged(object sender, EventArgs e)
        {
            decimal amp = ((AVDDOCH.Value * 5) + 20) / 10;
            AVDDOCA.Value = amp;
            AVDDOCSL.Value = (int)AVDDOCH.Value;
            W0A.Value = (int)AVDDOCH.Value << 6 | ((int)W0A.Value & 0x3F);
        }

        private void AVDDOCSL_Scroll(object sender, ScrollEventArgs e)
        {
            AVDDOCH.Value = AVDDOCSL.Value;
        }

        private void W0A_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W0A.Value;
            AVDDOCH.Value = ((int)data & 0xC0) >> 6;
            CBDRVN.SelectedIndex = ((int)data & 0x38) >> 3;
            CBDRVP.SelectedIndex = ((int)data & 0x7);
        }


        private void CBPhase_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBLine.SelectedIndex != -1 && CBPhase.SelectedIndex != -1)
            {
                W0B.Value = CBLS_Dis.SelectedIndex << 4 | ((int)W0B.Value & 0x0F);
                W0B.Value = CBPhase.SelectedIndex << 2 | ((int)W0B.Value & 0xF3);
                W0B.Value = CBLine.SelectedIndex | ((int)W0B.Value & 0xFC);
            }
        }

        private void W0B_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W0B.Value;
            CBLS_Dis.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBPhase.SelectedIndex = ((int)data & 0xC) >> 2;
            CBLine.SelectedIndex = ((int)data & 0x3);
        }

        private void CBOCPLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBOCPLevel.SelectedIndex != -1 &&
                CBOCPDly.SelectedIndex != -1 &&
                CBLSOCPEnable.SelectedIndex != -1)
            {
                W0C.Value = CBOCPLevel.SelectedIndex << 4 | ((int)W0C.Value & 0x0F);
                W0C.Value = CBOCPDly.SelectedIndex << 1 | ((int)W0C.Value & 0xF1);
                W0C.Value = CBLSOCPEnable.SelectedIndex | ((int)W0C.Value & 0xFE);
            }
        }

        private void W0C_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W0C.Value;
            CBOCPLevel.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBOCPDly.SelectedIndex = ((int)data & 0x0E) >> 1;
            CBLSOCPEnable.SelectedIndex = ((int)data & 0x01);
        }

        private void CBDelayFrame_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBDelayFrame.SelectedIndex != -1 &&
                CBClockCycle.SelectedIndex != -1)
            {
                W0D.Value = CBDelayFrame.SelectedIndex << 4 | ((int)W0D.Value & 0x0F);
                W0D.Value = CBClockCycle.SelectedIndex | ((int)W0D.Value & 0xF0);
            }
        }

        private void W0D_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W0D.Value;
            CBDelayFrame.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBClockCycle.SelectedIndex = ((int)data & 0x0F);
        }

        private void GAM1H_ValueChanged(object sender, EventArgs e)
        {
            GAM1SL.Value = (int)GAM1H.Value;
            W0E.Value = ((int)GAM1H.Value & 0x300) >> 8 | ((int)W0E.Value & 0xFC);
            W0F.Value = (int)GAM1H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM1SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM1H.Value = GAM1SL.Value;
        }

        private void GAM2H_ValueChanged(object sender, EventArgs e)
        {
            GAM2SL.Value = (int)GAM2H.Value;
            W10.Value = ((int)GAM2H.Value & 0x300) >> 8 | ((int)W10.Value & 0xFC);
            W11.Value = ((int)GAM2H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM2H.Value = GAM2SL.Value;
        }

        private void GAM3H_ValueChanged(object sender, EventArgs e)
        {
            GAM3SL.Value = (int)GAM3H.Value;
            W12.Value = ((int)GAM3H.Value & 0x300) >> 8 | ((int)W12.Value & 0xFC);
            W13.Value = ((int)GAM3H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM4H_ValueChanged(object sender, EventArgs e)
        {
            GAM4SL.Value = (int)GAM4H.Value;
            W14.Value = ((int)GAM4H.Value & 0x300) >> 8 | ((int)W14.Value & 0xFC); ;
            W15.Value = ((int)GAM4H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM5H_ValueChanged(object sender, EventArgs e)
        {
            GAM5SL.Value = ((int)GAM5H.Value);
            W16.Value = ((int)GAM5H.Value & 0x300) >> 8 | ((int)W16.Value & 0xFC); ;
            W17.Value = ((int)GAM5H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM6H_ValueChanged(object sender, EventArgs e)
        {
            GAM6SL.Value = ((int)GAM6H.Value);
            W18.Value = ((int)GAM6H.Value & 0x300) >> 8 | ((int)W18.Value & 0xFC); ;
            W19.Value = ((int)GAM6H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM7H_ValueChanged(object sender, EventArgs e)
        {
            GAM7SL.Value = ((int)GAM7H.Value);
            W1A.Value = ((int)GAM7H.Value & 0x300) >> 8 | ((int)W1A.Value & 0xFC); ;
            W1B.Value = ((int)GAM7H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM8H_ValueChanged(object sender, EventArgs e)
        {
            GAM8SL.Value = ((int)GAM8H.Value);
            W1C.Value = ((int)GAM8H.Value & 0x300) >> 8 | ((int)W1C.Value & 0xFC); ;
            W1D.Value = ((int)GAM8H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM9H_ValueChanged(object sender, EventArgs e)
        {
            GAM9SL.Value = ((int)GAM9H.Value);
            W1E.Value = ((int)GAM9H.Value & 0x300) >> 8 | ((int)W1E.Value & 0xFC); ;
            W1F.Value = ((int)GAM9H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM10H_ValueChanged(object sender, EventArgs e)
        {
            GAM10SL.Value = ((int)GAM10H.Value);
            W20.Value = ((int)GAM10H.Value & 0x300) >> 8 | ((int)W20.Value & 0xFC); ;
            W21.Value = ((int)GAM10H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM11H_ValueChanged(object sender, EventArgs e)
        {
            GAM11SL.Value = ((int)GAM11H.Value);
            W22.Value = ((int)GAM11H.Value & 0x300) >> 8 | ((int)W22.Value & 0xFC); ;
            W23.Value = ((int)GAM11H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM12H_ValueChanged(object sender, EventArgs e)
        {
            GAM12SL.Value = ((int)GAM12H.Value);
            W24.Value = ((int)GAM12H.Value & 0x300) >> 8 | ((int)W24.Value & 0xFC); ;
            W25.Value = ((int)GAM12H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM13H_ValueChanged(object sender, EventArgs e)
        {
            GAM13SL.Value = ((int)GAM13H.Value);
            W26.Value = ((int)GAM13H.Value & 0x300) >> 8 | ((int)W26.Value & 0xFC); ;
            W27.Value = ((int)GAM13H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void GAM14H_ValueChanged(object sender, EventArgs e)
        {
            GAM14SL.Value = ((int)GAM14H.Value);
            W28.Value = ((int)GAM14H.Value & 0x300) >> 8 | ((int)W28.Value & 0xFC); ;
            W29.Value = ((int)GAM14H.Value & 0xFF);
            CalculateGAM_VCOM();
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

        private void CBDRVN_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBDRVP.SelectedIndex != -1 && CBDRVN.SelectedIndex != -1)
            {
                W0A.Value = (CBDRVN.SelectedIndex << 3) | CBDRVP.SelectedIndex | ((int)W0A.Value & 0xC0);
            }
        }

        private void VCOM1H_ValueChanged(object sender, EventArgs e)
        {
            VCOM1SL.Value = ((int)VCOM1H.Value);
            W2A.Value = ((int)VCOM1H.Value & 0x300) >> 8 | ((int)W2A.Value & 0xFC); ;
            W2B.Value = ((int)VCOM1H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void VCOM2H_ValueChanged(object sender, EventArgs e)
        {
            VCOM2SL.Value = ((int)VCOM2H.Value);
            W2C.Value = ((int)VCOM2H.Value & 0x300) >> 8 | ((int)W2C.Value & 0xFC); ;
            W2D.Value = ((int)VCOM2H.Value & 0xFF);
            CalculateGAM_VCOM();
        }

        private void VCOM1SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM1H.Value = VCOM1SL.Value;
        }

        private void VCOM2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM2H.Value = VCOM2SL.Value;
        }

        private void BTReadtoWrite_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < WriteTable.Length; i++)
            {
                WriteTable[i].Value = ReadTable[i].Value;
            }
        }

        private void W0E_ValueChanged(object sender, EventArgs e)
        {
            GAM1H.Value = (int)W0F.Value | (((int)W0E.Value & 0x03) << 8);
        }

        private void W10_ValueChanged(object sender, EventArgs e)
        {
            GAM2H.Value = (int)W11.Value | (((int)W10.Value & 0x03) << 8);
        }

        private void W12_ValueChanged(object sender, EventArgs e)
        {
            GAM3H.Value = (int)W13.Value | (((int)W12.Value & 0x03) << 8);
        }

        private void W14_ValueChanged(object sender, EventArgs e)
        {
            GAM4H.Value = (int)W15.Value | (((int)W14.Value & 0x03) << 8);
        }

        private void W16_ValueChanged(object sender, EventArgs e)
        {
            GAM5H.Value = (int)W17.Value | (((int)W16.Value & 0x03) << 8);
        }

        private void W18_ValueChanged(object sender, EventArgs e)
        {
            GAM6H.Value = (int)W19.Value | (((int)W18.Value & 0x03) << 8);
        }

        private void W1A_ValueChanged(object sender, EventArgs e)
        {
            GAM7H.Value = (int)W1B.Value | (((int)W1A.Value & 0x03) << 8);
        }

        private void W1C_ValueChanged(object sender, EventArgs e)
        {
            GAM8H.Value = (int)W1D.Value | (((int)W1C.Value & 0x03) << 8);
        }

        private void W1E_ValueChanged(object sender, EventArgs e)
        {
            GAM9H.Value = (int)W1F.Value | (((int)W1E.Value & 0x03) << 8);
        }

        private void W20_ValueChanged(object sender, EventArgs e)
        {
            GAM10H.Value = (int)W21.Value | (((int)W20.Value & 0x03) << 8);
        }

        private void W22_ValueChanged(object sender, EventArgs e)
        {
            GAM11H.Value = (int)W23.Value | (((int)W22.Value & 0x03) << 8);
        }

        private void W24_ValueChanged(object sender, EventArgs e)
        {
            GAM12H.Value = (int)W25.Value | (((int)W24.Value & 0x03) << 8);
        }

        private void W26_ValueChanged(object sender, EventArgs e)
        {
            GAM13H.Value = (int)W27.Value | (((int)W26.Value & 0x03) << 8);
        }

        private void W28_ValueChanged(object sender, EventArgs e)
        {
            GAM14H.Value = (int)W29.Value | (((int)W28.Value & 0x03) << 8);
        }

        private void W2A_ValueChanged(object sender, EventArgs e)
        {
            VCOM1H.Value = (int)W2B.Value | (((int)W2A.Value & 0x03) << 8);
        }

        private void W2C_ValueChanged(object sender, EventArgs e)
        {
            VCOM2H.Value = (int)W2D.Value | (((int)W2C.Value & 0x03) << 8);
        }

        private void BTWriteBankA_Click(object sender, EventArgs e)
        {
            List<byte> WriteBuf = new List<byte>();
            for (int i = 0; i < WriteTable.Length; i++)
            {
                WriteBuf.Add(Convert.ToByte(WriteTable[i].Value));
            }

            RTDev.I2C_Write((byte)((int)nuSlave.Value >> 1), 0x00, WriteBuf.ToArray());
        }

        private void BTReadBankA_Click(object sender, EventArgs e)
        {
            byte[] ReadBuf = new byte[WriteTable.Length];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), 0x00, ref ReadBuf);

            for(int i = 0; i < ReadTable.Length; i++)
            {
                ReadTable[i].Value = ReadBuf[i];
            }

            byte[] buf = new byte[1];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), 0x64, ref ReadBuf);
            CBMode.SelectedIndex = (ReadBuf[0] & 0x80) >> 7;
        }

        private void ScanSlaveID()
        {
            tbSlave.Invoke((MethodInvoker)(() => tbSlave.Text = ""));
            System.Threading.Thread.Sleep(100);
            List<byte> list =  RTDev.ScanSlaveID();
            if(list == null || list.Count == 0)
            {
                tbSlave.Invoke((MethodInvoker)(() => tbSlave.Text = "No Found Slave Address"));
            }
            else
            {
                string tmp = "Slave Address (8bits) : ";
                for(int i = 0; i < list.Count; i++)
                {
                    if(i == list.Count - 1)
                    {
                        tmp += "0x" + (list[i] << 1).ToString("x").ToUpper();
                    }
                    else
                    {
                        tmp += "0x" + (list[i] << 1).ToString("x").ToUpper() + ", ";
                    }
                }
                tbSlave.Invoke((MethodInvoker)(() => tbSlave.Text = tmp));
                nuSlave.Invoke((MethodInvoker)(() => nuSlave.Value = list[0] << 1));
            }
        }


        private void btScan_Click(object sender, EventArgs e)
        {
            thread = new Thread(ScanSlaveID);
            thread.IsBackground = true;
            thread.Start();
        }

        private void saveBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Bin File|*.bin";
            if(saveDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = saveDlg.FileName;
                List<byte> bin_buf = new List<byte>();
                BinaryWriter bw = new BinaryWriter(new FileStream(file_name, FileMode.Create));

                for(int i = 0; i < 0x100; i++)
                {
                    if(i < WriteTable.Length)
                    {
                        bin_buf.Add(Convert.ToByte(WriteTable[i].Value));
                    }
                    else if(i >= WriteTable.Length && i < WriteTable2.Length + WriteTable.Length)
                    {
                        bin_buf.Add(Convert.ToByte(WriteTable2[i - WriteTable2.Length].Value));
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
            if(openDlg.ShowDialog() == DialogResult.OK)
            {
                byte[] ReadBuf = new byte[255];
                string file_name = openDlg.FileName;
                BinaryReader br = new BinaryReader(new FileStream(file_name, FileMode.Open));

                br.Read(ReadBuf, 0, 0xff);
                
                for(int i = 0; i < 0x100; i++)
                {
                    if(i < WriteTable.Length)
                    {
                        WriteTable[i].Value = ReadBuf[i];
                    }
                    else if(i >= WriteTable.Length && i < WriteTable2.Length + WriteTable.Length)
                    {
                        WriteTable2[i - WriteTable.Length].Value = ReadBuf[i];
                    }
                }
                br.Close();
            }
        }

        private void CKVDD2_CheckedChanged(object sender, EventArgs e)
        {
            if(CBSEL_5V12V2.SelectedIndex != -1)
            {
                byte data = 0x00;
                CheckBox[] _30h_table = new CheckBox[]
                {
                CKVDD2, CKVGL2, CKVSS2, CKControl2, CKAVDD2, CKVGH2
                };

                for (int i = 0; i < _30h_table.Length; i++)
                {
                    data |= (byte)((_30h_table[i].Checked ? 0x01 << i : 0x00));
                }
                W30.Value = data | CBSEL_5V12V2.SelectedIndex << 7 | (int)W30.Value & 0x40;
            }

        }

        private void CBSEL_5V12V2_SelectedIndexChanged(object sender, EventArgs e)
        {
            byte data = 0x00;
            CheckBox[] _30h_table = new CheckBox[]
            {
                CKVDD2, CKVGL2, CKVSS2, CKControl2, CKAVDD2, CKVGH2
            };

            for (int i = 0; i < _30h_table.Length; i++)
            {
                data |= (byte)((_30h_table[i].Checked ? 0x01 << i : 0x00));
            }

            W30.Value = data | CBSEL_5V12V2.SelectedIndex << 7 | (int)W30.Value & 0x40;
        }

        private void CBDLY0_2_SelectedIndexChanged(object sender, EventArgs e)
        {
            byte data = 0;
            ComboBox[] _31h_table = new ComboBox[]
            {
                CBDLY0_2, CBDLY1_2, CBDLY2_2, CBAVDDSS2
            };

            data |= (byte)(CBDLY0_2.SelectedIndex << 0);       // 1:0
            data |= (byte)(CBDLY1_2.SelectedIndex << 2);       // 3:2
            data |= (byte)(CBDLY2_2.SelectedIndex << 4);       // 5:4
            data |= (byte)(CBAVDDSS2.SelectedIndex << 6);      // 6
            W31.Value = data | ((byte)W31.Value & 0x80);
        }

        private void CBSWFreq2_SelectedIndexChanged(object sender, EventArgs e)
        {
            W32.Value = (byte)(CBLSDis2.SelectedIndex << 7) | (byte)(CBSWFreq2.SelectedIndex << 6) | (byte)AVDD2H.Value;
        }

        private void CBAVDDOCP2_SelectedIndexChanged(object sender, EventArgs e)
        {
            W33.Value = (byte)(CBAVDDOCP2.SelectedIndex << 5) | (int)VDD2H.Value | (int)W33.Value & 0xC0;
        }

        private void CBDRVN2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBDRVP2.SelectedIndex != -1 && CBDRVN2.SelectedIndex != -1)
            {
                W3A.Value = (CBDRVN2.SelectedIndex << 3) | CBDRVP2.SelectedIndex | ((int)W3A.Value & 0xC0);
            }

        }

        private void CBPhase2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBLine2.SelectedIndex != -1 && CBPhase2.SelectedIndex != -1)
            {
                W3B.Value = CBLS_Dis2.SelectedIndex << 4 | ((int)W3B.Value & 0x0F);
                W3B.Value = CBPhase2.SelectedIndex << 2 | ((int)W3B.Value & 0xF3);
                W3B.Value = CBLine2.SelectedIndex | ((int)W3B.Value & 0xFC);
            }
        }

        private void CBOCPLevel2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBOCPLevel2.SelectedIndex != -1 &&
                CBOCPDly2.SelectedIndex != -1 &&
                CBLSOCPEnable2.SelectedIndex != -1)
            {
                W3C.Value = CBOCPLevel2.SelectedIndex << 4 | ((int)W3C.Value & 0x0F);
                W3C.Value = CBOCPDly2.SelectedIndex << 1 | ((int)W3C.Value & 0xF1);
                W3C.Value = CBLSOCPEnable2.SelectedIndex | ((int)W3C.Value & 0xFE);
            }
        }

        private void CBDelayFrame2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBDelayFrame2.SelectedIndex != -1 &&
                CBClockCycle2.SelectedIndex != -1)
            {
                W3D.Value = CBDelayFrame2.SelectedIndex << 4 | ((int)W3D.Value & 0x0F);
                W3D.Value = CBClockCycle2.SelectedIndex | ((int)W3D.Value & 0xF0);
            }
        }

        private void AVDD2H_ValueChanged(object sender, EventArgs e)
        {
            byte code = (byte)AVDD2H.Value;
            if(code > 0)
            {
                double volt = (double)(((AVDD2H.Value - 1) + 130) / 10);
                
                if (volt <= 19.2)
                    AVDD2V.Value = (decimal)volt;
            }
            AVDD2SL.Value = (int)AVDD2H.Value;
            int other = ((int)W32.Value & 0xC0);
            W32.Value = ((int)AVDD2H.Value | other);
        }

        private void VDD2H_ValueChanged(object sender, EventArgs e)
        {
            double volt = ((((double)VDD2H.Value * 5) / 100) + 2.2);
            VDD2SL.Value = (int)VDD2H.Value;
            if (volt <= 3.7)
                VDD2V.Value = (decimal)volt;

            W33.Value = (int)VDD2H.Value | ((int)W33.Value & 0xE0);
        }

        private void VSS2H_ValueChanged(object sender, EventArgs e)
        {
            double volt = (((double)VSS2H.Value * 5) + 45) / (-10);
            if (volt >= -16)
                VSS2V.Value = (decimal)volt;

            VSS2SL.Value = (int)VSS2H.Value;
            W34.Value = (int)VSS2H.Value | ((int)W34.Value & 0xE0);
        }

        private void VGL2H_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = ((VGL2H.Value * 5) + 45) / -10;
            VGL2V.Value = volt;
            VGL2SL.Value = (int)VGL2H.Value;
            W36.Value = (int)VGL2H.Value | ((int)W36.Value & 0xE0);
        }

        private void VGH2H_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = VGH2H.Value * 1 + 20;
            if (volt <= 40)
                VGH2V.Value = volt;
            VGH2SL.Value = (int)VGH2H.Value;
            W38.Value = (int)VGH2H.Value | ((int)W38.Value & 0xE0);
        }

        private void OPLDO2H_ValueChanged(object sender, EventArgs e)
        {
            decimal volt = (OPLDO2H.Value * 2 + 130) / 10;
            if (volt <= 18)
                OPLDO2V.Value = volt;
            OPLDO2SL.Value = (int)OPLDO2H.Value;
            W39.Value = (int)OPLDO2H.Value | ((int)W39.Value & 0xE0);
        }

        private void AVDD_OC2H_ValueChanged(object sender, EventArgs e)
        {
            decimal amp = ((AVDD_OC2H.Value * 5) + 20) / 10;
            AVDD_OC2A.Value = amp;
            AVDDOC2SL.Value = (int)AVDD_OC2H.Value;
            W3A.Value = (int)AVDD_OC2H.Value << 6 | ((int)W3A.Value & 0x3F);
        }


        private void GAM1_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM1_2SL.Value = (int)GAM1_2H.Value;
            W3E.Value = ((int)GAM1_2H.Value & 0x300) >> 8 | ((int)W3E.Value & 0xFC);
            W3F.Value = (int)GAM1_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM2_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM2_2SL.Value = (int)GAM2_2H.Value;
            W40.Value = ((int)GAM2_2H.Value & 0x300) >> 8 | ((int)W40.Value & 0xFC);
            W41.Value = (int)GAM2_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM3_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM3_2SL.Value = (int)GAM3_2H.Value;
            W42.Value = ((int)GAM3_2H.Value & 0x300) >> 8 | ((int)W42.Value & 0xFC);
            W43.Value = (int)GAM3_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM4_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM4_2SL.Value = (int)GAM4_2H.Value;
            W44.Value = ((int)GAM4_2H.Value & 0x300) >> 8 | ((int)W44.Value & 0xFC);
            W45.Value = (int)GAM4_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM5_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM5_2SL.Value = (int)GAM5_2H.Value;
            W46.Value = ((int)GAM5_2H.Value & 0x300) >> 8 | ((int)W46.Value & 0xFC);
            W47.Value = (int)GAM5_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM6_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM6_2SL.Value = (int)GAM6_2H.Value;
            W48.Value = ((int)GAM6_2H.Value & 0x300) >> 8 | ((int)W48.Value & 0xFC);
            W49.Value = (int)GAM6_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM7_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM7_2SL.Value = (int)GAM7_2H.Value;
            W4A.Value = ((int)GAM7_2H.Value & 0x300) >> 8 | ((int)W4A.Value & 0xFC);
            W4B.Value = (int)GAM7_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM8_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM8_2SL.Value = (int)GAM8_2H.Value;
            W4C.Value = ((int)GAM8_2H.Value & 0x300) >> 8 | ((int)W4C.Value & 0xFC);
            W4D.Value = (int)GAM8_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM9_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM9_2SL.Value = (int)GAM9_2H.Value;
            W4E.Value = ((int)GAM9_2H.Value & 0x300) >> 8 | ((int)W4E.Value & 0xFC);
            W4F.Value = (int)GAM9_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM10_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM10_2SL.Value = (int)GAM10_2H.Value;
            W50.Value = ((int)GAM10_2H.Value & 0x300) >> 8 | ((int)W50.Value & 0xFC);
            W51.Value = (int)GAM10_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM11_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM11_2SL.Value = (int)GAM11_2H.Value;
            W52.Value = ((int)GAM11_2H.Value & 0x300) >> 8 | ((int)W52.Value & 0xFC);
            W53.Value = (int)GAM11_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM12_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM12_2SL.Value = (int)GAM12_2H.Value;
            W54.Value = ((int)GAM12_2H.Value & 0x300) >> 8 | ((int)W54.Value & 0xFC);
            W55.Value = (int)GAM12_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }


        private void GAM13_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM13_2SL.Value = (int)GAM13_2H.Value;
            W56.Value = ((int)GAM13_2H.Value & 0x300) >> 8 | ((int)W56.Value & 0xFC);
            W57.Value = (int)GAM13_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void GAM14_2H_ValueChanged(object sender, EventArgs e)
        {
            GAM14_2SL.Value = (int)GAM14_2H.Value;
            W58.Value = ((int)GAM14_2H.Value & 0x300) >> 8 | ((int)W58.Value & 0xFC);
            W59.Value = (int)GAM14_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void VCOM1_2H_ValueChanged(object sender, EventArgs e)
        {
            VCOM1_2SL.Value = (int)VCOM1_2H.Value;
            W5A.Value = ((int)VCOM1_2H.Value & 0x300) >> 8 | ((int)W5A.Value & 0xFC);
            W5B.Value = (int)VCOM1_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void VCOM2_2H_ValueChanged(object sender, EventArgs e)
        {
            VCOM2_2SL.Value = (int)VCOM2_2H.Value;
            W5C.Value = ((int)VCOM2_2H.Value & 0x300) >> 8 | ((int)W5C.Value & 0xFC);
            W5D.Value = (int)VCOM2_2H.Value & 0xff;
            CalculateGAM_VCOM();
        }

        private void AVDD2SL_Scroll(object sender, ScrollEventArgs e)
        {
            AVDD2H.Value = AVDD2SL.Value;
        }

        private void VDD2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VDD2H.Value = VDD2SL.Value;
        }

        private void VSS2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VSS2H.Value = VSS2SL.Value;
        }

        private void VGL2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VGL2H.Value = VGL2SL.Value;
        }

        private void VGH2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VGH2H.Value = VGH2SL.Value;
        }

        private void OPLDO2SL_Scroll(object sender, ScrollEventArgs e)
        {
            OPLDO2H.Value = OPLDO2SL.Value;
        }

        private void AVDDOC2SL_Scroll(object sender, ScrollEventArgs e)
        {
            AVDD_OC2H.Value = AVDDOC2SL.Value;
        }


        private void GAM1_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM1_2H.Value = GAM1_2SL.Value;
        }

        private void GAM2_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM2_2H.Value = GAM2_2SL.Value;
        }

        private void GAM3_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM3_2H.Value = GAM3_2SL.Value;
        }

        private void GAM4_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM4_2H.Value = GAM4_2SL.Value;
        }

        private void GAM5_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM5_2H.Value = GAM5_2SL.Value;
        }

        private void GAM6_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM6_2H.Value = GAM6_2SL.Value;
        }

        private void GAM7_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM7_2H.Value = GAM7_2SL.Value;
        }

        private void GAM8_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM8_2H.Value = GAM8_2SL.Value;
        }

        private void GAM9_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM9_2H.Value = GAM9_2SL.Value;
        }

        private void GAM10_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM10_2H.Value = GAM10_2SL.Value;
        }

        private void GAM11_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM11_2H.Value = GAM11_2SL.Value;
        }

        private void GAM12_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM12_2H.Value = GAM12_2SL.Value;
        }

        private void GAM13_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM13_2H.Value = GAM13_2SL.Value;
        }

        private void GAM14_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            GAM14_2H.Value = GAM14_2SL.Value;
        }

        private void VCOM1_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM1_2H.Value = VCOM1_2SL.Value;
        }

        private void VCOM2_2SL_Scroll(object sender, ScrollEventArgs e)
        {
            VCOM2_2H.Value = VCOM2_2SL.Value;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < WriteTable2.Length; i++)
            {
                WriteTable2[i].Value = ReadTable2[i].Value;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<byte> WriteBuf = new List<byte>();
            for (int i = 0; i < WriteTable2.Length; i++)
            {
                WriteBuf.Add(Convert.ToByte(WriteTable2[i].Value));
            }
            RTDev.I2C_Write((byte)((int)nuSlave.Value >> 1), 0x30, WriteBuf.ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            byte[] ReadBuf = new byte[WriteTable.Length];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), 0x30, ref ReadBuf);
            for (int i = 0; i < ReadTable.Length; i++)
            {
                ReadTable2[i].Value = ReadBuf[i];
            }

            byte[] buf = new byte[1];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), 0x64, ref ReadBuf);
            CBMode.SelectedIndex = (ReadBuf[0] & 0x80) >> 7;
        }

        private void W30_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W30.Value;
            CheckBox[] _30h_table = new CheckBox[]
            {
                CKVDD2, CKVGL2, CKVSS2, CKControl2, CKAVDD2, CKVGH2
            };

            for (int i = 0; i < _30h_table.Length; i++)
            {
                if ((data & (0x01 << i)) == bit_table[i])
                {
                    _30h_table[i].Checked = true;
                }
                else
                {
                    _30h_table[i].Checked = false;
                }
            }

            CBSEL_5V12V2.SelectedIndex = (data & 0x80) >> 7;
        }

        private void W31_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W31.Value;

            CBDLY0_2.SelectedIndex = (data & 0x03);
            CBDLY1_2.SelectedIndex = (data & 0x0C) >> 2;
            CBDLY2_2.SelectedIndex = (data & 0x30) >> 4;
            CBAVDDSS2.SelectedIndex = (data & 0x40) >> 6;
        }

        private void W32_ValueChanged(object sender, EventArgs e)
        {
            byte data = (byte)W32.Value;
            AVDD2H.Value = data & 0x3f;
            CBSWFreq2.SelectedIndex = (data & 0x40) >> 6;
            CBLSDis2.SelectedIndex = (data & 0x80) >> 7;
        }

        private void W33_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W33.Value;
            VDD2H.Value = (int)data & 0x1F;
            CBAVDDOCP2.SelectedIndex = ((int)data & 0x20) >> 5;
        }

        private void W34_ValueChanged(object sender, EventArgs e)
        {
            VSS2H.Value = (int)W34.Value & 0x1F;
        }

        private void W36_ValueChanged(object sender, EventArgs e)
        {
            VGL2H.Value = ((int)W36.Value & 0x1F);
        }

        private void VGLSL_Scroll(object sender, ScrollEventArgs e)
        {
            VGLH.Value = VGLSL.Value;
        }

        private void W38_ValueChanged(object sender, EventArgs e)
        {
            VGH2H.Value = (int)W38.Value & 0x1F;
        }

        private void W39_ValueChanged(object sender, EventArgs e)
        {
            OPLDOH.Value = (int)W39.Value & 0x1F;
        }

        private void W3A_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W3A.Value;
            AVDD_OC2H.Value = ((int)data & 0xC0) >> 6;
            CBDRVN2.SelectedIndex = ((int)data & 0x38) >> 3;
            CBDRVP2.SelectedIndex = ((int)data & 0x7);
        }

        private void W3B_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W3B.Value;
            CBLS_Dis2.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBPhase2.SelectedIndex = ((int)data & 0xC) >> 2;
            CBLine2.SelectedIndex = ((int)data & 0x3);
        }

        private void W3C_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W3C.Value;
            CBOCPLevel2.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBOCPDly2.SelectedIndex = ((int)data & 0x0E) >> 1;
            CBLSOCPEnable2.SelectedIndex = ((int)data & 0x01);
        }

        private void W3D_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)W3D.Value;
            CBDelayFrame2.SelectedIndex = ((int)data & 0xF0) >> 4;
            CBClockCycle2.SelectedIndex = ((int)data & 0x0F);
        }

        private void W3E_ValueChanged(object sender, EventArgs e)
        {
            GAM1_2H.Value = (int)W3F.Value | (((int)W3E.Value & 0x03) << 8);
        }

        private void W40_ValueChanged(object sender, EventArgs e)
        {
            GAM2_2H.Value = (int)W41.Value | (((int)W40.Value & 0x03) << 8);
        }

        private void W42_ValueChanged(object sender, EventArgs e)
        {
            GAM3_2H.Value = (int)W43.Value | (((int)W42.Value & 0x03) << 8);
        }

        private void W44_ValueChanged(object sender, EventArgs e)
        {
            GAM4_2H.Value = (int)W45.Value | (((int)W44.Value & 0x03) << 8);
        }

        private void W46_ValueChanged(object sender, EventArgs e)
        {
            GAM5_2H.Value = (int)W47.Value | (((int)W46.Value & 0x03) << 8);
        }

        private void W48_ValueChanged(object sender, EventArgs e)
        {
            GAM6_2H.Value = (int)W49.Value | (((int)W48.Value & 0x03) << 8);
        }

        private void W4A_ValueChanged(object sender, EventArgs e)
        {
            GAM7_2H.Value = (int)W4B.Value | (((int)W4A.Value & 0x03) << 8);
        }

        private void W4C_ValueChanged(object sender, EventArgs e)
        {
            GAM8_2H.Value = (int)W4D.Value | (((int)W4C.Value & 0x03) << 8);
        }

        private void W4E_ValueChanged(object sender, EventArgs e)
        {
            GAM9_2H.Value = (int)W4F.Value | (((int)W4E.Value & 0x03) << 8);
        }

        private void W50_ValueChanged(object sender, EventArgs e)
        {
            GAM10_2H.Value = (int)W51.Value | (((int)W50.Value & 0x03) << 8);
        }

        private void W52_ValueChanged(object sender, EventArgs e)
        {
            GAM11_2H.Value = (int)W53.Value | (((int)W52.Value & 0x03) << 8);
        }

        private void W54_ValueChanged(object sender, EventArgs e)
        {
            GAM12_2H.Value = (int)W55.Value | (((int)W54.Value & 0x03) << 8);
        }

        private void W56_ValueChanged(object sender, EventArgs e)
        {
            GAM13_2H.Value = (int)W57.Value | (((int)W56.Value & 0x03) << 8);

        }

        private void W58_ValueChanged(object sender, EventArgs e)
        {
            GAM14_2H.Value = (int)W59.Value | (((int)W58.Value & 0x03) << 8);

        }

        private void W5A_ValueChanged(object sender, EventArgs e)
        {
            VCOM1_2H.Value = (int)W5B.Value | (((int)W5A.Value & 0x03) << 8);
        }

        private void W5C_ValueChanged(object sender, EventArgs e)
        {
            VCOM2_2H.Value = (int)W5D.Value | (((int)W5C.Value & 0x03) << 8);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            byte[] ReadBuf = new byte[5];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), 0x60, ref ReadBuf);

            R60.Value = (decimal)ReadBuf[0];
            R61.Value = (decimal)ReadBuf[1];
            R62.Value = (decimal)ReadBuf[2];
            R63.Value = (decimal)ReadBuf[3];
            //R64.Value = (decimal)ReadBuf[4];

            byte CRC0, CRC1, CRC2, CRC3, CRC4, CRC5, CRC6, CRC7;
            byte CRC10, CRC11, CRC12, CRC13, CRC14, CRC15, CRC16, CRC17;
            byte R60_data = ReadBuf[0];
            byte R61_data = ReadBuf[1];
            byte R62_data = ReadBuf[2];
            byte R63_data = ReadBuf[3];

            CRC0 = (byte)((R60_data & 0x80) >> 7);
            CRC12 = (byte)((R60_data & 0x20) >> 5);
            CRC5 = (byte)((R60_data & 0x10) >> 4);
            CRC14 = (byte)((R60_data & 0x02) >> 1);

            CRC6 = (byte)((R61_data & 0x20) >> 5);
            CRC2 = (byte)((R61_data & 0x8) >> 3);
            CRC10 = (byte)((R61_data & 0x2) >> 1);
            CRC16 = (byte)((R61_data & 0x1) >> 0);

            CRC13 = (byte)((R62_data & 0x80) >> 7);
            CRC1 = (byte)((R62_data & 0x20) >> 5);
            CRC7 = (byte)((R62_data & 0x08) >> 3);
            CRC3 = (byte)((R62_data & 0x1) >> 0);

            CRC11 = (byte)((R63_data & 0x20) >> 5);
            CRC17 = (byte)((R63_data & 0x10) >> 4);
            CRC4 = (byte)((R63_data & 0x04) >> 2);
            CRC15 = (byte)((R63_data & 0x02) >> 1);

            int CRC = (CRC0 |
                CRC1 << 1 |
                CRC2 << 2 |
                CRC3 << 3 |
                CRC4 << 4 |
                CRC5 << 5 |
                CRC6 << 6 |
                CRC7 << 7) |
                (CRC10 << 0 |
                CRC11 << 1 |
                CRC12 << 2 |
                CRC13 << 3 |
                CRC14 << 4 |
                CRC15 << 5 |
                CRC16 << 6 |
                CRC17 << 7) << 8;

            tbCRC.Text = CRC.ToString("x").ToUpper();
            tbCRC0_7.Text = (CRC & 0xFF).ToString("x").ToUpper();
            tbCRC10_17.Text = ((CRC & 0xFF00) >> 8).ToString("x").ToUpper();


        }

        private void R60_ValueChanged(object sender, EventArgs e)
        {
            decimal R60_data = R60.Value;
            decimal R61_data = R61.Value;
            decimal R62_data = R62.Value;
            decimal R63_data = R63.Value;
            //decimal R64_data = R64.Value;

            //if (((int)R64_data & (0x01 << 7)) != 0)
            //    tbS5_7.BackColor = Color.Red;
            //else
            //    tbS5_7.BackColor = Color.LawnGreen;

            for (int i = 0; i < 8; i++)
            {
                if(((int)R60_data & (0x01 << (7 - i))) != 0)
                    StatusReg1[i].BackColor = Color.LawnGreen;
                else
                    StatusReg1[i].BackColor = Color.Red;

                if (((int)R61_data & (0x01 << (7 - i))) != 0)
                    StatusReg2[i].BackColor = Color.LawnGreen;
                else
                    StatusReg2[i].BackColor = Color.Red;


                if (((int)R62_data & (0x01 << (7 - i))) != 0)
                    StatusReg3[i].BackColor = Color.LawnGreen;
                else
                    StatusReg3[i].BackColor = Color.Red;

                if (((int)R63_data & (0x01 << (7 - i))) != 0)
                    StatusReg4[i].BackColor = Color.LawnGreen;
                else
                    StatusReg4[i].BackColor = Color.Red;
            }
        }

        private void BTSingleWrite_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[] { (byte)singleWriteData.Value };
            RTDev.I2C_Write((byte)((int)nuSlave.Value >> 1), (byte)singleWriteAddr.Value, buf);
        }

        private void BTSingleRead_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[1];
            RTDev.I2C_Read((byte)((int)nuSlave.Value >> 1), (byte)singleReadAddr.Value, ref buf);
            singleReadData.Value = buf[0];
        }

        private void CBMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            byte[] buf = new byte[] { (byte)(CBMode.SelectedIndex << 7) };
            RTDev.I2C_Write((byte)((int)nuSlave.Value >> 1), 0x64, buf);
        }

        private void linkRTBridgeBoardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool status = RTDev.BoadInit();
            if (status)
                MessageBox.Show("Linking RTBridge Board Successful!!!", this.Text);
            else
                MessageBox.Show("Linking RTBridge Board Fail!!!", this.Text);
        }

        private void BT_IntoTestMode_Click(object sender, EventArgs e)
        {
            byte[] tm_code = new byte[]
            {
                0x5A, 0xA5, 0x62, 0x86, 0x68, 0x26, 0x5A, 0xA5
            };

            RTDev.I2C_Write((byte)((int)nuSlave.Value >> 1), 0xF0, tm_code);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            byte[] exit_code = new byte[]
            {
                0x87
            };

            RTDev.I2C_Write((byte)(0xD0 >> 1), 0xF7, exit_code);
        }

        private void BTWriteTM_Click(object sender, EventArgs e)
        {
            List<byte> write_buf = new List<byte>();

            for(int i = 0; i < WriteTMTable.Length; i++)
            {
                write_buf.Add((byte)WriteTMTable[i].Value);
            }
            RTDev.I2C_Write(0xD0 >> 1, 0x70, write_buf.ToArray());

            write_buf.Clear();
            for (int i = 0; i < WriteTMTable2.Length; i++)
            {
                write_buf.Add((byte)WriteTMTable2[i].Value);
            }
            RTDev.I2C_Write(0xD0 >> 1, 0xE0, write_buf.ToArray());
        }

        private void button6_Click(object sender, EventArgs e)
        {
            byte[] write_buf = new byte[] { (byte)comboBox1.SelectedIndex };
            RTDev.I2C_Write(0xD0 >> 1, 0xFC, write_buf);

            byte[] read_buf = new byte[ReadTMTable.Length];

            RTDev.I2C_Read(0xD0 >> 1, 0x70, ref read_buf);
            for(int i = 0; i < read_buf.Length; i++)
            {
                ReadTMTable[i].Value = read_buf[i];
            }

            read_buf = new byte[ReadTMTable2.Length];
            RTDev.I2C_Read(0xD0 >> 1, 0xE0, ref read_buf);
            for (int i = 0; i < read_buf.Length; i++)
            {
                ReadTMTable2[i].Value = read_buf[i];
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ReadTMTable.Length; i++)
            {
                WriteTMTable[i].Value = ReadTMTable[i].Value;
            }


            for (int i = 0; i < ReadTMTable2.Length; i++)
            {
                WriteTMTable2[i].Value = ReadTMTable2[i].Value;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[2];
            // 0xF8 TMD
            // 0xF9 TMG
            buf[0] = (byte)WTMD.Value; // TMD
            buf[1] = (byte)WTMG.Value; // TMG
            RTDev.I2C_Write(0xD0 >> 1, 0xF8, buf);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[2];
            RTDev.I2C_Read(0xD0 >> 1, 0xF8, ref buf);
            // 0xF8 TMD
            // 0xF9 TMG
            RTMD.Value = buf[0]; // TMD
            RTMG.Value = buf[1]; // TMG
        }

        private void button10_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[1];
            buf[0] = (byte)0xA1;
            RTDev.I2C_Write(0xD0 >> 1, 0xFB, buf);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            byte[] buf = new byte[1];
            buf[0] = (byte)0xA8;
            RTDev.I2C_Write(0xD0 >> 1, 0xFB, buf);
        }

        private void bt_bit0_Click(object sender, EventArgs e)
        {
            Button[] bt_arr = new Button[] { bt_bit0, bt_bit1, bt_bit2, bt_bit3, bt_bit4, bt_bit5, bt_bit6, bt_bit7 };
            Button bt = (Button)sender;
            int idx = bt.TabIndex;

            if (bt.Text == "0") bt.Text = "1";
            else if(bt.Text == "1") bt.Text = "0";

            int data = 0;
            for(int i = 0; i < 8; i++)
            {
                if (bt_arr[i].Text == "1")
                    data |= 1 << i;
            }

            singleWriteData.Value = data;
        }

        private void singleWriteData_ValueChanged(object sender, EventArgs e)
        {
            int data = (int)singleWriteData.Value;
            Button[] bt_arr = new Button[] { bt_bit0, bt_bit1, bt_bit2, bt_bit3, bt_bit4, bt_bit5, bt_bit6, bt_bit7 };

            for(int i = 0; i < 8; i++)
            {
                if ((data & (1 << i)) != 0) bt_arr[i].Text = "1";
                else bt_arr[i].Text = "0";
            }
        }

    }
}

