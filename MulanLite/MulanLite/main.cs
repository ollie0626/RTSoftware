using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using System.IO;
using Sunny.UI;

namespace MulanLite
{
    public partial class main : UIForm
    {
        private const int WriteCmd = 0x2D;
        private const int ReadCmd = 0x1E;
        private const int BLUpdateCmd = 0x5A;
        private const int LEDPacketCmd = 0x3C;
        private const int BroadcastCmd = 0xFF;
        RTBBControl RTDev;
        NumericUpDown[] WriteTable;
        NumericUpDown[] ReadTable;
        UITrackBar[] TrackBarTable;
        NumericUpDown[] LEDCHxTable;
        Control[] FileTable;

        private bool initial_wr_en = false;

        public void GUIInit()
        {
            RTDev = new RTBBControl();
            RTDev.BoardInit();

            cb_allowone.SelectedIndex = 0;
            cb_ditheren.SelectedIndex = 1;
            cb_m_factor.SelectedIndex = 0;
            cb_centred.SelectedIndex = 0;
            CiFreq.SelectedIndex = 0;
            RCLK_DIV.SelectedIndex = 3;
            CiEnable.Active = true;
            cb_pulse_rf.SelectedIndex = 3;
            cb_vhr_open.SelectedIndex = 0;
            cb_vhr_short.SelectedIndex = 3;
            cb_open_dgl.SelectedIndex = 2;
            cb_short_dgl.SelectedIndex = 2;
            cb_thresh_clk_missing.SelectedIndex = 0;
            cb_vhr_hyst.SelectedIndex = 2;
            cb_vhr_up.SelectedIndex = 4;
            Host_CRC.Checked = true;

            cb_debug_en.SelectedIndex = 0;
            cb_switch_filter_time.SelectedIndex = 0;
            cb_blanking_time.SelectedIndex = 0;
            cb_co_do_keep0.SelectedIndex = 0;
            cb_debug_out.SelectedIndex = 0;
            cb_cal_modex1.SelectedIndex = 0;
            cb_cal_modex8.SelectedIndex = 0;
            cb_low_drive.SelectedIndex = 0;
            cb_range_x8_x1.SelectedIndex = 0;
            cb_ch_num.SelectedIndex = 0;
            cb_min_count.SelectedIndex = 0;
            initial_wr_en = true;

            bt_crc_en.Style = UIStyle.Gray;
            bt_rdo_en.Style = UIStyle.Gray;
            bt_badlen_en.Style = UIStyle.Gray;
            bt_badadd_en.Style = UIStyle.Gray;
            bt_badid_en.Style = UIStyle.Gray;
            bt_badcmd_en.Style = UIStyle.Gray;


            WriteTable = new NumericUpDown[]
            {
                W00, W01, W02, W03, W04, W05, W06, W07, W08, W09, W0A, W0B, W0C, W0D, W0E, W0F,
                W10, W11, W12, W13, W14, W15, W16, W17, W18, W19, W1A, W1B, W1C, W1D, W1E, W1F,
                W20, W21, W22, W23, W24, W25, W26, W27, W28, W29, W2A, W2B, W2C, W2D, W2E, W2F,
                W30, W31, W32, W33, W34, W35, W36, W37, W38, W39, W3A, W3B, W3C, W3D, W3E, W3F,
                W40, W41, W42, W43, W44, W45, W46, W47, W48, W49, W4A, W4B, W4C, W4D, W4E, W4F,
                W50, W51, W52, W53, W54, W55, W56, W57, W58, W59, W5A, W5B, W5C, W5D, W5E, W5F
                // W60, W61, W62
            };

            ReadTable = new NumericUpDown[]
            {
                R00, R01, R02, R03, R04, R05, R06, R07, R08, R09, R0A, R0B, R0C, R0D, R0E, R0F,
                R10, R11, R12, R13, R14, R15, R16, R17, R18, R19, R1A, R1B, R1C, R1D, R1E, R1F,
                R20, R21, R22, R23, R24, R25, R26, R27, R28, R29, R2A, R2B, R2C, R2D, R2E, R2F,
                R30, R31, R32, R33, R34, R35, R36, R37, R38, R39, R3A, R3B, R3C, R3D, R3E, R3F,
                R40, R41, R42, R43, R44, R45, R46, R47, R48, R49, R4A, R4B, R4C, R4D, R4E, R4F,
                R50, R51, R52, R53, R54, R55, R56, R57, R58, R59, R5A, R5B, R5C, R5D, R5E, R5F
                // R60, R61, R62
            };

            TrackBarTable = new UITrackBar[]
            {
                trackCH0x8SL, trackCH1x8SL, trackCH2x8SL, trackCH3x8SL, trackCH0x1SL, trackCH1x1SL, trackCH2x1SL, trackCH3x1SL
            };

            LEDCHxTable = new NumericUpDown[]
            {
                nu_CH0x8, nu_CH1x8, nu_CH2x8, nu_CH3x8, nu_CH0x1, nu_CH1x1, nu_CH2x1, nu_CH3x1
            };

            FileTable = new Control[]
            {
                nu_persentid, nuFirst, nuEnd, RCLK_DIV, CiEnable, Host_CRC, cb_allowone, cb_ditheren, cb_m_factor, nuPWMcycle, nuMaxpulse, nuMinpulse, cb_centred,
                nuCy0, nuCy1, nuCy2, nuCy3, nuCy4, nuCy5, nuCy6, nuCy7, ck_short_mask, ck_open_mask, ck_clk_missing, ck_fuse_mask, ck_tsd_mask,
                nu_mulan_qty, nu_start_offset, nu_startid, nu_endid, nu_data, nu_speciedid, nu_start_zone, nu_spe_offset1, nu_spe_offset2, nu_spe_offset2, nu_spe_offset3, nu_spe_offset4,
                nu_specified_data, nu_fault_qty, trackCH0x8SL, trackCH1x8SL, trackCH2x8SL, trackCH3x8SL, trackCH0x1SL, trackCH1x1SL, trackCH2x1SL, trackCH3x1SL,
                cb_thresh_clk_missing, cb_pulse_rf, cb_debug_en, cb_debug_out, cb_vhr_open, cb_vhr_short, cb_vhr_hyst, cb_vhr_up,
                cb_switch_filter_time, cb_open_dgl, cb_ch_num, cb_cal_modex1, cb_cal_modex8, cb_low_drive, cb_range_x8_x1, cb_min_count
            };
            initial_wr_en = true;
        }

        public main()
        {
            InitializeComponent();
            GUIInit();
        }

        private Task<int> WDataTask(byte id, byte addr, byte len ,byte[] buf)
        {
            if(initial_wr_en == false)
            {
                return Task.Factory.StartNew(() => 0);
            }
            return Task.Factory.StartNew(() => RTDev.WriteFunc(id, WriteCmd, addr, len, buf));
        }

        private Task<byte[]> RDataTask(byte id, byte len, byte addr)
        {
            return Task.Factory.StartNew(() => RTDev.ReadFunc(id, len, addr));
        }

        private async void bt_config_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            // identify
            RTDev.Identify((byte)nuFirst.Value);

            uiProcessBar1.Value = 0;
            uiProcessBar2.Value = 0;
            uiProcessBar1.Maximum = 3;
            uiProcessBar2.Maximum = 3;

            // broadcast write config
            // len follow n + 1 format
            byte[] buffer = new byte[15];
            buffer[0] = (byte)(cb_ditheren.SelectedIndex << 4 |
                               cb_m_factor.SelectedIndex << 5 |
                               cb_centred.SelectedIndex << 7);
            buffer[1] = (byte)(cb_allowone.SelectedIndex);
            await WDataTask(0xff, 0x32, 1, buffer); /* id, addr, len, data */
            uiProcessBar2.Value += 1;
            uiProcessBar1.Value += 1;
            System.Threading.Thread.Sleep(15);

            buffer[0] = (byte)((int)nuPWMcycle.Value & 0xFF);
            buffer[1] = (byte)(((int)nuPWMcycle.Value & 0xFF00) >> 8);
            buffer[3] = (byte)((int)nuMaxpulse.Value & 0xFF);
            buffer[4] = (byte)(((int)nuMaxpulse.Value & 0xFF00) >> 8);
            buffer[5] = (byte)nuMinpulse.Value;
            await WDataTask(0xff, 0x36, 4, buffer);
            uiProcessBar2.Value += 1;
            uiProcessBar1.Value += 1;
            System.Threading.Thread.Sleep(15);

            buffer[0] = (byte)nuCy0.Value;
            buffer[1] = (byte)nuCy1.Value;
            buffer[2] = (byte)nuCy2.Value;
            buffer[3] = (byte)nuCy3.Value;
            buffer[4] = (byte)nuCy4.Value;
            buffer[5] = (byte)nuCy5.Value;
            buffer[6] = (byte)nuCy6.Value;
            buffer[7] = (byte)nuCy7.Value;
            await WDataTask(0xff, 0x50, 7, buffer);
            uiProcessBar2.Value += 1;
            uiProcessBar1.Value += 1;
            System.Threading.Thread.Sleep(15);

            // set EOC
            int last_id = (int)nuEnd.Value;
            byte[] data = new byte[1];
            data[0] = 0x01;
            RTDev.WriteFunc((byte)last_id, WriteCmd, 0x32, 0x00, data);
            bt.Enabled = true;
        }

        private void CiFreq_SelectedIndexChanged(object sender, EventArgs e)
        {
            RTDev.SetCiClock(CiFreq.SelectedIndex);
            double ci_freq = 18000000;
            switch (CiFreq.SelectedIndex)
            {
                case 0:
                    ci_freq = 18000000;
                    break;
                case 1:
                    ci_freq = 7800000;
                    break;
                case 2:
                    ci_freq = 6000000;
                    break;
                case 3:
                    ci_freq = 7000000;
                    break;
                case 4:
                    ci_freq = 6000000;
                    break;
            }
            nuPWMout.Value = (decimal)((ci_freq / (double)nuPWMcycle.Value) / 1000);
        }

        private async void bt_open_ch4_Click(object sender, EventArgs e)
        {
            byte id = (byte)nu_persentid.Value;
            byte[] Rdbuffer = new byte[1];
            Rdbuffer = await RDataTask(id, 0, 0x7);

            nuopen_ch4.Value = (Rdbuffer[2] & 0x80) >> 7;
            nuopen_ch3.Value = (Rdbuffer[2] & 0x40) >> 6;
            nuopen_ch2.Value = (Rdbuffer[2] & 0x20) >> 5;
            nuopen_ch1.Value = (Rdbuffer[2] & 0x10) >> 4;

            nushort_ch4.Value = (Rdbuffer[2] & 0x08) >> 3;
            nushort_ch3.Value = (Rdbuffer[2] & 0x04) >> 2;
            nushort_ch2.Value = (Rdbuffer[2] & 0x02) >> 1;
            nushort_ch1.Value = (Rdbuffer[2] & 0x01) >> 0;

            Rdbuffer = await RDataTask(id, 0, 0x04);
            nu_dont_lower.Value = (Rdbuffer[2] & 0x40) >> 6;
            nu_raise.Value = (Rdbuffer[2] & 0x20) >> 5;
        }

        private async void bt_w1c_open4_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            int idx = bt.TabIndex;
            byte id = (byte)nu_persentid.Value;
            byte[] Wrbuffer = new byte[1];
            Wrbuffer[0] = (byte)(0x01 << idx);
            await WDataTask(id, 0x07, 0, Wrbuffer);
        }

        private async void ck_short_mask_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox ck = (CheckBox)sender;
            int idx = ck.TabIndex;
            byte id = (byte)nu_persentid.Value;

            byte[] Rdbuffer = new byte[1];
            Rdbuffer = await RDataTask(id, 0, 0x30);

            byte[] Wrbuffer = new byte[1];
            byte data = (byte)(1 << idx);
            Wrbuffer[0] = ck.Checked ? (byte)(Rdbuffer[2] | data) : (byte)(Rdbuffer[2] & ~data);
            await WDataTask(id, 0x30, 0, Wrbuffer);
        }

        private async void bt_crc_en_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            int idx = bt.TabIndex;
            byte id = (byte)nu_persentid.Value;
            byte[] Rdbuffer = new byte[1];
            byte[] Wrbuffer = new byte[1];
            byte data = (byte)(1 << idx);
            Rdbuffer = await RDataTask(id, 0x00, 0x31);
            Rdbuffer[2] = (byte)(Rdbuffer[2] & ~data);
            if (bt.Style == UIStyle.Gray)
            {
                bt.Style = UIStyle.LightBlue;
                Wrbuffer[0] = (byte)(Rdbuffer[2] | data);
                await WDataTask(id, 0x31, 0x0, Wrbuffer);
            }
            else
            {
                bt.Style = UIStyle.Gray;
                Wrbuffer[0] = (byte)(Rdbuffer[2] & ~data);
                await WDataTask(id, 0x31, 0x0, Wrbuffer);
            }
            bt.Enabled = true;
        }

        private async void bt_rd_crc_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            byte id = (byte)nu_persentid.Value;
            byte[] Rdbuffer = new byte[1];
            Rdbuffer = await RDataTask(id, 0x00, 0x06);

            byte bit5, bit4, bit3, bit2, bit1, bit0;
            bit5 = (byte)((Rdbuffer[2] & 0x20) >> 5);
            bit4 = (byte)((Rdbuffer[2] & 0x10) >> 4);
            bit3 = (byte)((Rdbuffer[2] & 0x08) >> 3);
            bit2 = (byte)((Rdbuffer[2] & 0x04) >> 2);
            bit1 = (byte)((Rdbuffer[2] & 0x02) >> 1);
            bit0 = (byte)((Rdbuffer[2] & 0x01) >> 0);

            nu_crc.Value = bit5;
            nu_rdo.Value = bit4;
            nu_badlen.Value = bit3;
            nu_badadd.Value = bit2;
            nu_badid.Value = bit1;
            nu_badcmd.Value = bit0;
        }

        private async void bt_w1c_rcr_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            int idx = bt.TabIndex;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(0x01 << idx);
            byte[] Wrbuffer = new byte[1] { data };
            await WDataTask(id, 0x06, 0x00, Wrbuffer);
        }

        private void bt_blenable_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte[] WData = new byte[1] { 0x02 };
            RTDev.WriteFunc(id, 0x2D, 0x03, 0x00, WData);
            bt.Enabled = true;
        }

        private void bt_bldisable_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte[] WData = new byte[1] { 0x00 };
            RTDev.WriteFunc(id, 0x2D, 0x03, 0x00, WData);
            bt.Enabled = true;
        }

        private async void bt_usLed_control_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            byte[] ZoneNum = new byte[8];
            byte[] startOffset = new byte[8];
            int offset = (int)nu_start_offset.Value;
            int max = (int)(nu_mulan_qty.Value) + (int)(nu_endid.Value - nu_startid.Value);
            uiProcessBar1.Value = 0;
            uiProcessBar2.Value = 0;
            uiProcessBar1.Maximum = max;
            uiProcessBar2.Maximum = max;

            for (int i = (int)nuFirst.Value; i < nu_mulan_qty.Value; i++)
            {
                int zone1 = ((i * 4) + 1) * offset;
                int zone2 = ((i * 4) + 2) * offset;
                int zone3 = ((i * 4) + 3) * offset;
                int zone4 = ((i * 4) + 4) * offset;

                ZoneNum[0] = (byte)(zone1 & 0xFF);
                ZoneNum[1] = (byte)((zone1 & 0xFF00) >> 8);

                ZoneNum[2] = (byte)(zone2 & 0xFF);
                ZoneNum[3] = (byte)((zone2 & 0xFF00) >> 8);
                
                ZoneNum[4] = (byte)(zone3 & 0xFF);
                ZoneNum[5] = (byte)((zone3 & 0xFF00) >> 8);

                ZoneNum[6] = (byte)(zone4 & 0xFF);
                ZoneNum[7] = (byte)((zone4 & 0xFF00) >> 8);                                
                // mulna zone address start from 0x18
                // and mulan-lite zone address start from 0x10
                //RTDev.WriteFunc((byte)(i + 1), 0x2D, 0x10, 7, ZoneNum);

                await WDataTask((byte)(i), 0x10, 7, ZoneNum);
                uiProcessBar1.Value += 1;
                uiProcessBar2.Value += 1;
            }

            int led_data = (int)nu_data.Value;
            byte bit16 = (byte)((led_data & 0x10000) >> 16);
            byte MSB = (byte)((led_data & 0xFF00) >> 8);
            byte LSB = (byte)(led_data & 0xFF);
            byte[] buffer = new byte[3] { LSB, MSB, bit16 };

            for(int id = (int)nu_startid.Value; id < nu_endid.Value; id++)
            {
                await WDataTask((byte)id, 0x0B, 2, buffer);
                uiProcessBar1.Value += 1;
                uiProcessBar2.Value += 1;
            }
            bt.Enabled = true;
        }

        private async void bt_specified_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            // write data buffer parameter
            byte id = (byte)nu_speciedid.Value;
            int data = (int)nu_specified_data.Value;
            byte MSB = (byte)((data & 0xFF00) >> 8);
            byte LSB = (byte)(data & 0xFF);
            byte bit16 = (byte)((data & 0x10000) >> 16);
            byte[] buffer = new byte[3] { LSB, MSB, bit16 };
            int max = 3;
            uiProcessBar1.Value = 0;
            uiProcessBar2.Value = 0;
            uiProcessBar1.Maximum = max;
            uiProcessBar2.Maximum = max;

            // zone number setting
            int start_zone = (int)nu_start_zone.Value;
            int zone1 = start_zone;
            int zone2 = start_zone + 1;
            int zone3 = start_zone + 2;
            int zone4 = start_zone + 3;
            byte[] ZoneNum = new byte[8];
            ZoneNum[0] = (byte)((zone1 & 0x1fe0) >> 5);
            ZoneNum[1] = (byte)(zone1 & 0x1f);
            ZoneNum[2] = (byte)((zone2 & 0x1fe0) >> 5);
            ZoneNum[3] = (byte)(zone2 & 0x1f);
            ZoneNum[4] = (byte)((zone3 & 0x1fe0) >> 5);
            ZoneNum[5] = (byte)(zone3 & 0x1f);
            ZoneNum[6] = (byte)((zone4 & 0x1fe0) >> 5);
            ZoneNum[7] = (byte)(zone4 & 0x1f);
            await WDataTask(id, 0x10, 7, ZoneNum);
            System.Threading.Thread.Sleep(50);
            uiProcessBar1.Value += 1;
            uiProcessBar2.Value += 1;
            await WDataTask(id, 0x0B, 2, buffer);
            System.Threading.Thread.Sleep(50);
            uiProcessBar1.Value += 1;
            uiProcessBar2.Value += 1;

            // for start offset setting
            int start_offset1 = (int)nu_spe_offset1.Value;
            int start_offset2 = (int)nu_spe_offset2.Value;
            int start_offset3 = (int)nu_spe_offset3.Value;
            int start_offset4 = (int)nu_spe_offset4.Value;
            buffer = new byte[8]{ 
                (byte)(start_offset1 & 0xFF), (byte)((start_offset1 & 0xFF00) >> 8),
                (byte)(start_offset2 & 0xFF), (byte)((start_offset2 & 0xFF00) >> 8),
                (byte)(start_offset4 & 0xFF), (byte)((start_offset4 & 0xFF00) >> 8),
                (byte)(start_offset3 & 0xFF), (byte)((start_offset3 & 0xFF00) >> 8),
            };
            await WDataTask(id, 0x18, 7, buffer);
            uiProcessBar1.Value += 1;
            uiProcessBar2.Value += 1;
            bt.Enabled = true;
        }

        private void bt_inquiry_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            string[] FlagName_talbe = new string[]{
                "Not Ready",                /* 0 */
                "Disable LED",              /* 1 */
                "Therm SHUT",               /* 2 */
                "EFUSE CRCERR",             /* 3 */
                "CLOCK MISSING",            /* 4 */
                "RAISE",                    /* 5 */
                "DONT LOWER",               /* 6 */
                "COMM ERR",                 /* 7 */
                "LATE UPD",                 /* 8 */
                "OPEN",                     /* 9 */
                "SHORT",                    /* 10 */
                "SMALL BLANKED",            /* 11 */
                "BIG BLANKED"               /* 12 */
            };
            
            NumericUpDown[] FlagTable = new NumericUpDown[]
            {
                flag1, flag2, flag3, flag4, flag5, flag6, flag7, flag8, flag9, flag10, flag11, flag12, flag13
            };
            byte[] RData = RTDev.Inquiry();
            int flag = RData[1] | (RData[2] << 8);
            if(flag != 0x00)
            {
                for(int i = 0; i < FlagTable.Length; i++)
                {
                    if((flag & (1 << i)) == (1 << i))
                        FlagTable[i].Value = 1;
                    else
                        FlagTable[i].Value = 0;
                }
            }
            else
            {
                for(int i = 0; i < FlagTable.Length; i++) FlagTable[i].Value = 0;
            }
            bt.Enabled = true;
        }

        private void bt_repones_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            NumericUpDown[] FlagTable = new NumericUpDown[]
            {
                flag1, flag2, flag3, flag4, flag5, flag6, flag7, flag8, flag9, flag10, flag11, flag12, flag13
            };

            TextBox[] showFlagTable = new TextBox[]
            {
                textBox1, textBox2, textBox3, textBox4, textBox5, textBox6, textBox7, textBox8, textBox9, textBox10, textBox11, textBox12, textBox13
            };

            for(int flag_idx = 0; flag_idx < FlagTable.Length; flag_idx++)
            {
                showFlagTable[flag_idx].Text = "ID:";
                if(FlagTable[flag_idx].Value == 1)
                {
                    for(int i = 0; i < nu_fault_qty.Value; i++)
                    {
                        //7:4 = Target FLAG number 0..15. 3:0 = ~FLAG
                        byte flag = (byte)(((flag_idx & 0x0F) << 4) | (~flag_idx & 0x0F));
                        byte[] RData = RTDev.ResponesID(flag);
                        if (RData[1] == 0xff)
                        {
                            if(nu_fault_qty.Value == 1) showFlagTable[flag_idx].Text += "0x" + RData[2].ToString("X");
                            else                        showFlagTable[flag_idx].Text += "0x" + RData[2].ToString("X") + ", ";
                        }
                    }
                }
            }
            bt.Enabled = true;
        }

        private void Host_CRC_CheckedChanged(object sender, EventArgs e)
        {
            if (Host_CRC.Checked) RTBBControl.CRC_En = true;
            else RTBBControl.CRC_En = false;
        }

        private void RCLK_DIV_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte[] RData = RTDev.ReadFunc(id, 0x00, 0x34);
            byte[] WData = new byte[1];
            WData[0] = (byte)((RCLK_DIV.SelectedIndex << 4) | (RData[2] & 0x0F));
            RTDev.WriteFunc(id, WriteCmd, 0x34, 0x00, WData);
            cb.Enabled = true;
        }

        private void bt_readtowrite_Click(object sender, EventArgs e)
        {
            for(int i = 0; i < WriteTable.Length; i++)
            {
                WriteTable[i].Value = ReadTable[i].Value;
            }
        }

        private void bt_writeall_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            int max = 6;
            uiProcessBar1.Maximum = max;
            uiProcessBar2.Maximum = max;
            //for (int i = 0; i < max; i++)
            //{
            //    byte[] Data = new byte[16];
            //    Data[0] = (byte)WriteTable[(i * 16) + 0].Value;
            //    Data[1] = (byte)WriteTable[(i * 16) + 1].Value;
            //    Data[2] = (byte)WriteTable[(i * 16) + 2].Value;
            //    Data[3] = (byte)WriteTable[(i * 16) + 3].Value;
            //    Data[4] = (byte)WriteTable[(i * 16) + 4].Value;
            //    Data[5] = (byte)WriteTable[(i * 16) + 5].Value;
            //    Data[6] = (byte)WriteTable[(i * 16) + 6].Value;
            //    Data[7] = (byte)WriteTable[(i * 16) + 7].Value;
            //    Data[8] = (byte)WriteTable[(i * 16) + 8].Value;
            //    Data[9] = (byte)WriteTable[(i * 16) + 9].Value;
            //    Data[10] = (byte)WriteTable[(i * 16) + 10].Value;
            //    Data[11] = (byte)WriteTable[(i * 16) + 11].Value;
            //    Data[12] = (byte)WriteTable[(i * 16) + 12].Value;
            //    Data[13] = (byte)WriteTable[(i * 16) + 13].Value;
            //    Data[14] = (byte)WriteTable[(i * 16) + 14].Value;
            //    Data[15] = (byte)WriteTable[(i * 16) + 15].Value;
            //    await WDataTask(id, (byte)(i * 16), (byte)15, Data);
            //    uiProcessBar1.Value += 1;
            //    uiProcessBar2.Value += 1;
            //}


            //RTDev.WriteFunc((byte)nuSID.Value, WriteCmd, (byte)nuSAddr.Value, (byte)0x01, write_buf);
            byte[] data = new byte[WriteTable.Length];
            for (int i = 0; i < data.Length; i++) data[i] = (byte)WriteTable[i].Value;
            RTDev.WriteFunc(id, WriteCmd, (byte)0x00, data.Length, data);
            bt.Enabled = true;
        }

        private async void bt_readall_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            int max = 12;
            byte[] Before = new byte[ReadTable.Length];
            byte[] After = new byte[ReadTable.Length];
            for (int i = 0; i < Before.Length; i++) Before[i] = (byte)ReadTable[i].Value;
            byte[] RData = new byte[7];
            byte id = (byte)nu_persentid.Value;

            uiProcessBar1.Value = 0;
            uiProcessBar2.Value = 0;
            uiProcessBar1.Maximum = max;
            uiProcessBar2.Maximum = max;

            for(int i = 0; i < max; i++)
            {
                byte addr = (byte)(i * 8);
                RData = await RDataTask(id, 7, addr);
                ReadTable[i * 8 + 0].Value = RData[2];
                ReadTable[i * 8 + 1].Value = RData[3];
                ReadTable[i * 8 + 2].Value = RData[4];
                ReadTable[i * 8 + 3].Value = RData[5];
                ReadTable[i * 8 + 4].Value = RData[6];
                ReadTable[i * 8 + 5].Value = RData[7];
                ReadTable[i * 8 + 6].Value = RData[8];
                ReadTable[i * 8 + 7].Value = RData[9];

                ReadTable[i * 8 + 0].BackColor = Before[i * 8 + 0] != RData[2] ? Color.Red : Color.White;
                ReadTable[i * 8 + 1].BackColor = Before[i * 8 + 1] != RData[3] ? Color.Red : Color.White;
                ReadTable[i * 8 + 2].BackColor = Before[i * 8 + 2] != RData[4] ? Color.Red : Color.White;
                ReadTable[i * 8 + 3].BackColor = Before[i * 8 + 3] != RData[5] ? Color.Red : Color.White;
                ReadTable[i * 8 + 4].BackColor = Before[i * 8 + 4] != RData[6] ? Color.Red : Color.White;
                ReadTable[i * 8 + 5].BackColor = Before[i * 8 + 5] != RData[7] ? Color.Red : Color.White;
                ReadTable[i * 8 + 6].BackColor = Before[i * 8 + 6] != RData[8] ? Color.Red : Color.White;
                ReadTable[i * 8 + 7].BackColor = Before[i * 8 + 7] != RData[9] ? Color.Red : Color.White;
                System.Threading.Thread.Sleep(5);
                uiProcessBar1.Value += 1;
                uiProcessBar2.Value += 1;
            }
            bt.Enabled = true;
        }

        private async void WRReg(byte id, byte mask, byte addr, byte data)
        {
            try
            {
                if (initial_wr_en == false) return;

                byte[] RData = await RDataTask(id, 0x00, addr);
                byte Wrin = (byte)((RData[0] & mask) | data);
                byte[] WData = new byte[1] { Wrin };
                await WDataTask(id, addr, 0x00, WData);
            }
            catch
            {
                Console.WriteLine("WRReg Func error");
            }
        }

        private void cb_m_factor_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x32;
            byte mask = 0x9F;
            byte data = (byte)(cb_m_factor.SelectedIndex << 5);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_allowone_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x33;
            byte mask = 0xFE;
            byte data = (byte)cb_allowone.SelectedIndex;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_ditheren_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x32;
            byte mask = 0xEF;
            byte data = (byte)cb_allowone.SelectedIndex;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_centred_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x32;
            byte mask = 0x7F;
            byte data = (byte)(cb_centred.SelectedIndex << 7);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private async void nuPWMcycle_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown nu = (NumericUpDown)sender;
            nu.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            int Data = (int)nuPWMcycle.Value;
            byte addr = 0x36;
            byte len = 1;
            byte MSB = (byte)(Data & 0xFF);
            byte LSB = (byte)((Data & 0xFF00) >> 8);
            byte[] WData = new byte[2] { MSB, LSB };
            await WDataTask(id, addr, len, WData);
            nu.Enabled = true;
        }

        private async void nuMaxpulse_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown nu = (NumericUpDown)sender;
            nu.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            int Data = (int)nuPWMcycle.Value;
            byte addr = 0x38;
            byte len = 1;
            byte MSB = (byte)(Data & 0xFF);
            byte LSB = (byte)((Data & 0xFF00) >> 8);
            byte[] WData = new byte[2] { MSB, LSB };
            await WDataTask(id, addr, len, WData);
            nu.Enabled = true;
        }

        private void nuMinpulse_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown nu = (NumericUpDown)sender;
            nu.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x00;
            switch(nu.TabIndex)
            {
                case 0: addr = 0x3A; break;
                case 1: addr = 0x50; break;
                case 2: addr = 0x51; break;
                case 3: addr = 0x52; break;
                case 4: addr = 0x53; break;
                case 5: addr = 0x54; break;
                case 6: addr = 0x55; break;
                case 7: addr = 0x56; break;
                case 8: addr = 0x57; break;
            }
            byte mask = 0x00;
            byte data = (byte)(nu.Value);
            WRReg(id, mask, addr, data);
            nu.Enabled = true;
        }

        private async void trackCH0x8SL_ValueChanged(object sender, EventArgs e)
        {
            List<byte> DataList = new List<byte>();
            for(int i = 0; i < TrackBarTable.Length; i++)
            {
                LEDCHxTable[i].Value = TrackBarTable[i].Value;
                DataList.Add((byte)LEDCHxTable[i].Value);
            }
            byte[] WData = DataList.ToArray();
            byte id = (byte)nu_persentid.Value;
            await WDataTask(id, 0x28, 7, WData);
        }

        private void nu_CH0x8_ValueChanged(object sender, EventArgs e)
        {
            for(int i = 0; i < LEDCHxTable.Length; i++)
            {
                TrackBarTable[i].Value = (int)LEDCHxTable[i].Value;
            }
        }

        private void cb_thresh_clk_missing_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x4B;
            byte mask = 0xFC;
            byte data = (byte)(cb_thresh_clk_missing.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_pulse_rf_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x44;
            byte mask = 0xFC;
            byte data = (byte)(cb_pulse_rf.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_vhr_open_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x45;
            byte mask = 0xCF;
            byte data = (byte)(cb_vhr_open.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_vhr_short_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x45;
            byte mask = 0xFC;
            byte data = (byte)(cb_vhr_short.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_vhr_hyst_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x46;
            byte mask = 0x8F;
            byte data = (byte)(cb_vhr_hyst.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_vhr_up_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x46;
            byte mask = 0xF8;
            byte data = (byte)(cb_vhr_hyst.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_open_dgl_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x47;
            byte mask = 0xF3;
            byte data = (byte)(cb_open_dgl.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_short_dgl_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x47;
            byte mask = 0xFC;
            byte data = (byte)(cb_short_dgl.SelectedIndex);
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void CiEnable_ValueChanged(object sender, bool value)
        {
            if(CiEnable.Active == true) RTDev.CiEnable();
            else                        RTDev.CiDisable();
        }

        private void track_bl_late_ValueChanged(object sender, EventArgs e)
        {
            nu_bl_late.Value = track_bl_late.Value;

            byte id = (byte)nu_persentid.Value;
            byte addr = 0x4A;
            byte mask = 0x00;
            byte data = (byte)track_bl_late.Value;
            WRReg(id, mask, addr, data);

        }

        private void nu_bl_late_ValueChanged(object sender, EventArgs e)
        {
            track_bl_late.Value = (int)nu_bl_late.Value;
        }

        private void cb_switch_filter_time_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)((cb_switch_filter_time.SelectedIndex << 4) | cb_blanking_time.SelectedIndex);
            byte addr = 0x58;
            byte mask = 0x00;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_debug_en_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_debug_en.SelectedIndex << 7 | cb_co_do_keep0.SelectedIndex << 6 | cb_debug_out.SelectedIndex);
            byte addr = 0x5E;
            byte mask = 0x00;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private async void bt_user_test_mode_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            bt.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte addr = 0x60;
            byte[] RData = await RDataTask(id, 0x01, addr);

            // get status data index = 2, 3
            byte data = RData[2]; // 0x60 data;
            byte bit7 = (byte)((data & 0x80) >> 7);
            byte bit6 = (byte)((data & 0x40) >> 6);
            byte bit5 = (byte)((data & 0x20) >> 5);
            byte bit4 = (byte)((data & 0x10) >> 4);
            byte bit3 = (byte)((data & 0x08) >> 3);
            byte bit2 = (byte)((data & 0x04) >> 2);
            byte bit1 = (byte)((data & 0x02) >> 1);
            byte bit0 = (byte)((data & 0x01) >> 0);
            nu_test_mode.Value = bit5;
            nu_stat_dis.Value = bit4;
            nu_stat_norm.Value = bit3;
            nu_stat_stb.Value = bit2;
            nu_stat_iden.Value = bit1;
            nu_stat_init.Value = bit0;

            data = RData[3]; // 0x61 data;
            bit7 = (byte)((data & 0x80) >> 7);
            bit6 = (byte)((data & 0x40) >> 6);
            bit5 = (byte)((data & 0x20) >> 5);
            bit4 = (byte)((data & 0x10) >> 4);
            bit3 = (byte)((data & 0x08) >> 3);
            bit2 = (byte)((data & 0x04) >> 2);
            bit1 = (byte)((data & 0x02) >> 1);
            bit0 = (byte)((data & 0x01) >> 0);
            nu_efuse_load.Value = bit2;
            nu_tsd_mask.Value = bit1;
            nu_tsd.Value = bit0;
        
            bt.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            RTDev.BoardInit();
            timer1.Enabled = false;
            Console.WriteLine("one time shot timer !!!");
        }

        private void main_Load(object sender, EventArgs e)
        {
            //Guid hidGuid = new Guid("745a17a0-74d3-11d0-b6fe-00a0c90f57da");
            //Dbt.HidD_GetHidGuid(ref hidGuid);
            //RegisterNotification(hidGuid);
            timer1.Interval = 500;
            timer1.Enabled = false;
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            //UnregisterNotification();
        }

        private void cb_ch_num_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)cb_ch_num.SelectedIndex;
            byte addr = 0x0A;
            byte mask = 0x00;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_cal_modex1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_cal_modex1.SelectedIndex << 2);
            byte addr = 0x32;
            byte mask = 0xFB;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;   
        }

        private void cb_cal_modex8_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_cal_modex1.SelectedIndex << 2);
            byte addr = 0x32;
            byte mask = 0xFD;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_low_drive_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_low_drive.SelectedIndex << 7);
            byte addr = 0x33;
            byte mask = 0x7F;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_range_x8_x1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_range_x8_x1.SelectedIndex << 6);
            byte addr = 0x33;
            byte mask = 0xBF;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void cb_min_count_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_min_count.SelectedIndex);
            byte addr = 0x35;
            byte mask = 0xF0;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void bt_blupdate_Click(object sender, EventArgs e)
        {
            RTDev.BLUpdate();
        }

        private void bt_openbin_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Open Setting|*.bin";
            openDlg.Title = "Open Mulan Lite Setting";
            if(openDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = openDlg.FileName;
                read_setting(file_name);
            }
            
        }

        private void bt_savebin_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Save Setting|*.bin";
            saveDlg.Title = "Save Mulan Lite Setting";
            if(saveDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = saveDlg.FileName;
                write_setting(file_name);
            }
        }

        private void read_setting(string file_name)
        {
            FileStream fs = File.OpenRead(file_name);
            BinaryReader sr = new BinaryReader(fs);

            for(int i = 0; i < FileTable.Length; i++)
            {
                string type_name = FileTable[i].GetType().ToString();
                decimal value;
                switch(type_name)
                {
                    case "System.Windows.Forms.NumericUpDown":
                        value = (sr.ReadByte() << 8)| sr.ReadByte();
                        ((NumericUpDown)FileTable[i]).Value = value;
                        break;
                    case "System.Windows.Forms.ComboBox":
                        value = sr.ReadByte();
                        ((ComboBox)FileTable[i]).SelectedIndex = (int)value;
                        break;
                    case "System.Windows.Forms.CheckBox":
                        value = sr.ReadByte();
                        ((CheckBox)FileTable[i]).Checked = (value == 1 ? true : false);
                        break;
                    case "Sunny.UI.UICheckBox":
                        value = sr.ReadByte();
                        ((UICheckBox)FileTable[i]).Checked = (value == 1 ? true : false);
                        break;
                    case "Sunny.UI.UISwitch":
                        value = sr.ReadByte();
                        ((UISwitch)FileTable[i]).Active = (value == 1 ? true : false);
                        break;
                }
            }

            for(int i = 0; i < ReadTable.Length; i++)
            {
                WriteTable[i].Value = sr.ReadByte();
            }

            sr.Close();
            sr.Dispose();
            fs.Close();
            fs.Dispose();
        }


        private void write_setting(string file_name)
        {
            FileStream fs = File.Create(file_name);
            BinaryWriter sw = new BinaryWriter(fs);
            List<byte> gui_setting = new List<byte>();
            for(int i = 0; i < FileTable.Length; i++)
            {
                string type_name = FileTable[i].GetType().ToString();
                byte tmp = 0;
                switch(type_name)
                {
                    case "System.Windows.Forms.NumericUpDown":
                        int value = (int)(((NumericUpDown)FileTable[i]).Value);
                        byte MSB = (byte)((value & 0xFF00) >> 8);
                        byte LSB = (byte)(value & 0xFF);
                        gui_setting.Add(MSB);
                        gui_setting.Add(LSB);
                        break;
                    case "System.Windows.Forms.ComboBox":
                        tmp = (byte)(((ComboBox)FileTable[i]).SelectedIndex);
                        gui_setting.Add(tmp);
                        break;
                    case "System.Windows.Forms.CheckBox":
                        tmp = (byte)((((CheckBox)FileTable[i]).Checked) ? 1 : 0);
                        gui_setting.Add(tmp);
                        break;
                    case "Sunny.UI.UICheckBox":
                        tmp = (byte)((((UICheckBox)FileTable[i]).Checked) ? 1 : 0);
                        gui_setting.Add(tmp);
                        break;
                    case "Sunny.UI.UISwitch":
                        tmp = (byte)((((UISwitch)FileTable[i]).Active) ? 0 : 1);
                        gui_setting.Add(tmp);
                        break;
                }
                
            }
            sw.Write(gui_setting.ToArray());

            for(int i = 0; i < WriteTable.Length; i++)
            {
                sw.Write((byte)WriteTable[i].Value);
            }

            sw.Close();
            sw.Dispose();
            fs.Close();
            fs.Dispose();
        }

        private void uibt_write_Click(object sender, EventArgs e)
        {
            byte[] write_buf = new byte[1] { (byte)nuSWrite.Value };
            
            RTDev.WriteFunc((byte)nuSID.Value, WriteCmd, (byte)nuSAddr.Value, (byte)0x0, write_buf);
        }

        private void uibt_read_Click(object sender, EventArgs e)
        {
            byte[] data = RTDev.ReadFunc((byte)nuSID.Value, 0x00, (byte)nuSAddr.Value);
            nuSRead.Value = data[2];
        }

        private void uibt_flag_setting_Click(object sender, EventArgs e)
        {
            byte id = (byte)nu_conf_id.Value;
            byte addr = (byte)nu_conf_addr.Value;
            byte bit = (byte)nu_conf_bit.Value;

            //byte[] buf = RTDev.ReadFunc((byte)nuSID.Value, 0x00, 0x32);
            //RTDev.WriteFunc((byte)nuSID.Value, WriteCmd, (byte)0x09, (byte)0x00, buf);


            // step 1, config turn on flag
            byte[] data = new byte[] { 0x3D, 0xAE, 0xDD, bit };
            RTDev.SPIWrite(data);
            // step 2, read flag di do switch
            data = new byte[] { 0xAD, 0xAE, 0x04 };
            RTDev.SPIWrite(data);
            // step 3, ram write data enable
            data = new byte[] { 0xAD, 0xAE, 0x01 };
            RTDev.SPIWrite(data);
            // step 4, send read packet
            //data = new byte[] { 0xAC, 0x5A, 0x00, 0xAC, 0x1E, id, (byte)~id, 0x00, 0x00, addr};
            data = new byte[] { 0xAC, 0x1E, id, (byte)~id, 0x00, 0x00, addr };
            RTDev.SPIWrite(data);
            // step 5, start run real time read
            data = new byte[] { 0xAD, 0xAE, 0x02};
            RTDev.SPIWrite(data);

        }

        private void uiButton1_Click(object sender, EventArgs e)
        {
            byte[] data = new byte[] { 0xAD, 0xAE, 0x06 };
            RTDev.SPIWrite(data);
            data = new byte[] { 0xAD, 0xAE, 0x03 };
            RTDev.SPIWrite(data);
        }

        private void cb_ldoio_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            cb.Enabled = false;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)(cb_ldoio.SelectedIndex << 5);
            byte addr = 0x33;
            byte mask = 0xDF;
            WRReg(id, mask, addr, data);
            cb.Enabled = true;
        }

        private void ck_CH0_en_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox ck = (CheckBox)sender;
            int idx = ck.TabIndex;
            byte id = (byte)nu_persentid.Value;
            byte data = (byte)((ck_CH0_en.Checked ? 0x01 : 0x00) | (ck_CH1_en.Checked ? 0x02 : 0x00) | (ck_CH2_en.Checked ? 0x04 : 0x00) | (ck_CH3_en.Checked ? 0x08 : 0x00));
            byte addr = 0x34;
            byte mask = 0xF0;
            WRReg(id, mask, addr, data);
        }
    }
}
