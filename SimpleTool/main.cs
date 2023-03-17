using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace SimpleTool
{
    public partial class main : Form
    {

        string win_name = "SimpleTool v1.3";
        RTBBControl RTDev = new RTBBControl();

        public main()
        {
            InitializeComponent();
            RTDev.BoadInit();
            this.Text = win_name;
        }

        private void bt_into_testmode_Click(object sender, EventArgs e)
        {
            byte[] test_mode_code = new byte[] { 0x5A, 0xA5, 0x62, 0x86, 0x68, 0x26, 0x5A, 0xA5 };
            byte slave = (byte)nuSlave.Value;
            byte addr = (byte)nuAddr.Value;
            RTDev.I2C_Write((byte)(slave >> 1), addr, test_mode_code);
        }



        private void IntoTestMode()
        {
            byte[] test_mode_code = new byte[] { 0x5A, 0xA5, 0x62, 0x86, 0x68, 0x26, 0x5A, 0xA5 };
            byte slave = (byte)nuSlave.Value;
            RTDev.I2C_Write((byte)(slave >> 1), 0xF0, test_mode_code);
        }

        private void ExitTestMode()
        {
            byte[] test_mode_code = new byte[] { 0x08 };
            byte slave = (byte)nuSlave.Value;
            RTDev.I2C_Write((byte)(0xD0 >> 1), 0xDA, test_mode_code);
        }


        private void button1_Click(object sender, EventArgs e)
        {
            IntoTestMode();
            if (numericUpDown2.Value < 255) numericUpDown2.Value = numericUpDown2.Value + 1;
            byte slave = (byte)nuSlave.Value;
            byte addr = (byte)nuAddr.Value;
            byte data = (byte)numericUpDown2.Value;
            RTDev.I2C_Write(0xD0 >> 1, addr, new byte[] { data });
            ExitTestMode();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IntoTestMode();
            if (numericUpDown2.Value > 0) numericUpDown2.Value = numericUpDown2.Value - 1;
            byte slave = (byte)nuSlave.Value;
            byte addr = (byte)nuAddr.Value;
            byte data = (byte)numericUpDown2.Value;
            RTDev.I2C_Write(0xD0 >> 1, addr, new byte[] { data });
            ExitTestMode();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            IntoTestMode();
            byte slave = (byte)nuSlave.Value;
            byte addr = (byte)nuAddr.Value;
            byte[] buf = new byte[1];
            RTDev.I2C_Read(0xD0 >> 1, addr, buf);
            numericUpDown3.Value = buf[0];
            ExitTestMode();

        }
    }
}
