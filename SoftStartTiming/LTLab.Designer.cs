﻿
namespace SoftStartTiming
{
    partial class LTLab
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.nuslave = new System.Windows.Forms.NumericUpDown();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Addr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Data = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bt_up = new System.Windows.Forms.Button();
            this.bt_down = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.tb_vinList = new System.Windows.Forms.TextBox();
            this.CBChannel = new System.Windows.Forms.ComboBox();
            this.tb_power = new System.Windows.Forms.TextBox();
            this.led_power = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.uibt_osc_connect = new System.Windows.Forms.Button();
            this.tb_osc = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.list_ins = new System.Windows.Forms.ListBox();
            this.led_osc = new System.Windows.Forms.TextBox();
            this.BTScan = new System.Windows.Forms.Button();
            this.CBPower = new System.Windows.Forms.ComboBox();
            this.BTPause = new System.Windows.Forms.Button();
            this.BTStop = new System.Windows.Forms.Button();
            this.BTRun = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.nuTimeScale = new System.Windows.Forms.NumericUpDown();
            this.BT_LoadSetting = new System.Windows.Forms.Button();
            this.BT_SaveSetting = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.num_Iin_detect = new System.Windows.Forms.NumericUpDown();
            this.num_detect_cnt = new System.Windows.Forms.NumericUpDown();
            this.nu_delay_s = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.tbWave = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.nuslave)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nuTimeScale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Iin_detect)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_detect_cnt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_delay_s)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 137;
            this.label1.Text = "Slave ID";
            this.label1.Visible = false;
            // 
            // nuslave
            // 
            this.nuslave.Hexadecimal = true;
            this.nuslave.Location = new System.Drawing.Point(75, 40);
            this.nuslave.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nuslave.Name = "nuslave";
            this.nuslave.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nuslave.Size = new System.Drawing.Size(57, 22);
            this.nuslave.TabIndex = 138;
            this.nuslave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nuslave.Value = new decimal(new int[] {
            74,
            0,
            0,
            0});
            this.nuslave.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Addr,
            this.Data,
            this.Column1});
            this.dataGridView1.Location = new System.Drawing.Point(12, 330);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(329, 141);
            this.dataGridView1.TabIndex = 139;
            // 
            // Addr
            // 
            this.Addr.HeaderText = "Addr";
            this.Addr.Name = "Addr";
            this.Addr.Width = 80;
            // 
            // Data
            // 
            this.Data.HeaderText = "Data";
            this.Data.Name = "Data";
            this.Data.Width = 80;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Vout";
            this.Column1.Name = "Column1";
            // 
            // bt_up
            // 
            this.bt_up.Location = new System.Drawing.Point(348, 330);
            this.bt_up.Name = "bt_up";
            this.bt_up.Size = new System.Drawing.Size(28, 23);
            this.bt_up.TabIndex = 140;
            this.bt_up.Text = "▲";
            this.bt_up.UseVisualStyleBackColor = true;
            this.bt_up.Click += new System.EventHandler(this.bt_up_Click);
            // 
            // bt_down
            // 
            this.bt_down.Location = new System.Drawing.Point(348, 359);
            this.bt_down.Name = "bt_down";
            this.bt_down.Size = new System.Drawing.Size(28, 23);
            this.bt_down.TabIndex = 141;
            this.bt_down.Text = "▼";
            this.bt_down.UseVisualStyleBackColor = true;
            this.bt_down.Click += new System.EventHandler(this.bt_down_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.tb_vinList);
            this.groupBox4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox4.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox4.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.groupBox4.Location = new System.Drawing.Point(12, 272);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox4.Size = new System.Drawing.Size(258, 53);
            this.groupBox4.TabIndex = 142;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Vin Range (V)";
            // 
            // tb_vinList
            // 
            this.tb_vinList.Location = new System.Drawing.Point(6, 21);
            this.tb_vinList.Name = "tb_vinList";
            this.tb_vinList.Size = new System.Drawing.Size(231, 22);
            this.tb_vinList.TabIndex = 49;
            this.tb_vinList.Text = "3.3";
            // 
            // CBChannel
            // 
            this.CBChannel.FormattingEnabled = true;
            this.CBChannel.Items.AddRange(new object[] {
            "E3632",
            "E3633"});
            this.CBChannel.Location = new System.Drawing.Point(161, 243);
            this.CBChannel.Name = "CBChannel";
            this.CBChannel.Size = new System.Drawing.Size(215, 20);
            this.CBChannel.TabIndex = 159;
            // 
            // tb_power
            // 
            this.tb_power.Enabled = false;
            this.tb_power.Location = new System.Drawing.Point(418, 42);
            this.tb_power.Name = "tb_power";
            this.tb_power.Size = new System.Drawing.Size(298, 22);
            this.tb_power.TabIndex = 150;
            this.tb_power.Text = "Power: ";
            // 
            // led_power
            // 
            this.led_power.BackColor = System.Drawing.Color.Red;
            this.led_power.Location = new System.Drawing.Point(386, 42);
            this.led_power.Name = "led_power";
            this.led_power.Size = new System.Drawing.Size(25, 22);
            this.led_power.TabIndex = 146;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(11, 246);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(92, 12);
            this.label10.TabIndex = 158;
            this.label10.Text = "Channel Select:";
            // 
            // uibt_osc_connect
            // 
            this.uibt_osc_connect.Location = new System.Drawing.Point(198, 12);
            this.uibt_osc_connect.Name = "uibt_osc_connect";
            this.uibt_osc_connect.Size = new System.Drawing.Size(178, 23);
            this.uibt_osc_connect.TabIndex = 144;
            this.uibt_osc_connect.Text = "Instrument Connect";
            this.uibt_osc_connect.UseVisualStyleBackColor = true;
            this.uibt_osc_connect.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // tb_osc
            // 
            this.tb_osc.Enabled = false;
            this.tb_osc.Location = new System.Drawing.Point(418, 14);
            this.tb_osc.Name = "tb_osc";
            this.tb_osc.Size = new System.Drawing.Size(298, 22);
            this.tb_osc.TabIndex = 145;
            this.tb_osc.Text = "Scope: ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(11, 217);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(123, 12);
            this.label9.TabIndex = 157;
            this.label9.Text = "Power Supply Select:";
            // 
            // list_ins
            // 
            this.list_ins.FormattingEnabled = true;
            this.list_ins.ItemHeight = 12;
            this.list_ins.Location = new System.Drawing.Point(13, 70);
            this.list_ins.Name = "list_ins";
            this.list_ins.Size = new System.Drawing.Size(363, 136);
            this.list_ins.TabIndex = 156;
            // 
            // led_osc
            // 
            this.led_osc.BackColor = System.Drawing.Color.Red;
            this.led_osc.Location = new System.Drawing.Point(386, 14);
            this.led_osc.Name = "led_osc";
            this.led_osc.Size = new System.Drawing.Size(25, 22);
            this.led_osc.TabIndex = 143;
            // 
            // BTScan
            // 
            this.BTScan.Location = new System.Drawing.Point(12, 12);
            this.BTScan.Name = "BTScan";
            this.BTScan.Size = new System.Drawing.Size(178, 23);
            this.BTScan.TabIndex = 154;
            this.BTScan.Text = "Scan Instrument";
            this.BTScan.UseVisualStyleBackColor = true;
            this.BTScan.Click += new System.EventHandler(this.BTScan_Click);
            // 
            // CBPower
            // 
            this.CBPower.FormattingEnabled = true;
            this.CBPower.Location = new System.Drawing.Point(161, 214);
            this.CBPower.Name = "CBPower";
            this.CBPower.Size = new System.Drawing.Size(215, 20);
            this.CBPower.TabIndex = 155;
            this.CBPower.SelectedIndexChanged += new System.EventHandler(this.CBPower_SelectedIndexChanged);
            // 
            // BTPause
            // 
            this.BTPause.Location = new System.Drawing.Point(467, 70);
            this.BTPause.Name = "BTPause";
            this.BTPause.Size = new System.Drawing.Size(75, 32);
            this.BTPause.TabIndex = 162;
            this.BTPause.Text = "Pause";
            this.BTPause.UseVisualStyleBackColor = true;
            this.BTPause.Click += new System.EventHandler(this.BTPause_Click);
            // 
            // BTStop
            // 
            this.BTStop.Location = new System.Drawing.Point(548, 70);
            this.BTStop.Name = "BTStop";
            this.BTStop.Size = new System.Drawing.Size(75, 32);
            this.BTStop.TabIndex = 161;
            this.BTStop.Text = "Stop";
            this.BTStop.UseVisualStyleBackColor = true;
            this.BTStop.Click += new System.EventHandler(this.BTStop_Click);
            // 
            // BTRun
            // 
            this.BTRun.Location = new System.Drawing.Point(386, 70);
            this.BTRun.Name = "BTRun";
            this.BTRun.Size = new System.Drawing.Size(75, 32);
            this.BTRun.TabIndex = 160;
            this.BTRun.Text = "Run";
            this.BTRun.UseVisualStyleBackColor = true;
            this.BTRun.Click += new System.EventHandler(this.BTRun_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(409, 451);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(133, 12);
            this.label2.TabIndex = 163;
            this.label2.Text = "Initial Time Scale (ms)";
            this.label2.Visible = false;
            // 
            // nuTimeScale
            // 
            this.nuTimeScale.DecimalPlaces = 3;
            this.nuTimeScale.Location = new System.Drawing.Point(548, 449);
            this.nuTimeScale.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.nuTimeScale.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nuTimeScale.Name = "nuTimeScale";
            this.nuTimeScale.Size = new System.Drawing.Size(120, 22);
            this.nuTimeScale.TabIndex = 164;
            this.nuTimeScale.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nuTimeScale.Visible = false;
            // 
            // BT_LoadSetting
            // 
            this.BT_LoadSetting.Location = new System.Drawing.Point(494, 108);
            this.BT_LoadSetting.Name = "BT_LoadSetting";
            this.BT_LoadSetting.Size = new System.Drawing.Size(102, 23);
            this.BT_LoadSetting.TabIndex = 166;
            this.BT_LoadSetting.Text = "Load Setting";
            this.BT_LoadSetting.UseVisualStyleBackColor = true;
            this.BT_LoadSetting.Click += new System.EventHandler(this.BT_LoadSetting_Click);
            // 
            // BT_SaveSetting
            // 
            this.BT_SaveSetting.Location = new System.Drawing.Point(386, 108);
            this.BT_SaveSetting.Name = "BT_SaveSetting";
            this.BT_SaveSetting.Size = new System.Drawing.Size(102, 23);
            this.BT_SaveSetting.TabIndex = 165;
            this.BT_SaveSetting.Text = "Save Setting";
            this.BT_SaveSetting.UseVisualStyleBackColor = true;
            this.BT_SaveSetting.Click += new System.EventHandler(this.BT_SaveSetting_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(384, 134);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 12);
            this.label3.TabIndex = 167;
            this.label3.Text = "Iin(A) Compare";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(384, 173);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 12);
            this.label4.TabIndex = 168;
            this.label4.Text = "Detect count";
            // 
            // num_Iin_detect
            // 
            this.num_Iin_detect.DecimalPlaces = 5;
            this.num_Iin_detect.Increment = new decimal(new int[] {
            1,
            0,
            0,
            327680});
            this.num_Iin_detect.Location = new System.Drawing.Point(386, 149);
            this.num_Iin_detect.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.num_Iin_detect.Name = "num_Iin_detect";
            this.num_Iin_detect.Size = new System.Drawing.Size(120, 22);
            this.num_Iin_detect.TabIndex = 169;
            this.num_Iin_detect.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // num_detect_cnt
            // 
            this.num_detect_cnt.DecimalPlaces = 3;
            this.num_detect_cnt.Location = new System.Drawing.Point(386, 188);
            this.num_detect_cnt.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.num_detect_cnt.Name = "num_detect_cnt";
            this.num_detect_cnt.Size = new System.Drawing.Size(120, 22);
            this.num_detect_cnt.TabIndex = 170;
            this.num_detect_cnt.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // nu_delay_s
            // 
            this.nu_delay_s.DecimalPlaces = 3;
            this.nu_delay_s.Location = new System.Drawing.Point(386, 232);
            this.nu_delay_s.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.nu_delay_s.Name = "nu_delay_s";
            this.nu_delay_s.Size = new System.Drawing.Size(120, 22);
            this.nu_delay_s.TabIndex = 172;
            this.nu_delay_s.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(384, 217);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(84, 12);
            this.label5.TabIndex = 171;
            this.label5.Text = "Delay Time(s)";
            // 
            // tbWave
            // 
            this.tbWave.Location = new System.Drawing.Point(386, 275);
            this.tbWave.Name = "tbWave";
            this.tbWave.Size = new System.Drawing.Size(215, 22);
            this.tbWave.TabIndex = 173;
            this.tbWave.Text = "D:\\";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(384, 260);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(95, 12);
            this.label6.TabIndex = 174;
            this.label6.Text = "Waveform Path:";
            // 
            // LTLab
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(731, 483);
            this.Controls.Add(this.tbWave);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.nu_delay_s);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.num_detect_cnt);
            this.Controls.Add(this.num_Iin_detect);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.BT_LoadSetting);
            this.Controls.Add(this.BT_SaveSetting);
            this.Controls.Add(this.nuTimeScale);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.BTPause);
            this.Controls.Add(this.BTStop);
            this.Controls.Add(this.BTRun);
            this.Controls.Add(this.CBChannel);
            this.Controls.Add(this.tb_power);
            this.Controls.Add(this.led_power);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.uibt_osc_connect);
            this.Controls.Add(this.tb_osc);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.list_ins);
            this.Controls.Add(this.led_osc);
            this.Controls.Add(this.BTScan);
            this.Controls.Add(this.CBPower);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.bt_down);
            this.Controls.Add(this.bt_up);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.nuslave);
            this.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "LTLab";
            this.Text = "LTLab";
            this.Load += new System.EventHandler(this.LTLab_Load);
            ((System.ComponentModel.ISupportInitialize)(this.nuslave)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nuTimeScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_Iin_detect)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.num_detect_cnt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_delay_s)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown nuslave;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button bt_up;
        private System.Windows.Forms.Button bt_down;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox tb_vinList;
        private System.Windows.Forms.ComboBox CBChannel;
        private System.Windows.Forms.TextBox tb_power;
        private System.Windows.Forms.TextBox led_power;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button uibt_osc_connect;
        private System.Windows.Forms.TextBox tb_osc;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ListBox list_ins;
        private System.Windows.Forms.TextBox led_osc;
        private System.Windows.Forms.Button BTScan;
        private System.Windows.Forms.ComboBox CBPower;
        private System.Windows.Forms.Button BTPause;
        private System.Windows.Forms.Button BTStop;
        private System.Windows.Forms.Button BTRun;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nuTimeScale;
        private System.Windows.Forms.DataGridViewTextBoxColumn Addr;
        private System.Windows.Forms.DataGridViewTextBoxColumn Data;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Button BT_LoadSetting;
        private System.Windows.Forms.Button BT_SaveSetting;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown num_Iin_detect;
        private System.Windows.Forms.NumericUpDown num_detect_cnt;
        private System.Windows.Forms.NumericUpDown nu_delay_s;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbWave;
        private System.Windows.Forms.Label label6;
    }
}