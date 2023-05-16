
namespace SoftStartTiming
{
    partial class VIDIO
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
            this.CBChannel = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.BT_LoadSetting = new System.Windows.Forms.Button();
            this.uibt_osc_connect = new System.Windows.Forms.Button();
            this.BT_SaveSetting = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.nuslave = new System.Windows.Forms.NumericUpDown();
            this.tbWave = new System.Windows.Forms.TextBox();
            this.list_ins = new System.Windows.Forms.ListBox();
            this.BTScan = new System.Windows.Forms.Button();
            this.CBPower = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tb_power = new System.Windows.Forms.TextBox();
            this.tb_eload = new System.Windows.Forms.TextBox();
            this.led_power = new System.Windows.Forms.TextBox();
            this.tb_daq = new System.Windows.Forms.TextBox();
            this.tb_osc = new System.Windows.Forms.TextBox();
            this.led_eload = new System.Windows.Forms.TextBox();
            this.tb_chamber = new System.Windows.Forms.TextBox();
            this.led_chamber = new System.Windows.Forms.TextBox();
            this.led_osc = new System.Windows.Forms.TextBox();
            this.led_daq = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.ck_chamber_en = new System.Windows.Forms.CheckBox();
            this.nu_steady = new System.Windows.Forms.NumericUpDown();
            this.label14 = new System.Windows.Forms.Label();
            this.tb_templist = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.nuDisLoad = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.nuDischarge = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridViewComboBoxColumn1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.dataGridViewComboBoxColumn2 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.dataGridViewComboBoxColumn3 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.tb_iout = new System.Windows.Forms.TextBox();
            this.tb_connect3 = new System.Windows.Forms.TextBox();
            this.BT_Sub = new System.Windows.Forms.Button();
            this.tb_connect2 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tb_vinList = new System.Windows.Forms.TextBox();
            this.tb_connect1 = new System.Windows.Forms.TextBox();
            this.BT_Add = new System.Windows.Forms.Button();
            this.BTPause = new System.Windows.Forms.Button();
            this.BTStop = new System.Windows.Forms.Button();
            this.BTRun = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.labStatus = new System.Windows.Forms.Label();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.nuslave)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nuDisLoad)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nuDischarge)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 110;
            this.label1.Text = "Slave ID";
            this.label1.Visible = false;
            // 
            // CBChannel
            // 
            this.CBChannel.FormattingEnabled = true;
            this.CBChannel.Items.AddRange(new object[] {
            "E3632",
            "E3633"});
            this.CBChannel.Location = new System.Drawing.Point(156, 241);
            this.CBChannel.Name = "CBChannel";
            this.CBChannel.Size = new System.Drawing.Size(215, 20);
            this.CBChannel.TabIndex = 119;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(27, 244);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(92, 12);
            this.label10.TabIndex = 118;
            this.label10.Text = "Channel Select:";
            // 
            // BT_LoadSetting
            // 
            this.BT_LoadSetting.Location = new System.Drawing.Point(252, 9);
            this.BT_LoadSetting.Name = "BT_LoadSetting";
            this.BT_LoadSetting.Size = new System.Drawing.Size(106, 23);
            this.BT_LoadSetting.TabIndex = 113;
            this.BT_LoadSetting.Text = "Load Setting";
            this.BT_LoadSetting.UseVisualStyleBackColor = true;
            this.BT_LoadSetting.Click += new System.EventHandler(this.BT_LoadSetting_Click);
            // 
            // uibt_osc_connect
            // 
            this.uibt_osc_connect.Location = new System.Drawing.Point(194, 38);
            this.uibt_osc_connect.Name = "uibt_osc_connect";
            this.uibt_osc_connect.Size = new System.Drawing.Size(178, 23);
            this.uibt_osc_connect.TabIndex = 109;
            this.uibt_osc_connect.Text = "Instrument Connect";
            this.uibt_osc_connect.UseVisualStyleBackColor = true;
            this.uibt_osc_connect.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // BT_SaveSetting
            // 
            this.BT_SaveSetting.Location = new System.Drawing.Point(140, 9);
            this.BT_SaveSetting.Name = "BT_SaveSetting";
            this.BT_SaveSetting.Size = new System.Drawing.Size(106, 23);
            this.BT_SaveSetting.TabIndex = 112;
            this.BT_SaveSetting.Text = "Save Setting";
            this.BT_SaveSetting.UseVisualStyleBackColor = true;
            this.BT_SaveSetting.Click += new System.EventHandler(this.BT_SaveSetting_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(27, 218);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(123, 12);
            this.label9.TabIndex = 117;
            this.label9.Text = "Power Supply Select:";
            // 
            // nuslave
            // 
            this.nuslave.Hexadecimal = true;
            this.nuslave.Location = new System.Drawing.Point(65, 9);
            this.nuslave.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nuslave.Name = "nuslave";
            this.nuslave.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.nuslave.Size = new System.Drawing.Size(64, 22);
            this.nuslave.TabIndex = 111;
            this.nuslave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nuslave.Value = new decimal(new int[] {
            74,
            0,
            0,
            0});
            this.nuslave.Visible = false;
            // 
            // tbWave
            // 
            this.tbWave.Location = new System.Drawing.Point(156, 267);
            this.tbWave.Name = "tbWave";
            this.tbWave.Size = new System.Drawing.Size(215, 22);
            this.tbWave.TabIndex = 120;
            this.tbWave.Text = "D:\\Waveform\\VID";
            // 
            // list_ins
            // 
            this.list_ins.FormattingEnabled = true;
            this.list_ins.ItemHeight = 12;
            this.list_ins.Location = new System.Drawing.Point(8, 71);
            this.list_ins.Name = "list_ins";
            this.list_ins.Size = new System.Drawing.Size(363, 136);
            this.list_ins.TabIndex = 116;
            // 
            // BTScan
            // 
            this.BTScan.Location = new System.Drawing.Point(8, 38);
            this.BTScan.Name = "BTScan";
            this.BTScan.Size = new System.Drawing.Size(178, 23);
            this.BTScan.TabIndex = 114;
            this.BTScan.Text = "Scan Instrument";
            this.BTScan.UseVisualStyleBackColor = true;
            this.BTScan.Click += new System.EventHandler(this.BTScan_Click);
            // 
            // CBPower
            // 
            this.CBPower.FormattingEnabled = true;
            this.CBPower.Location = new System.Drawing.Point(156, 215);
            this.CBPower.Name = "CBPower";
            this.CBPower.Size = new System.Drawing.Size(215, 20);
            this.CBPower.TabIndex = 115;
            this.CBPower.SelectedIndexChanged += new System.EventHandler(this.CBPower_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 270);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 121;
            this.label2.Text = "Waveform Path:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column4,
            this.Column8});
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(480, 73);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(255, 230);
            this.dataGridView1.TabIndex = 122;
            this.dataGridView1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView1_MouseDown);
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Vout (V)";
            this.Column4.Name = "Column4";
            // 
            // Column8
            // 
            this.Column8.HeaderText = "Vout (after)";
            this.Column8.Name = "Column8";
            // 
            // tb_power
            // 
            this.tb_power.Enabled = false;
            this.tb_power.Location = new System.Drawing.Point(410, 68);
            this.tb_power.Name = "tb_power";
            this.tb_power.Size = new System.Drawing.Size(298, 22);
            this.tb_power.TabIndex = 129;
            this.tb_power.Text = "Power: ";
            // 
            // tb_eload
            // 
            this.tb_eload.Enabled = false;
            this.tb_eload.Location = new System.Drawing.Point(410, 96);
            this.tb_eload.Name = "tb_eload";
            this.tb_eload.Size = new System.Drawing.Size(298, 22);
            this.tb_eload.TabIndex = 130;
            this.tb_eload.Text = "ELoad:";
            // 
            // led_power
            // 
            this.led_power.BackColor = System.Drawing.Color.Red;
            this.led_power.Location = new System.Drawing.Point(378, 68);
            this.led_power.Name = "led_power";
            this.led_power.Size = new System.Drawing.Size(25, 22);
            this.led_power.TabIndex = 125;
            // 
            // tb_daq
            // 
            this.tb_daq.Enabled = false;
            this.tb_daq.Location = new System.Drawing.Point(410, 126);
            this.tb_daq.Name = "tb_daq";
            this.tb_daq.Size = new System.Drawing.Size(298, 22);
            this.tb_daq.TabIndex = 131;
            this.tb_daq.Text = "DAQ:";
            // 
            // tb_osc
            // 
            this.tb_osc.Enabled = false;
            this.tb_osc.Location = new System.Drawing.Point(410, 40);
            this.tb_osc.Name = "tb_osc";
            this.tb_osc.Size = new System.Drawing.Size(298, 22);
            this.tb_osc.TabIndex = 124;
            this.tb_osc.Text = "Scope: ";
            // 
            // led_eload
            // 
            this.led_eload.BackColor = System.Drawing.Color.Red;
            this.led_eload.Location = new System.Drawing.Point(378, 96);
            this.led_eload.Name = "led_eload";
            this.led_eload.Size = new System.Drawing.Size(25, 22);
            this.led_eload.TabIndex = 126;
            // 
            // tb_chamber
            // 
            this.tb_chamber.Enabled = false;
            this.tb_chamber.Location = new System.Drawing.Point(410, 154);
            this.tb_chamber.Name = "tb_chamber";
            this.tb_chamber.Size = new System.Drawing.Size(298, 22);
            this.tb_chamber.TabIndex = 132;
            this.tb_chamber.Text = "Chanber:GPIB0::3::INSTR";
            // 
            // led_chamber
            // 
            this.led_chamber.BackColor = System.Drawing.Color.Red;
            this.led_chamber.Location = new System.Drawing.Point(378, 154);
            this.led_chamber.Name = "led_chamber";
            this.led_chamber.Size = new System.Drawing.Size(25, 22);
            this.led_chamber.TabIndex = 128;
            // 
            // led_osc
            // 
            this.led_osc.BackColor = System.Drawing.Color.Red;
            this.led_osc.Location = new System.Drawing.Point(378, 40);
            this.led_osc.Name = "led_osc";
            this.led_osc.Size = new System.Drawing.Size(25, 22);
            this.led_osc.TabIndex = 123;
            // 
            // led_daq
            // 
            this.led_daq.BackColor = System.Drawing.Color.Red;
            this.led_daq.Location = new System.Drawing.Point(378, 126);
            this.led_daq.Name = "led_daq";
            this.led_daq.Size = new System.Drawing.Size(25, 22);
            this.led_daq.TabIndex = 127;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(834, 601);
            this.tabControl1.TabIndex = 133;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.DarkGray;
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.tb_power);
            this.tabPage1.Controls.Add(this.CBPower);
            this.tabPage1.Controls.Add(this.tb_eload);
            this.tabPage1.Controls.Add(this.BTScan);
            this.tabPage1.Controls.Add(this.led_power);
            this.tabPage1.Controls.Add(this.list_ins);
            this.tabPage1.Controls.Add(this.tb_daq);
            this.tabPage1.Controls.Add(this.tbWave);
            this.tabPage1.Controls.Add(this.tb_osc);
            this.tabPage1.Controls.Add(this.nuslave);
            this.tabPage1.Controls.Add(this.led_eload);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.tb_chamber);
            this.tabPage1.Controls.Add(this.BT_SaveSetting);
            this.tabPage1.Controls.Add(this.led_chamber);
            this.tabPage1.Controls.Add(this.uibt_osc_connect);
            this.tabPage1.Controls.Add(this.led_osc);
            this.tabPage1.Controls.Add(this.BT_LoadSetting);
            this.tabPage1.Controls.Add(this.led_daq);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.CBChannel);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(826, 575);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "General Setting";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.progressBar1);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Controls.Add(this.ck_chamber_en);
            this.groupBox3.Controls.Add(this.nu_steady);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.tb_templist);
            this.groupBox3.Location = new System.Drawing.Point(378, 182);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(363, 117);
            this.groupBox3.TabIndex = 133;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Chamber Crtl";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(127, 17);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(126, 23);
            this.progressBar1.TabIndex = 62;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 12);
            this.label3.TabIndex = 61;
            this.label3.Text = "count down: 5:00";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(113, 84);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(85, 12);
            this.label15.TabIndex = 60;
            this.label15.Text = "Steady time(s)";
            // 
            // ck_chamber_en
            // 
            this.ck_chamber_en.AutoSize = true;
            this.ck_chamber_en.Location = new System.Drawing.Point(6, 21);
            this.ck_chamber_en.Name = "ck_chamber_en";
            this.ck_chamber_en.Size = new System.Drawing.Size(116, 16);
            this.ck_chamber_en.TabIndex = 61;
            this.ck_chamber_en.Text = "Chamber Enable";
            this.ck_chamber_en.UseVisualStyleBackColor = true;
            // 
            // nu_steady
            // 
            this.nu_steady.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nu_steady.Location = new System.Drawing.Point(202, 79);
            this.nu_steady.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.nu_steady.Maximum = new decimal(new int[] {
            6000,
            0,
            0,
            0});
            this.nu_steady.Name = "nu_steady";
            this.nu_steady.Size = new System.Drawing.Size(51, 23);
            this.nu_steady.TabIndex = 59;
            this.nu_steady.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 49);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(90, 12);
            this.label14.TabIndex = 59;
            this.label14.Text = "Chamber Temp";
            // 
            // tb_templist
            // 
            this.tb_templist.Location = new System.Drawing.Point(102, 46);
            this.tb_templist.Name = "tb_templist";
            this.tb_templist.Size = new System.Drawing.Size(152, 22);
            this.tb_templist.TabIndex = 60;
            this.tb_templist.Text = "25,40,80";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.DarkGray;
            this.tabPage2.Controls.Add(this.nuDisLoad);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.nuDischarge);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.dataGridView2);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.tb_connect3);
            this.tabPage2.Controls.Add(this.BT_Sub);
            this.tabPage2.Controls.Add(this.tb_connect2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.tb_connect1);
            this.tabPage2.Controls.Add(this.dataGridView1);
            this.tabPage2.Controls.Add(this.BT_Add);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(826, 575);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Test Parameter";
            // 
            // nuDisLoad
            // 
            this.nuDisLoad.DecimalPlaces = 3;
            this.nuDisLoad.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nuDisLoad.Location = new System.Drawing.Point(364, 337);
            this.nuDisLoad.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.nuDisLoad.Maximum = new decimal(new int[] {
            6000,
            0,
            0,
            0});
            this.nuDisLoad.Name = "nuDisLoad";
            this.nuDisLoad.Size = new System.Drawing.Size(77, 23);
            this.nuDisLoad.TabIndex = 139;
            this.nuDisLoad.Value = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(176, 341);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(129, 12);
            this.label5.TabIndex = 140;
            this.label5.Text = "Discharge Loading(A)";
            // 
            // nuDischarge
            // 
            this.nuDischarge.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nuDischarge.Location = new System.Drawing.Point(364, 310);
            this.nuDischarge.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.nuDischarge.Maximum = new decimal(new int[] {
            6000,
            0,
            0,
            0});
            this.nuDischarge.Name = "nuDischarge";
            this.nuDischarge.Size = new System.Drawing.Size(77, 23);
            this.nuDischarge.TabIndex = 137;
            this.nuDischarge.Value = new decimal(new int[] {
            100,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(176, 314);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(182, 12);
            this.label4.TabIndex = 138;
            this.label4.Text = "LPM Mode Discharge Time(ms)";
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewComboBoxColumn1,
            this.dataGridViewComboBoxColumn2,
            this.dataGridViewComboBoxColumn3,
            this.dataGridViewTextBoxColumn1});
            this.dataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView2.Location = new System.Drawing.Point(15, 73);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 24;
            this.dataGridView2.Size = new System.Drawing.Size(455, 230);
            this.dataGridView2.TabIndex = 137;
            this.dataGridView2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridView2_MouseDown);
            // 
            // dataGridViewComboBoxColumn1
            // 
            this.dataGridViewComboBoxColumn1.HeaderText = "LPM";
            this.dataGridViewComboBoxColumn1.Name = "dataGridViewComboBoxColumn1";
            this.dataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // dataGridViewComboBoxColumn2
            // 
            this.dataGridViewComboBoxColumn2.HeaderText = "G1 logic";
            this.dataGridViewComboBoxColumn2.Name = "dataGridViewComboBoxColumn2";
            this.dataGridViewComboBoxColumn2.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // dataGridViewComboBoxColumn3
            // 
            this.dataGridViewComboBoxColumn3.HeaderText = "G2 logic";
            this.dataGridViewComboBoxColumn3.Name = "dataGridViewComboBoxColumn3";
            this.dataGridViewComboBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewComboBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "Vout (V)";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.tb_iout);
            this.groupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox2.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.groupBox2.Location = new System.Drawing.Point(279, 15);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox2.Size = new System.Drawing.Size(258, 53);
            this.groupBox2.TabIndex = 136;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Iout Range (A)";
            // 
            // tb_iout
            // 
            this.tb_iout.Location = new System.Drawing.Point(6, 21);
            this.tb_iout.Name = "tb_iout";
            this.tb_iout.Size = new System.Drawing.Size(231, 22);
            this.tb_iout.TabIndex = 49;
            this.tb_iout.Text = "0.5";
            // 
            // tb_connect3
            // 
            this.tb_connect3.Enabled = false;
            this.tb_connect3.Location = new System.Drawing.Point(15, 367);
            this.tb_connect3.Name = "tb_connect3";
            this.tb_connect3.Size = new System.Drawing.Size(146, 22);
            this.tb_connect3.TabIndex = 136;
            this.tb_connect3.Text = "G2 = GPIO2.2";
            this.tb_connect3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BT_Sub
            // 
            this.BT_Sub.Location = new System.Drawing.Point(519, 309);
            this.BT_Sub.Name = "BT_Sub";
            this.BT_Sub.Size = new System.Drawing.Size(33, 23);
            this.BT_Sub.TabIndex = 134;
            this.BT_Sub.Text = "-";
            this.BT_Sub.UseVisualStyleBackColor = true;
            this.BT_Sub.Click += new System.EventHandler(this.BT_Sub_Click);
            // 
            // tb_connect2
            // 
            this.tb_connect2.Enabled = false;
            this.tb_connect2.Location = new System.Drawing.Point(15, 339);
            this.tb_connect2.Name = "tb_connect2";
            this.tb_connect2.Size = new System.Drawing.Size(146, 22);
            this.tb_connect2.TabIndex = 135;
            this.tb_connect2.Text = "G1 = GPIO2.1";
            this.tb_connect2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tb_vinList);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.groupBox1.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.groupBox1.Location = new System.Drawing.Point(15, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox1.Size = new System.Drawing.Size(258, 53);
            this.groupBox1.TabIndex = 135;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Vin Range (V)";
            // 
            // tb_vinList
            // 
            this.tb_vinList.Location = new System.Drawing.Point(6, 21);
            this.tb_vinList.Name = "tb_vinList";
            this.tb_vinList.Size = new System.Drawing.Size(231, 22);
            this.tb_vinList.TabIndex = 49;
            this.tb_vinList.Text = "3.3";
            // 
            // tb_connect1
            // 
            this.tb_connect1.Enabled = false;
            this.tb_connect1.Location = new System.Drawing.Point(15, 311);
            this.tb_connect1.Name = "tb_connect1";
            this.tb_connect1.Size = new System.Drawing.Size(146, 22);
            this.tb_connect1.TabIndex = 134;
            this.tb_connect1.Text = "LPM = GPIO2.0";
            this.tb_connect1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BT_Add
            // 
            this.BT_Add.Location = new System.Drawing.Point(480, 309);
            this.BT_Add.Name = "BT_Add";
            this.BT_Add.Size = new System.Drawing.Size(33, 23);
            this.BT_Add.TabIndex = 133;
            this.BT_Add.Text = "+";
            this.BT_Add.UseVisualStyleBackColor = true;
            this.BT_Add.Click += new System.EventHandler(this.BT_Add_Click);
            // 
            // BTPause
            // 
            this.BTPause.Location = new System.Drawing.Point(672, 619);
            this.BTPause.Name = "BTPause";
            this.BTPause.Size = new System.Drawing.Size(75, 32);
            this.BTPause.TabIndex = 68;
            this.BTPause.Text = "Pause";
            this.BTPause.UseVisualStyleBackColor = true;
            this.BTPause.Click += new System.EventHandler(this.BTPause_Click);
            // 
            // BTStop
            // 
            this.BTStop.Location = new System.Drawing.Point(753, 619);
            this.BTStop.Name = "BTStop";
            this.BTStop.Size = new System.Drawing.Size(75, 32);
            this.BTStop.TabIndex = 67;
            this.BTStop.Text = "Stop";
            this.BTStop.UseVisualStyleBackColor = true;
            this.BTStop.Click += new System.EventHandler(this.BTStop_Click);
            // 
            // BTRun
            // 
            this.BTRun.Location = new System.Drawing.Point(591, 619);
            this.BTRun.Name = "BTRun";
            this.BTRun.Size = new System.Drawing.Size(75, 32);
            this.BTRun.TabIndex = 66;
            this.BTRun.Text = "Run";
            this.BTRun.UseVisualStyleBackColor = true;
            this.BTRun.Click += new System.EventHandler(this.BTRun_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(496, 619);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(85, 32);
            this.button1.TabIndex = 134;
            this.button1.Text = "Excel Kill";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // labStatus
            // 
            this.labStatus.AutoSize = true;
            this.labStatus.Location = new System.Drawing.Point(22, 619);
            this.labStatus.Name = "labStatus";
            this.labStatus.Size = new System.Drawing.Size(97, 12);
            this.labStatus.TabIndex = 135;
            this.labStatus.Text = "Progress Status :";
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(21, 634);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(371, 21);
            this.progressBar2.TabIndex = 136;
            // 
            // VIDIO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(858, 670);
            this.Controls.Add(this.labStatus);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.BTPause);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.BTStop);
            this.Controls.Add(this.BTRun);
            this.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "VIDIO";
            this.Text = "VIDIO v1.1";
            ((System.ComponentModel.ISupportInitialize)(this.nuslave)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nuDisLoad)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nuDischarge)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox CBChannel;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button BT_LoadSetting;
        private System.Windows.Forms.Button uibt_osc_connect;
        private System.Windows.Forms.Button BT_SaveSetting;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.NumericUpDown nuslave;
        private System.Windows.Forms.TextBox tbWave;
        private System.Windows.Forms.ListBox list_ins;
        private System.Windows.Forms.Button BTScan;
        private System.Windows.Forms.ComboBox CBPower;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox tb_power;
        private System.Windows.Forms.TextBox tb_eload;
        private System.Windows.Forms.TextBox led_power;
        private System.Windows.Forms.TextBox tb_daq;
        private System.Windows.Forms.TextBox tb_osc;
        private System.Windows.Forms.TextBox led_eload;
        private System.Windows.Forms.TextBox tb_chamber;
        private System.Windows.Forms.TextBox led_chamber;
        private System.Windows.Forms.TextBox led_osc;
        private System.Windows.Forms.TextBox led_daq;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button BT_Sub;
        private System.Windows.Forms.Button BT_Add;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tb_vinList;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox tb_iout;
        private System.Windows.Forms.Button BTPause;
        private System.Windows.Forms.Button BTStop;
        private System.Windows.Forms.Button BTRun;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.CheckBox ck_chamber_en;
        private System.Windows.Forms.NumericUpDown nu_steady;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox tb_templist;
        private System.Windows.Forms.TextBox tb_connect3;
        private System.Windows.Forms.TextBox tb_connect2;
        private System.Windows.Forms.TextBox tb_connect1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label labStatus;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column8;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn1;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn2;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewComboBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown nuDischarge;
        private System.Windows.Forms.NumericUpDown nuDisLoad;
        private System.Windows.Forms.Label label5;
    }
}