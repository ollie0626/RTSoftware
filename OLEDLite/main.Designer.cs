﻿
namespace OLEDLite
{
    partial class main
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(main));
            this.materialTabSelector1 = new MaterialSkin.Controls.MaterialTabSelector();
            this.materialTabControl1 = new MaterialSkin.Controls.MaterialTabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.ck_func = new System.Windows.Forms.CheckBox();
            this.tb_res_func = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.bt_func_set = new System.Windows.Forms.Button();
            this.nu_Tf = new System.Windows.Forms.NumericUpDown();
            this.nu_Tr = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tb_Low_level = new System.Windows.Forms.TextBox();
            this.tb_High_level = new System.Windows.Forms.TextBox();
            this.nu_duty = new System.Windows.Forms.NumericUpDown();
            this.nu_Freq = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.bt_connect = new MaterialSkin.Controls.MaterialButton();
            this.cb_item = new System.Windows.Forms.ComboBox();
            this.bt_scanIns = new MaterialSkin.Controls.MaterialButton();
            this.label17 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.nu_swire_num = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.swireTable = new System.Windows.Forms.DataGridView();
            this.tb_initial_bin = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.tb_wave_path = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.tb_bin = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.nu_slave = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.CK_I2c = new System.Windows.Forms.CheckBox();
            this.bt_stop = new MaterialSkin.Controls.MaterialButton();
            this.list_ins = new System.Windows.Forms.ListBox();
            this.bt_pause = new MaterialSkin.Controls.MaterialButton();
            this.tb_res_scope = new System.Windows.Forms.TextBox();
            this.bt_run = new MaterialSkin.Controls.MaterialButton();
            this.tb_res_daq = new System.Windows.Forms.TextBox();
            this.ck_chamber = new System.Windows.Forms.CheckBox();
            this.tb_res_eload = new System.Windows.Forms.TextBox();
            this.tb_res_chamber = new System.Windows.Forms.TextBox();
            this.tb_res_power = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.tb_Vin = new System.Windows.Forms.TextBox();
            this.ChamGroup = new System.Windows.Forms.GroupBox();
            this.cb_chamber = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.ck_chaber_en = new System.Windows.Forms.CheckBox();
            this.nu_steady = new System.Windows.Forms.NumericUpDown();
            this.label14 = new System.Windows.Forms.Label();
            this.tb_templist = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.ck_slave = new System.Windows.Forms.CheckBox();
            this.ck_multi_chamber = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.nu_load4 = new System.Windows.Forms.NumericUpDown();
            this.nu_load3 = new System.Windows.Forms.NumericUpDown();
            this.nu_load2 = new System.Windows.Forms.NumericUpDown();
            this.nu_load1 = new System.Windows.Forms.NumericUpDown();
            this.ck_ch4_en = new System.Windows.Forms.CheckBox();
            this.ck_ch3_en = new System.Windows.Forms.CheckBox();
            this.ck_ch2_en = new System.Windows.Forms.CheckBox();
            this.ck_ch1_en = new System.Windows.Forms.CheckBox();
            this.ck_Iout_mode = new System.Windows.Forms.CheckBox();
            this.bt_eload_sub = new System.Windows.Forms.Button();
            this.bt_eload_add = new System.Windows.Forms.Button();
            this.tb_Iout = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.Eload_DG = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ck_scope = new System.Windows.Forms.CheckBox();
            this.ck_meter_out = new System.Windows.Forms.CheckBox();
            this.ck_daq = new System.Windows.Forms.CheckBox();
            this.tb_res_meter_out = new System.Windows.Forms.TextBox();
            this.ck_eload = new System.Windows.Forms.CheckBox();
            this.ck_meter_in = new System.Windows.Forms.CheckBox();
            this.ck_power = new System.Windows.Forms.CheckBox();
            this.tb_res_meter_in = new System.Windows.Forms.TextBox();
            this.materialTabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Tf)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Tr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_duty)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Freq)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_swire_num)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.swireTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_slave)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.ChamGroup.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Eload_DG)).BeginInit();
            this.SuspendLayout();
            // 
            // materialTabSelector1
            // 
            this.materialTabSelector1.BaseTabControl = this.materialTabControl1;
            this.materialTabSelector1.CharacterCasing = MaterialSkin.Controls.MaterialTabSelector.CustomCharacterCasing.Normal;
            this.materialTabSelector1.Depth = 0;
            resources.ApplyResources(this.materialTabSelector1, "materialTabSelector1");
            this.materialTabSelector1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialTabSelector1.Name = "materialTabSelector1";
            // 
            // materialTabControl1
            // 
            this.materialTabControl1.Controls.Add(this.tabPage2);
            this.materialTabControl1.Depth = 0;
            resources.ApplyResources(this.materialTabControl1, "materialTabControl1");
            this.materialTabControl1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialTabControl1.Multiline = true;
            this.materialTabControl1.Name = "materialTabControl1";
            this.materialTabControl1.SelectedIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.ck_func);
            this.tabPage2.Controls.Add(this.tb_res_func);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.bt_connect);
            this.tabPage2.Controls.Add(this.cb_item);
            this.tabPage2.Controls.Add(this.bt_scanIns);
            this.tabPage2.Controls.Add(this.label17);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.bt_stop);
            this.tabPage2.Controls.Add(this.list_ins);
            this.tabPage2.Controls.Add(this.bt_pause);
            this.tabPage2.Controls.Add(this.tb_res_scope);
            this.tabPage2.Controls.Add(this.bt_run);
            this.tabPage2.Controls.Add(this.tb_res_daq);
            this.tabPage2.Controls.Add(this.ck_chamber);
            this.tabPage2.Controls.Add(this.tb_res_eload);
            this.tabPage2.Controls.Add(this.tb_res_chamber);
            this.tabPage2.Controls.Add(this.tb_res_power);
            this.tabPage2.Controls.Add(this.groupBox5);
            this.tabPage2.Controls.Add(this.ChamGroup);
            this.tabPage2.Controls.Add(this.groupBox4);
            this.tabPage2.Controls.Add(this.ck_scope);
            this.tabPage2.Controls.Add(this.ck_meter_out);
            this.tabPage2.Controls.Add(this.ck_daq);
            this.tabPage2.Controls.Add(this.tb_res_meter_out);
            this.tabPage2.Controls.Add(this.ck_eload);
            this.tabPage2.Controls.Add(this.ck_meter_in);
            this.tabPage2.Controls.Add(this.ck_power);
            this.tabPage2.Controls.Add(this.tb_res_meter_in);
            resources.ApplyResources(this.tabPage2, "tabPage2");
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // ck_func
            // 
            resources.ApplyResources(this.ck_func, "ck_func");
            this.ck_func.Name = "ck_func";
            this.ck_func.UseVisualStyleBackColor = true;
            // 
            // tb_res_func
            // 
            resources.ApplyResources(this.tb_res_func, "tb_res_func");
            this.tb_res_func.Name = "tb_res_func";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.bt_func_set);
            this.groupBox1.Controls.Add(this.nu_Tf);
            this.groupBox1.Controls.Add(this.nu_Tr);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.tb_Low_level);
            this.groupBox1.Controls.Add(this.tb_High_level);
            this.groupBox1.Controls.Add(this.nu_duty);
            this.groupBox1.Controls.Add(this.nu_Freq);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // bt_func_set
            // 
            resources.ApplyResources(this.bt_func_set, "bt_func_set");
            this.bt_func_set.Name = "bt_func_set";
            this.bt_func_set.UseVisualStyleBackColor = true;
            this.bt_func_set.Click += new System.EventHandler(this.bt_func_set_Click);
            // 
            // nu_Tf
            // 
            resources.ApplyResources(this.nu_Tf, "nu_Tf");
            this.nu_Tf.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.nu_Tf.Name = "nu_Tf";
            this.nu_Tf.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // nu_Tr
            // 
            resources.ApplyResources(this.nu_Tr, "nu_Tr");
            this.nu_Tr.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.nu_Tr.Name = "nu_Tr";
            this.nu_Tr.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // tb_Low_level
            // 
            resources.ApplyResources(this.tb_Low_level, "tb_Low_level");
            this.tb_Low_level.Name = "tb_Low_level";
            // 
            // tb_High_level
            // 
            resources.ApplyResources(this.tb_High_level, "tb_High_level");
            this.tb_High_level.Name = "tb_High_level";
            // 
            // nu_duty
            // 
            this.nu_duty.DecimalPlaces = 2;
            resources.ApplyResources(this.nu_duty, "nu_duty");
            this.nu_duty.Name = "nu_duty";
            this.nu_duty.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            // 
            // nu_Freq
            // 
            this.nu_Freq.DecimalPlaces = 3;
            resources.ApplyResources(this.nu_Freq, "nu_Freq");
            this.nu_Freq.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.nu_Freq.Name = "nu_Freq";
            this.nu_Freq.Value = new decimal(new int[] {
            3,
            0,
            0,
            65536});
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // bt_connect
            // 
            resources.ApplyResources(this.bt_connect, "bt_connect");
            this.bt_connect.CharacterCasing = MaterialSkin.Controls.MaterialButton.CharacterCasingEnum.Normal;
            this.bt_connect.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.bt_connect.Depth = 0;
            this.bt_connect.HighEmphasis = true;
            this.bt_connect.Icon = null;
            this.bt_connect.MouseState = MaterialSkin.MouseState.HOVER;
            this.bt_connect.Name = "bt_connect";
            this.bt_connect.NoAccentTextColor = System.Drawing.Color.Empty;
            this.bt_connect.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.bt_connect.UseAccentColor = false;
            this.bt_connect.UseVisualStyleBackColor = true;
            this.bt_connect.Click += new System.EventHandler(this.bt_connect_Click);
            // 
            // cb_item
            // 
            this.cb_item.FormattingEnabled = true;
            this.cb_item.Items.AddRange(new object[] {
            resources.GetString("cb_item.Items")});
            resources.ApplyResources(this.cb_item, "cb_item");
            this.cb_item.Name = "cb_item";
            this.cb_item.SelectedIndexChanged += new System.EventHandler(this.cb_item_SelectedIndexChanged);
            // 
            // bt_scanIns
            // 
            resources.ApplyResources(this.bt_scanIns, "bt_scanIns");
            this.bt_scanIns.CharacterCasing = MaterialSkin.Controls.MaterialButton.CharacterCasingEnum.Normal;
            this.bt_scanIns.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.bt_scanIns.Depth = 0;
            this.bt_scanIns.HighEmphasis = true;
            this.bt_scanIns.Icon = null;
            this.bt_scanIns.MouseState = MaterialSkin.MouseState.HOVER;
            this.bt_scanIns.Name = "bt_scanIns";
            this.bt_scanIns.NoAccentTextColor = System.Drawing.Color.Empty;
            this.bt_scanIns.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.bt_scanIns.UseAccentColor = false;
            this.bt_scanIns.UseVisualStyleBackColor = true;
            this.bt_scanIns.Click += new System.EventHandler(this.bt_scanIns_Click);
            // 
            // label17
            // 
            resources.ApplyResources(this.label17, "label17");
            this.label17.Name = "label17";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.nu_swire_num);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.swireTable);
            this.groupBox2.Controls.Add(this.tb_initial_bin);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.tb_wave_path);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.tb_bin);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.nu_slave);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.CK_I2c);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // nu_swire_num
            // 
            this.nu_swire_num.Hexadecimal = true;
            resources.ApplyResources(this.nu_swire_num, "nu_swire_num");
            this.nu_swire_num.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nu_swire_num.Name = "nu_swire_num";
            this.nu_swire_num.ValueChanged += new System.EventHandler(this.nu_swire_num_ValueChanged);
            // 
            // label12
            // 
            resources.ApplyResources(this.label12, "label12");
            this.label12.Name = "label12";
            // 
            // swireTable
            // 
            this.swireTable.AllowUserToAddRows = false;
            this.swireTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.swireTable, "swireTable");
            this.swireTable.Name = "swireTable";
            this.swireTable.RowTemplate.Height = 24;
            // 
            // tb_initial_bin
            // 
            resources.ApplyResources(this.tb_initial_bin, "tb_initial_bin");
            this.tb_initial_bin.Name = "tb_initial_bin";
            // 
            // label10
            // 
            resources.ApplyResources(this.label10, "label10");
            this.label10.Name = "label10";
            // 
            // tb_wave_path
            // 
            resources.ApplyResources(this.tb_wave_path, "tb_wave_path");
            this.tb_wave_path.Name = "tb_wave_path";
            // 
            // label9
            // 
            resources.ApplyResources(this.label9, "label9");
            this.label9.Name = "label9";
            // 
            // tb_bin
            // 
            resources.ApplyResources(this.tb_bin, "tb_bin");
            this.tb_bin.Name = "tb_bin";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // nu_slave
            // 
            this.nu_slave.Hexadecimal = true;
            resources.ApplyResources(this.nu_slave, "nu_slave");
            this.nu_slave.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nu_slave.Name = "nu_slave";
            this.nu_slave.Value = new decimal(new int[] {
            70,
            0,
            0,
            0});
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            // 
            // CK_I2c
            // 
            resources.ApplyResources(this.CK_I2c, "CK_I2c");
            this.CK_I2c.Name = "CK_I2c";
            this.CK_I2c.UseVisualStyleBackColor = true;
            // 
            // bt_stop
            // 
            resources.ApplyResources(this.bt_stop, "bt_stop");
            this.bt_stop.CharacterCasing = MaterialSkin.Controls.MaterialButton.CharacterCasingEnum.Normal;
            this.bt_stop.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.bt_stop.Depth = 0;
            this.bt_stop.HighEmphasis = true;
            this.bt_stop.Icon = null;
            this.bt_stop.MouseState = MaterialSkin.MouseState.HOVER;
            this.bt_stop.Name = "bt_stop";
            this.bt_stop.NoAccentTextColor = System.Drawing.Color.Empty;
            this.bt_stop.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.bt_stop.UseAccentColor = false;
            this.bt_stop.UseVisualStyleBackColor = true;
            // 
            // list_ins
            // 
            this.list_ins.FormattingEnabled = true;
            resources.ApplyResources(this.list_ins, "list_ins");
            this.list_ins.Name = "list_ins";
            // 
            // bt_pause
            // 
            resources.ApplyResources(this.bt_pause, "bt_pause");
            this.bt_pause.CharacterCasing = MaterialSkin.Controls.MaterialButton.CharacterCasingEnum.Normal;
            this.bt_pause.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.bt_pause.Depth = 0;
            this.bt_pause.HighEmphasis = true;
            this.bt_pause.Icon = null;
            this.bt_pause.MouseState = MaterialSkin.MouseState.HOVER;
            this.bt_pause.Name = "bt_pause";
            this.bt_pause.NoAccentTextColor = System.Drawing.Color.Empty;
            this.bt_pause.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.bt_pause.UseAccentColor = false;
            this.bt_pause.UseVisualStyleBackColor = true;
            // 
            // tb_res_scope
            // 
            resources.ApplyResources(this.tb_res_scope, "tb_res_scope");
            this.tb_res_scope.Name = "tb_res_scope";
            // 
            // bt_run
            // 
            resources.ApplyResources(this.bt_run, "bt_run");
            this.bt_run.CharacterCasing = MaterialSkin.Controls.MaterialButton.CharacterCasingEnum.Normal;
            this.bt_run.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.bt_run.Depth = 0;
            this.bt_run.HighEmphasis = true;
            this.bt_run.Icon = null;
            this.bt_run.MouseState = MaterialSkin.MouseState.HOVER;
            this.bt_run.Name = "bt_run";
            this.bt_run.NoAccentTextColor = System.Drawing.Color.Empty;
            this.bt_run.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.bt_run.UseAccentColor = false;
            this.bt_run.UseVisualStyleBackColor = true;
            this.bt_run.Click += new System.EventHandler(this.bt_run_Click);
            // 
            // tb_res_daq
            // 
            resources.ApplyResources(this.tb_res_daq, "tb_res_daq");
            this.tb_res_daq.Name = "tb_res_daq";
            // 
            // ck_chamber
            // 
            resources.ApplyResources(this.ck_chamber, "ck_chamber");
            this.ck_chamber.Name = "ck_chamber";
            this.ck_chamber.UseVisualStyleBackColor = true;
            // 
            // tb_res_eload
            // 
            resources.ApplyResources(this.tb_res_eload, "tb_res_eload");
            this.tb_res_eload.Name = "tb_res_eload";
            // 
            // tb_res_chamber
            // 
            resources.ApplyResources(this.tb_res_chamber, "tb_res_chamber");
            this.tb_res_chamber.Name = "tb_res_chamber";
            // 
            // tb_res_power
            // 
            resources.ApplyResources(this.tb_res_power, "tb_res_power");
            this.tb_res_power.Name = "tb_res_power";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label18);
            this.groupBox5.Controls.Add(this.tb_Vin);
            resources.ApplyResources(this.groupBox5, "groupBox5");
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.TabStop = false;
            // 
            // label18
            // 
            resources.ApplyResources(this.label18, "label18");
            this.label18.Name = "label18";
            // 
            // tb_Vin
            // 
            resources.ApplyResources(this.tb_Vin, "tb_Vin");
            this.tb_Vin.Name = "tb_Vin";
            // 
            // ChamGroup
            // 
            this.ChamGroup.Controls.Add(this.cb_chamber);
            this.ChamGroup.Controls.Add(this.label11);
            this.ChamGroup.Controls.Add(this.label15);
            this.ChamGroup.Controls.Add(this.ck_chaber_en);
            this.ChamGroup.Controls.Add(this.nu_steady);
            this.ChamGroup.Controls.Add(this.label14);
            this.ChamGroup.Controls.Add(this.tb_templist);
            this.ChamGroup.Controls.Add(this.label13);
            this.ChamGroup.Controls.Add(this.ck_slave);
            this.ChamGroup.Controls.Add(this.ck_multi_chamber);
            resources.ApplyResources(this.ChamGroup, "ChamGroup");
            this.ChamGroup.Name = "ChamGroup";
            this.ChamGroup.TabStop = false;
            // 
            // cb_chamber
            // 
            this.cb_chamber.FormattingEnabled = true;
            resources.ApplyResources(this.cb_chamber, "cb_chamber");
            this.cb_chamber.Name = "cb_chamber";
            // 
            // label11
            // 
            resources.ApplyResources(this.label11, "label11");
            this.label11.Name = "label11";
            // 
            // label15
            // 
            resources.ApplyResources(this.label15, "label15");
            this.label15.Name = "label15";
            // 
            // ck_chaber_en
            // 
            resources.ApplyResources(this.ck_chaber_en, "ck_chaber_en");
            this.ck_chaber_en.Name = "ck_chaber_en";
            this.ck_chaber_en.UseVisualStyleBackColor = true;
            // 
            // nu_steady
            // 
            resources.ApplyResources(this.nu_steady, "nu_steady");
            this.nu_steady.Maximum = new decimal(new int[] {
            6000,
            0,
            0,
            0});
            this.nu_steady.Name = "nu_steady";
            this.nu_steady.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label14
            // 
            resources.ApplyResources(this.label14, "label14");
            this.label14.Name = "label14";
            // 
            // tb_templist
            // 
            resources.ApplyResources(this.tb_templist, "tb_templist");
            this.tb_templist.Name = "tb_templist";
            // 
            // label13
            // 
            resources.ApplyResources(this.label13, "label13");
            this.label13.Name = "label13";
            // 
            // ck_slave
            // 
            resources.ApplyResources(this.ck_slave, "ck_slave");
            this.ck_slave.Name = "ck_slave";
            this.ck_slave.UseVisualStyleBackColor = true;
            // 
            // ck_multi_chamber
            // 
            resources.ApplyResources(this.ck_multi_chamber, "ck_multi_chamber");
            this.ck_multi_chamber.Name = "ck_multi_chamber";
            this.ck_multi_chamber.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.groupBox6);
            this.groupBox4.Controls.Add(this.ck_Iout_mode);
            this.groupBox4.Controls.Add(this.bt_eload_sub);
            this.groupBox4.Controls.Add(this.bt_eload_add);
            this.groupBox4.Controls.Add(this.tb_Iout);
            this.groupBox4.Controls.Add(this.label16);
            this.groupBox4.Controls.Add(this.Eload_DG);
            resources.ApplyResources(this.groupBox4, "groupBox4");
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.TabStop = false;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.label22);
            this.groupBox6.Controls.Add(this.label21);
            this.groupBox6.Controls.Add(this.label20);
            this.groupBox6.Controls.Add(this.label19);
            this.groupBox6.Controls.Add(this.nu_load4);
            this.groupBox6.Controls.Add(this.nu_load3);
            this.groupBox6.Controls.Add(this.nu_load2);
            this.groupBox6.Controls.Add(this.nu_load1);
            this.groupBox6.Controls.Add(this.ck_ch4_en);
            this.groupBox6.Controls.Add(this.ck_ch3_en);
            this.groupBox6.Controls.Add(this.ck_ch2_en);
            this.groupBox6.Controls.Add(this.ck_ch1_en);
            resources.ApplyResources(this.groupBox6, "groupBox6");
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.TabStop = false;
            // 
            // label22
            // 
            resources.ApplyResources(this.label22, "label22");
            this.label22.Name = "label22";
            // 
            // label21
            // 
            resources.ApplyResources(this.label21, "label21");
            this.label21.Name = "label21";
            // 
            // label20
            // 
            resources.ApplyResources(this.label20, "label20");
            this.label20.Name = "label20";
            // 
            // label19
            // 
            resources.ApplyResources(this.label19, "label19");
            this.label19.Name = "label19";
            // 
            // nu_load4
            // 
            this.nu_load4.DecimalPlaces = 3;
            resources.ApplyResources(this.nu_load4, "nu_load4");
            this.nu_load4.Name = "nu_load4";
            // 
            // nu_load3
            // 
            this.nu_load3.DecimalPlaces = 3;
            resources.ApplyResources(this.nu_load3, "nu_load3");
            this.nu_load3.Name = "nu_load3";
            // 
            // nu_load2
            // 
            this.nu_load2.DecimalPlaces = 3;
            resources.ApplyResources(this.nu_load2, "nu_load2");
            this.nu_load2.Name = "nu_load2";
            // 
            // nu_load1
            // 
            this.nu_load1.DecimalPlaces = 3;
            resources.ApplyResources(this.nu_load1, "nu_load1");
            this.nu_load1.Name = "nu_load1";
            // 
            // ck_ch4_en
            // 
            resources.ApplyResources(this.ck_ch4_en, "ck_ch4_en");
            this.ck_ch4_en.Name = "ck_ch4_en";
            this.ck_ch4_en.UseVisualStyleBackColor = true;
            // 
            // ck_ch3_en
            // 
            resources.ApplyResources(this.ck_ch3_en, "ck_ch3_en");
            this.ck_ch3_en.Name = "ck_ch3_en";
            this.ck_ch3_en.UseVisualStyleBackColor = true;
            // 
            // ck_ch2_en
            // 
            resources.ApplyResources(this.ck_ch2_en, "ck_ch2_en");
            this.ck_ch2_en.Name = "ck_ch2_en";
            this.ck_ch2_en.UseVisualStyleBackColor = true;
            // 
            // ck_ch1_en
            // 
            resources.ApplyResources(this.ck_ch1_en, "ck_ch1_en");
            this.ck_ch1_en.Name = "ck_ch1_en";
            this.ck_ch1_en.UseVisualStyleBackColor = true;
            // 
            // ck_Iout_mode
            // 
            resources.ApplyResources(this.ck_Iout_mode, "ck_Iout_mode");
            this.ck_Iout_mode.Name = "ck_Iout_mode";
            this.ck_Iout_mode.UseVisualStyleBackColor = true;
            // 
            // bt_eload_sub
            // 
            resources.ApplyResources(this.bt_eload_sub, "bt_eload_sub");
            this.bt_eload_sub.Name = "bt_eload_sub";
            this.bt_eload_sub.UseVisualStyleBackColor = true;
            this.bt_eload_sub.Click += new System.EventHandler(this.bt_eload_sub_Click);
            // 
            // bt_eload_add
            // 
            resources.ApplyResources(this.bt_eload_add, "bt_eload_add");
            this.bt_eload_add.Name = "bt_eload_add";
            this.bt_eload_add.UseVisualStyleBackColor = true;
            this.bt_eload_add.Click += new System.EventHandler(this.bt_eload_add_Click);
            // 
            // tb_Iout
            // 
            resources.ApplyResources(this.tb_Iout, "tb_Iout");
            this.tb_Iout.Name = "tb_Iout";
            // 
            // label16
            // 
            resources.ApplyResources(this.label16, "label16");
            this.label16.Name = "label16";
            // 
            // Eload_DG
            // 
            this.Eload_DG.AllowUserToAddRows = false;
            this.Eload_DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Eload_DG.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
            resources.ApplyResources(this.Eload_DG, "Eload_DG");
            this.Eload_DG.Name = "Eload_DG";
            this.Eload_DG.RowHeadersVisible = false;
            this.Eload_DG.RowTemplate.Height = 24;
            // 
            // Column1
            // 
            resources.ApplyResources(this.Column1, "Column1");
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            resources.ApplyResources(this.Column2, "Column2");
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            resources.ApplyResources(this.Column3, "Column3");
            this.Column3.Name = "Column3";
            // 
            // ck_scope
            // 
            resources.ApplyResources(this.ck_scope, "ck_scope");
            this.ck_scope.Name = "ck_scope";
            this.ck_scope.UseVisualStyleBackColor = true;
            // 
            // ck_meter_out
            // 
            resources.ApplyResources(this.ck_meter_out, "ck_meter_out");
            this.ck_meter_out.Name = "ck_meter_out";
            this.ck_meter_out.UseVisualStyleBackColor = true;
            // 
            // ck_daq
            // 
            resources.ApplyResources(this.ck_daq, "ck_daq");
            this.ck_daq.Name = "ck_daq";
            this.ck_daq.UseVisualStyleBackColor = true;
            // 
            // tb_res_meter_out
            // 
            resources.ApplyResources(this.tb_res_meter_out, "tb_res_meter_out");
            this.tb_res_meter_out.Name = "tb_res_meter_out";
            // 
            // ck_eload
            // 
            resources.ApplyResources(this.ck_eload, "ck_eload");
            this.ck_eload.Name = "ck_eload";
            this.ck_eload.UseVisualStyleBackColor = true;
            // 
            // ck_meter_in
            // 
            resources.ApplyResources(this.ck_meter_in, "ck_meter_in");
            this.ck_meter_in.Name = "ck_meter_in";
            this.ck_meter_in.UseVisualStyleBackColor = true;
            // 
            // ck_power
            // 
            resources.ApplyResources(this.ck_power, "ck_power");
            this.ck_power.Name = "ck_power";
            this.ck_power.UseVisualStyleBackColor = true;
            // 
            // tb_res_meter_in
            // 
            resources.ApplyResources(this.tb_res_meter_in, "tb_res_meter_in");
            this.tb_res_meter_in.Name = "tb_res_meter_in";
            // 
            // main
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.materialTabControl1);
            this.Controls.Add(this.materialTabSelector1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Name = "main";
            this.Resize += new System.EventHandler(this.main_Resize);
            this.materialTabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Tf)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Tr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_duty)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_Freq)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_swire_num)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.swireTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_slave)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ChamGroup.ResumeLayout(false);
            this.ChamGroup.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_load1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Eload_DG)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private MaterialSkin.Controls.MaterialTabSelector materialTabSelector1;
        private MaterialSkin.Controls.MaterialTabControl materialTabControl1;
        private System.Windows.Forms.TabPage tabPage2;
        private MaterialSkin.Controls.MaterialButton bt_connect;
        private MaterialSkin.Controls.MaterialButton bt_scanIns;
        private System.Windows.Forms.ListBox list_ins;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown nu_Freq;
        private System.Windows.Forms.NumericUpDown nu_duty;
        private System.Windows.Forms.NumericUpDown nu_Tf;
        private System.Windows.Forms.NumericUpDown nu_Tr;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tb_Low_level;
        private System.Windows.Forms.TextBox tb_High_level;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox CK_I2c;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown nu_slave;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox tb_res_scope;
        private System.Windows.Forms.TextBox tb_res_daq;
        private System.Windows.Forms.TextBox tb_res_eload;
        private System.Windows.Forms.TextBox tb_res_power;
        private System.Windows.Forms.TextBox tb_bin;
        private System.Windows.Forms.TextBox tb_wave_path;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox tb_initial_bin;
        private System.Windows.Forms.GroupBox ChamGroup;
        private System.Windows.Forms.ComboBox cb_chamber;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.CheckBox ck_chaber_en;
        private System.Windows.Forms.NumericUpDown nu_steady;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox tb_templist;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox ck_slave;
        private System.Windows.Forms.CheckBox ck_multi_chamber;
        private System.Windows.Forms.CheckBox ck_scope;
        private System.Windows.Forms.CheckBox ck_daq;
        private System.Windows.Forms.CheckBox ck_eload;
        private System.Windows.Forms.CheckBox ck_power;
        private System.Windows.Forms.CheckBox ck_meter_in;
        private System.Windows.Forms.TextBox tb_res_meter_in;
        private System.Windows.Forms.CheckBox ck_meter_out;
        private System.Windows.Forms.TextBox tb_res_meter_out;
        private System.Windows.Forms.DataGridView swireTable;
        private System.Windows.Forms.NumericUpDown nu_swire_num;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox tb_Iout;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.DataGridView Eload_DG;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.TextBox tb_Vin;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Button bt_eload_add;
        private System.Windows.Forms.Button bt_eload_sub;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.CheckBox ck_Iout_mode;
        private System.Windows.Forms.CheckBox ck_chamber;
        private System.Windows.Forms.TextBox tb_res_chamber;
        private MaterialSkin.Controls.MaterialButton bt_run;
        private MaterialSkin.Controls.MaterialButton bt_pause;
        private MaterialSkin.Controls.MaterialButton bt_stop;
        private System.Windows.Forms.ComboBox cb_item;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.CheckBox ck_ch1_en;
        private System.Windows.Forms.CheckBox ck_ch4_en;
        private System.Windows.Forms.CheckBox ck_ch3_en;
        private System.Windows.Forms.CheckBox ck_ch2_en;
        private System.Windows.Forms.NumericUpDown nu_load1;
        private System.Windows.Forms.NumericUpDown nu_load4;
        private System.Windows.Forms.NumericUpDown nu_load3;
        private System.Windows.Forms.NumericUpDown nu_load2;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.Button bt_func_set;
        private System.Windows.Forms.CheckBox ck_func;
        private System.Windows.Forms.TextBox tb_res_func;
    }
}
