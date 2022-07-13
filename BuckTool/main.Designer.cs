
namespace BuckTool
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
            this.led_chamber = new Sunny.UI.UILedBulb();
            this.led_37940 = new Sunny.UI.UILedBulb();
            this.led_eload = new Sunny.UI.UILedBulb();
            this.led_power = new Sunny.UI.UILedBulb();
            this.led_osc = new Sunny.UI.UILedBulb();
            this.nu_chamber = new System.Windows.Forms.NumericUpDown();
            this.nu_34970A = new System.Windows.Forms.NumericUpDown();
            this.tb_osc = new System.Windows.Forms.TextBox();
            this.nu_power = new System.Windows.Forms.NumericUpDown();
            this.nu_eload = new System.Windows.Forms.NumericUpDown();
            this.uibt_chamber = new Sunny.UI.UISymbolButton();
            this.uibt_34970 = new Sunny.UI.UISymbolButton();
            this.ui_eload_connect = new Sunny.UI.UISymbolButton();
            this.uibt_power_connect = new Sunny.UI.UISymbolButton();
            this.uibt_osc_connect = new Sunny.UI.UISymbolButton();
            this.uibt_kill = new Sunny.UI.UISymbolButton();
            this.uibt_pause = new Sunny.UI.UISymbolButton();
            this.uiSymbolButton1 = new Sunny.UI.UISymbolButton();
            this.uibt_run = new Sunny.UI.UISymbolButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.bt_load_sub = new Sunny.UI.UISymbolButton();
            this.bt_load_add = new Sunny.UI.UISymbolButton();
            this.Eload_DG = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.cb_item = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_Vin = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ck_freq2 = new System.Windows.Forms.CheckBox();
            this.ck_freq1 = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.tb_chamber = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.uiProcessBar1 = new Sunny.UI.UIProcessBar();
            this.label15 = new System.Windows.Forms.Label();
            this.ck_chaber_en = new System.Windows.Forms.CheckBox();
            this.nu_steady = new System.Windows.Forms.NumericUpDown();
            this.label14 = new System.Windows.Forms.Label();
            this.tb_templist = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.ck_slave = new System.Windows.Forms.CheckBox();
            this.ck_multi_chamber = new System.Windows.Forms.CheckBox();
            this.uibt_specify = new Sunny.UI.UISymbolButton();
            this.uibt_Wavepath = new Sunny.UI.UISymbolButton();
            this.uibut_binfile = new Sunny.UI.UISymbolButton();
            this.tbWave = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.nu_chamber)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_34970A)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_power)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_eload)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Eload_DG)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).BeginInit();
            this.SuspendLayout();
            // 
            // led_chamber
            // 
            this.led_chamber.Location = new System.Drawing.Point(280, 162);
            this.led_chamber.Name = "led_chamber";
            this.led_chamber.Size = new System.Drawing.Size(16, 14);
            this.led_chamber.TabIndex = 43;
            this.led_chamber.Text = "uiLedBulb4";
            this.led_chamber.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_37940
            // 
            this.led_37940.Location = new System.Drawing.Point(280, 134);
            this.led_37940.Name = "led_37940";
            this.led_37940.Size = new System.Drawing.Size(16, 14);
            this.led_37940.TabIndex = 42;
            this.led_37940.Text = "uiLedBulb4";
            this.led_37940.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_eload
            // 
            this.led_eload.Location = new System.Drawing.Point(280, 105);
            this.led_eload.Name = "led_eload";
            this.led_eload.Size = new System.Drawing.Size(16, 14);
            this.led_eload.TabIndex = 41;
            this.led_eload.Text = "uiLedBulb3";
            this.led_eload.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_power
            // 
            this.led_power.Location = new System.Drawing.Point(280, 79);
            this.led_power.Name = "led_power";
            this.led_power.Size = new System.Drawing.Size(16, 14);
            this.led_power.TabIndex = 40;
            this.led_power.Text = "uiLedBulb2";
            this.led_power.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_osc
            // 
            this.led_osc.Location = new System.Drawing.Point(280, 51);
            this.led_osc.Name = "led_osc";
            this.led_osc.Size = new System.Drawing.Size(16, 14);
            this.led_osc.TabIndex = 39;
            this.led_osc.Text = "uiLedBulb1";
            this.led_osc.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // nu_chamber
            // 
            this.nu_chamber.Location = new System.Drawing.Point(302, 157);
            this.nu_chamber.Name = "nu_chamber";
            this.nu_chamber.Size = new System.Drawing.Size(62, 23);
            this.nu_chamber.TabIndex = 48;
            this.nu_chamber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nu_chamber.Value = new decimal(new int[] {
            3,
            0,
            0,
            0});
            // 
            // nu_34970A
            // 
            this.nu_34970A.Location = new System.Drawing.Point(302, 129);
            this.nu_34970A.Name = "nu_34970A";
            this.nu_34970A.Size = new System.Drawing.Size(62, 23);
            this.nu_34970A.TabIndex = 47;
            this.nu_34970A.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nu_34970A.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
            // 
            // tb_osc
            // 
            this.tb_osc.Location = new System.Drawing.Point(302, 42);
            this.tb_osc.Name = "tb_osc";
            this.tb_osc.Size = new System.Drawing.Size(222, 23);
            this.tb_osc.TabIndex = 46;
            this.tb_osc.Text = "TCPIP0::168.254.95.0::hislip0::INSTR";
            // 
            // nu_power
            // 
            this.nu_power.Location = new System.Drawing.Point(302, 71);
            this.nu_power.Name = "nu_power";
            this.nu_power.Size = new System.Drawing.Size(62, 23);
            this.nu_power.TabIndex = 45;
            this.nu_power.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nu_power.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // nu_eload
            // 
            this.nu_eload.Location = new System.Drawing.Point(302, 100);
            this.nu_eload.Name = "nu_eload";
            this.nu_eload.Size = new System.Drawing.Size(62, 23);
            this.nu_eload.TabIndex = 44;
            this.nu_eload.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nu_eload.Value = new decimal(new int[] {
            7,
            0,
            0,
            0});
            // 
            // uibt_chamber
            // 
            this.uibt_chamber.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_chamber.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_chamber.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_chamber.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_chamber.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_chamber.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_chamber.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_chamber.Location = new System.Drawing.Point(370, 155);
            this.uibt_chamber.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_chamber.Name = "uibt_chamber";
            this.uibt_chamber.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_chamber.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_chamber.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_chamber.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_chamber.Size = new System.Drawing.Size(143, 22);
            this.uibt_chamber.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_chamber.StyleCustomMode = true;
            this.uibt_chamber.Symbol = 61633;
            this.uibt_chamber.TabIndex = 4;
            this.uibt_chamber.Text = "Chamber Connect";
            this.uibt_chamber.Visible = false;
            this.uibt_chamber.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_chamber.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // uibt_34970
            // 
            this.uibt_34970.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_34970.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_34970.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_34970.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_34970.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_34970.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_34970.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_34970.Location = new System.Drawing.Point(370, 127);
            this.uibt_34970.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_34970.Name = "uibt_34970";
            this.uibt_34970.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_34970.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_34970.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_34970.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_34970.Size = new System.Drawing.Size(143, 22);
            this.uibt_34970.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_34970.StyleCustomMode = true;
            this.uibt_34970.Symbol = 61633;
            this.uibt_34970.TabIndex = 3;
            this.uibt_34970.Text = "34970 Connect";
            this.uibt_34970.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_34970.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // ui_eload_connect
            // 
            this.ui_eload_connect.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ui_eload_connect.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.ui_eload_connect.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.ui_eload_connect.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.ui_eload_connect.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.ui_eload_connect.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.ui_eload_connect.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ui_eload_connect.Location = new System.Drawing.Point(370, 98);
            this.ui_eload_connect.MinimumSize = new System.Drawing.Size(1, 1);
            this.ui_eload_connect.Name = "ui_eload_connect";
            this.ui_eload_connect.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.ui_eload_connect.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.ui_eload_connect.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.ui_eload_connect.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.ui_eload_connect.Size = new System.Drawing.Size(143, 22);
            this.ui_eload_connect.Style = Sunny.UI.UIStyle.Gray;
            this.ui_eload_connect.StyleCustomMode = true;
            this.ui_eload_connect.Symbol = 61633;
            this.ui_eload_connect.TabIndex = 2;
            this.ui_eload_connect.Text = "ELoad Connect";
            this.ui_eload_connect.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.ui_eload_connect.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // uibt_power_connect
            // 
            this.uibt_power_connect.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_power_connect.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_power_connect.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_power_connect.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_power_connect.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_power_connect.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_power_connect.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_power_connect.Location = new System.Drawing.Point(370, 72);
            this.uibt_power_connect.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_power_connect.Name = "uibt_power_connect";
            this.uibt_power_connect.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_power_connect.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_power_connect.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_power_connect.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_power_connect.Size = new System.Drawing.Size(143, 22);
            this.uibt_power_connect.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_power_connect.StyleCustomMode = true;
            this.uibt_power_connect.Symbol = 61633;
            this.uibt_power_connect.TabIndex = 1;
            this.uibt_power_connect.Text = "Power Connect";
            this.uibt_power_connect.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_power_connect.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // uibt_osc_connect
            // 
            this.uibt_osc_connect.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_osc_connect.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_osc_connect.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_osc_connect.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_osc_connect.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_osc_connect.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_osc_connect.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_osc_connect.Location = new System.Drawing.Point(530, 43);
            this.uibt_osc_connect.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_osc_connect.Name = "uibt_osc_connect";
            this.uibt_osc_connect.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_osc_connect.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_osc_connect.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_osc_connect.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_osc_connect.Size = new System.Drawing.Size(135, 22);
            this.uibt_osc_connect.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_osc_connect.StyleCustomMode = true;
            this.uibt_osc_connect.Symbol = 61633;
            this.uibt_osc_connect.TabIndex = 0;
            this.uibt_osc_connect.Text = "OSC Connect";
            this.uibt_osc_connect.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_osc_connect.Click += new System.EventHandler(this.uibt_osc_connect_Click);
            // 
            // uibt_kill
            // 
            this.uibt_kill.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_kill.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_kill.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_kill.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_kill.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_kill.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_kill.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_kill.Location = new System.Drawing.Point(530, 155);
            this.uibt_kill.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_kill.Name = "uibt_kill";
            this.uibt_kill.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_kill.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_kill.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_kill.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_kill.Size = new System.Drawing.Size(135, 21);
            this.uibt_kill.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_kill.StyleCustomMode = true;
            this.uibt_kill.Symbol = 61944;
            this.uibt_kill.SymbolSize = 18;
            this.uibt_kill.TabIndex = 53;
            this.uibt_kill.Text = "Kill Excel";
            this.uibt_kill.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_kill.Click += new System.EventHandler(this.uibt_kill_Click);
            // 
            // uibt_pause
            // 
            this.uibt_pause.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_pause.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_pause.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_pause.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(158)))), ((int)(((byte)(160)))), ((int)(((byte)(165)))));
            this.uibt_pause.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(123)))), ((int)(((byte)(129)))));
            this.uibt_pause.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(123)))), ((int)(((byte)(129)))));
            this.uibt_pause.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_pause.Location = new System.Drawing.Point(530, 100);
            this.uibt_pause.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_pause.Name = "uibt_pause";
            this.uibt_pause.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_pause.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(158)))), ((int)(((byte)(160)))), ((int)(((byte)(165)))));
            this.uibt_pause.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(123)))), ((int)(((byte)(129)))));
            this.uibt_pause.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(121)))), ((int)(((byte)(123)))), ((int)(((byte)(129)))));
            this.uibt_pause.Size = new System.Drawing.Size(135, 21);
            this.uibt_pause.Style = Sunny.UI.UIStyle.Custom;
            this.uibt_pause.StyleCustomMode = true;
            this.uibt_pause.Symbol = 61516;
            this.uibt_pause.SymbolSize = 18;
            this.uibt_pause.TabIndex = 63;
            this.uibt_pause.Text = "Pause";
            this.uibt_pause.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_pause.Click += new System.EventHandler(this.uibt_pause_Click);
            // 
            // uiSymbolButton1
            // 
            this.uiSymbolButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uiSymbolButton1.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uiSymbolButton1.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uiSymbolButton1.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uiSymbolButton1.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uiSymbolButton1.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uiSymbolButton1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uiSymbolButton1.Location = new System.Drawing.Point(530, 127);
            this.uiSymbolButton1.MinimumSize = new System.Drawing.Size(1, 1);
            this.uiSymbolButton1.Name = "uiSymbolButton1";
            this.uiSymbolButton1.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uiSymbolButton1.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uiSymbolButton1.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uiSymbolButton1.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uiSymbolButton1.Size = new System.Drawing.Size(135, 22);
            this.uiSymbolButton1.Style = Sunny.UI.UIStyle.Gray;
            this.uiSymbolButton1.StyleCustomMode = true;
            this.uiSymbolButton1.Symbol = 61517;
            this.uiSymbolButton1.SymbolSize = 18;
            this.uiSymbolButton1.TabIndex = 62;
            this.uiSymbolButton1.Text = "STOP";
            this.uiSymbolButton1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uiSymbolButton1.Click += new System.EventHandler(this.uiSymbolButton1_Click);
            // 
            // uibt_run
            // 
            this.uibt_run.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_run.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_run.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_run.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_run.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_run.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_run.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_run.Location = new System.Drawing.Point(530, 72);
            this.uibt_run.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_run.Name = "uibt_run";
            this.uibt_run.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_run.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_run.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_run.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_run.Size = new System.Drawing.Size(135, 22);
            this.uibt_run.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_run.StyleCustomMode = true;
            this.uibt_run.Symbol = 61515;
            this.uibt_run.SymbolSize = 18;
            this.uibt_run.TabIndex = 61;
            this.uibt_run.Text = "RUN";
            this.uibt_run.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_run.Click += new System.EventHandler(this.uibt_run_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.bt_load_sub);
            this.groupBox1.Controls.Add(this.bt_load_add);
            this.groupBox1.Controls.Add(this.Eload_DG);
            this.groupBox1.Location = new System.Drawing.Point(11, 247);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(403, 189);
            this.groupBox1.TabIndex = 64;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ELoad Setting";
            // 
            // bt_load_sub
            // 
            this.bt_load_sub.CircleRectWidth = 10;
            this.bt_load_sub.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bt_load_sub.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_sub.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_sub.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_sub.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.bt_load_sub.ImageInterval = 1;
            this.bt_load_sub.Location = new System.Drawing.Point(6, 57);
            this.bt_load_sub.MinimumSize = new System.Drawing.Size(1, 1);
            this.bt_load_sub.Name = "bt_load_sub";
            this.bt_load_sub.Radius = 18;
            this.bt_load_sub.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_sub.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_sub.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.RectSize = 2;
            this.bt_load_sub.Size = new System.Drawing.Size(29, 29);
            this.bt_load_sub.Style = Sunny.UI.UIStyle.Gray;
            this.bt_load_sub.Symbol = 61544;
            this.bt_load_sub.SymbolSize = 12;
            this.bt_load_sub.TabIndex = 66;
            this.bt_load_sub.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.bt_load_sub.Click += new System.EventHandler(this.bt_load_sub_Click);
            // 
            // bt_load_add
            // 
            this.bt_load_add.CircleRectWidth = 10;
            this.bt_load_add.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bt_load_add.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_add.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_add.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_add.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.bt_load_add.ImageInterval = 1;
            this.bt_load_add.Location = new System.Drawing.Point(6, 22);
            this.bt_load_add.MinimumSize = new System.Drawing.Size(1, 1);
            this.bt_load_add.Name = "bt_load_add";
            this.bt_load_add.Radius = 18;
            this.bt_load_add.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_add.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_add.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.RectSize = 2;
            this.bt_load_add.Size = new System.Drawing.Size(29, 29);
            this.bt_load_add.Style = Sunny.UI.UIStyle.Gray;
            this.bt_load_add.Symbol = 61543;
            this.bt_load_add.SymbolSize = 12;
            this.bt_load_add.TabIndex = 65;
            this.bt_load_add.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.bt_load_add.Click += new System.EventHandler(this.bt_load_add_Click);
            // 
            // Eload_DG
            // 
            this.Eload_DG.AllowUserToAddRows = false;
            this.Eload_DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Eload_DG.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
            this.Eload_DG.Location = new System.Drawing.Point(41, 22);
            this.Eload_DG.Name = "Eload_DG";
            this.Eload_DG.RowTemplate.Height = 24;
            this.Eload_DG.Size = new System.Drawing.Size(347, 149);
            this.Eload_DG.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "start (A)";
            this.Column1.Name = "Column1";
            // 
            // Column2
            // 
            this.Column2.HeaderText = "step (A)";
            this.Column2.Name = "Column2";
            // 
            // Column3
            // 
            this.Column3.HeaderText = "stop (A)";
            this.Column3.Name = "Column3";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 134);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 15);
            this.label1.TabIndex = 65;
            this.label1.Text = "test item select";
            // 
            // cb_item
            // 
            this.cb_item.FormattingEnabled = true;
            this.cb_item.Items.AddRange(new object[] {
            "1. Efficiency/Load Regulation",
            "2. Line Regulation"});
            this.cb_item.Location = new System.Drawing.Point(98, 133);
            this.cb_item.Name = "cb_item";
            this.cb_item.Size = new System.Drawing.Size(172, 23);
            this.cb_item.TabIndex = 66;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 165);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 15);
            this.label2.TabIndex = 67;
            this.label2.Text = "Vin";
            // 
            // tb_Vin
            // 
            this.tb_Vin.Location = new System.Drawing.Point(39, 162);
            this.tb_Vin.Name = "tb_Vin";
            this.tb_Vin.Size = new System.Drawing.Size(231, 23);
            this.tb_Vin.TabIndex = 68;
            this.tb_Vin.Text = "3,3.3,3.5";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ck_freq2);
            this.groupBox2.Controls.Add(this.ck_freq1);
            this.groupBox2.Location = new System.Drawing.Point(11, 191);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(403, 44);
            this.groupBox2.TabIndex = 69;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Freq Select";
            // 
            // ck_freq2
            // 
            this.ck_freq2.AutoSize = true;
            this.ck_freq2.Location = new System.Drawing.Point(128, 19);
            this.ck_freq2.Name = "ck_freq2";
            this.ck_freq2.Size = new System.Drawing.Size(101, 19);
            this.ck_freq2.TabIndex = 1;
            this.ck_freq2.Text = "Freq 1.25MHz";
            this.ck_freq2.UseVisualStyleBackColor = true;
            // 
            // ck_freq1
            // 
            this.ck_freq1.AutoSize = true;
            this.ck_freq1.Checked = true;
            this.ck_freq1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ck_freq1.Location = new System.Drawing.Point(13, 19);
            this.ck_freq1.Name = "ck_freq1";
            this.ck_freq1.Size = new System.Drawing.Size(94, 19);
            this.ck_freq1.TabIndex = 0;
            this.ck_freq1.Text = "Freq 2.5MHz";
            this.ck_freq1.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.tb_chamber);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.uiProcessBar1);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Controls.Add(this.ck_chaber_en);
            this.groupBox3.Controls.Add(this.nu_steady);
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.tb_templist);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.ck_slave);
            this.groupBox3.Controls.Add(this.ck_multi_chamber);
            this.groupBox3.Location = new System.Drawing.Point(420, 191);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(265, 245);
            this.groupBox3.TabIndex = 70;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Chamber Ctrl";
            // 
            // tb_chamber
            // 
            this.tb_chamber.FormattingEnabled = true;
            this.tb_chamber.Location = new System.Drawing.Point(102, 79);
            this.tb_chamber.Name = "tb_chamber";
            this.tb_chamber.Size = new System.Drawing.Size(121, 23);
            this.tb_chamber.TabIndex = 51;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 162);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(101, 15);
            this.label3.TabIndex = 50;
            this.label3.Text = "count down: 5:00";
            // 
            // uiProcessBar1
            // 
            this.uiProcessBar1.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.uiProcessBar1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uiProcessBar1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uiProcessBar1.Location = new System.Drawing.Point(5, 182);
            this.uiProcessBar1.MinimumSize = new System.Drawing.Size(70, 3);
            this.uiProcessBar1.Name = "uiProcessBar1";
            this.uiProcessBar1.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uiProcessBar1.Size = new System.Drawing.Size(245, 27);
            this.uiProcessBar1.Style = Sunny.UI.UIStyle.Gray;
            this.uiProcessBar1.TabIndex = 49;
            this.uiProcessBar1.Text = "uiProcessBar1";
            this.uiProcessBar1.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(95, 131);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(83, 15);
            this.label15.TabIndex = 11;
            this.label15.Text = "Steady time(s)";
            // 
            // ck_chaber_en
            // 
            this.ck_chaber_en.AutoSize = true;
            this.ck_chaber_en.Location = new System.Drawing.Point(9, 22);
            this.ck_chaber_en.Name = "ck_chaber_en";
            this.ck_chaber_en.Size = new System.Drawing.Size(115, 19);
            this.ck_chaber_en.TabIndex = 48;
            this.ck_chaber_en.Text = "Chamber Enable";
            this.ck_chaber_en.UseVisualStyleBackColor = true;
            // 
            // nu_steady
            // 
            this.nu_steady.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nu_steady.Location = new System.Drawing.Point(184, 129);
            this.nu_steady.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.nu_steady.Maximum = new decimal(new int[] {
            6000,
            0,
            0,
            0});
            this.nu_steady.Name = "nu_steady";
            this.nu_steady.Size = new System.Drawing.Size(66, 23);
            this.nu_steady.TabIndex = 10;
            this.nu_steady.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 50);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(87, 15);
            this.label14.TabIndex = 46;
            this.label14.Text = "Chamber Temp";
            // 
            // tb_templist
            // 
            this.tb_templist.Location = new System.Drawing.Point(102, 47);
            this.tb_templist.Name = "tb_templist";
            this.tb_templist.Size = new System.Drawing.Size(152, 23);
            this.tb_templist.TabIndex = 47;
            this.tb_templist.Text = "25,40,80";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 79);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(90, 15);
            this.label13.TabIndex = 42;
            this.label13.Text = "Chamber Name";
            // 
            // ck_slave
            // 
            this.ck_slave.AutoSize = true;
            this.ck_slave.Location = new System.Drawing.Point(9, 130);
            this.ck_slave.Name = "ck_slave";
            this.ck_slave.Size = new System.Drawing.Size(55, 19);
            this.ck_slave.TabIndex = 45;
            this.ck_slave.Text = "Slave";
            this.ck_slave.UseVisualStyleBackColor = true;
            // 
            // ck_multi_chamber
            // 
            this.ck_multi_chamber.AutoSize = true;
            this.ck_multi_chamber.Location = new System.Drawing.Point(8, 105);
            this.ck_multi_chamber.Name = "ck_multi_chamber";
            this.ck_multi_chamber.Size = new System.Drawing.Size(107, 19);
            this.ck_multi_chamber.TabIndex = 44;
            this.ck_multi_chamber.Text = "Multi Chamber";
            this.ck_multi_chamber.UseVisualStyleBackColor = true;
            // 
            // uibt_specify
            // 
            this.uibt_specify.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_specify.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_specify.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_specify.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_specify.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_specify.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_specify.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_specify.Location = new System.Drawing.Point(3, 100);
            this.uibt_specify.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_specify.Name = "uibt_specify";
            this.uibt_specify.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_specify.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_specify.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_specify.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_specify.Size = new System.Drawing.Size(109, 27);
            this.uibt_specify.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_specify.StyleCustomMode = true;
            this.uibt_specify.Symbol = 61462;
            this.uibt_specify.TabIndex = 76;
            this.uibt_specify.Text = "Specify Bin";
            this.uibt_specify.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_specify.Click += new System.EventHandler(this.uibt_specify_Click);
            // 
            // uibt_Wavepath
            // 
            this.uibt_Wavepath.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibt_Wavepath.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_Wavepath.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_Wavepath.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_Wavepath.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_Wavepath.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_Wavepath.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibt_Wavepath.Location = new System.Drawing.Point(3, 71);
            this.uibt_Wavepath.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibt_Wavepath.Name = "uibt_Wavepath";
            this.uibt_Wavepath.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibt_Wavepath.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibt_Wavepath.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_Wavepath.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibt_Wavepath.Size = new System.Drawing.Size(109, 27);
            this.uibt_Wavepath.Style = Sunny.UI.UIStyle.Gray;
            this.uibt_Wavepath.StyleCustomMode = true;
            this.uibt_Wavepath.Symbol = 61717;
            this.uibt_Wavepath.TabIndex = 75;
            this.uibt_Wavepath.Text = "Wave Path";
            this.uibt_Wavepath.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibt_Wavepath.Click += new System.EventHandler(this.uibt_Wavepath_Click);
            // 
            // uibut_binfile
            // 
            this.uibut_binfile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.uibut_binfile.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibut_binfile.FillColor2 = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibut_binfile.FillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibut_binfile.FillPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibut_binfile.FillSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibut_binfile.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uibut_binfile.Location = new System.Drawing.Point(3, 42);
            this.uibut_binfile.MinimumSize = new System.Drawing.Size(1, 1);
            this.uibut_binfile.Name = "uibut_binfile";
            this.uibut_binfile.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.uibut_binfile.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.uibut_binfile.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibut_binfile.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.uibut_binfile.Size = new System.Drawing.Size(109, 27);
            this.uibut_binfile.Style = Sunny.UI.UIStyle.Gray;
            this.uibut_binfile.StyleCustomMode = true;
            this.uibut_binfile.Symbol = 61787;
            this.uibut_binfile.TabIndex = 74;
            this.uibut_binfile.Text = "Bin File";
            this.uibut_binfile.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.uibut_binfile.Click += new System.EventHandler(this.uibut_binfile_Click);
            // 
            // tbWave
            // 
            this.tbWave.Location = new System.Drawing.Point(118, 75);
            this.tbWave.Name = "tbWave";
            this.tbWave.Size = new System.Drawing.Size(152, 23);
            this.tbWave.TabIndex = 73;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(118, 104);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(152, 23);
            this.textBox2.TabIndex = 72;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(118, 43);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(152, 23);
            this.textBox1.TabIndex = 71;
            this.textBox1.Text = "D:\\";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // main
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(699, 455);
            this.ControlBoxFillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.Controls.Add(this.uibt_specify);
            this.Controls.Add(this.uibt_Wavepath);
            this.Controls.Add(this.uibut_binfile);
            this.Controls.Add(this.tbWave);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.tb_Vin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cb_item);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.uibt_pause);
            this.Controls.Add(this.uiSymbolButton1);
            this.Controls.Add(this.uibt_run);
            this.Controls.Add(this.uibt_kill);
            this.Controls.Add(this.uibt_chamber);
            this.Controls.Add(this.uibt_34970);
            this.Controls.Add(this.ui_eload_connect);
            this.Controls.Add(this.uibt_power_connect);
            this.Controls.Add(this.uibt_osc_connect);
            this.Controls.Add(this.nu_chamber);
            this.Controls.Add(this.nu_34970A);
            this.Controls.Add(this.tb_osc);
            this.Controls.Add(this.nu_power);
            this.Controls.Add(this.nu_eload);
            this.Controls.Add(this.led_chamber);
            this.Controls.Add(this.led_37940);
            this.Controls.Add(this.led_eload);
            this.Controls.Add(this.led_power);
            this.Controls.Add(this.led_osc);
            this.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "main";
            this.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.Style = Sunny.UI.UIStyle.Gray;
            this.StyleCustomMode = true;
            this.Text = "Buck Tool v1";
            this.TitleColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.TitleFont = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 800, 450);
            ((System.ComponentModel.ISupportInitialize)(this.nu_chamber)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_34970A)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_power)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_eload)).EndInit();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Eload_DG)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nu_steady)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Sunny.UI.UILedBulb led_chamber;
        private Sunny.UI.UILedBulb led_37940;
        private Sunny.UI.UILedBulb led_eload;
        private Sunny.UI.UILedBulb led_power;
        private Sunny.UI.UILedBulb led_osc;
        private System.Windows.Forms.NumericUpDown nu_chamber;
        private System.Windows.Forms.NumericUpDown nu_34970A;
        private System.Windows.Forms.TextBox tb_osc;
        private System.Windows.Forms.NumericUpDown nu_power;
        private System.Windows.Forms.NumericUpDown nu_eload;
        private Sunny.UI.UISymbolButton uibt_chamber;
        private Sunny.UI.UISymbolButton uibt_34970;
        private Sunny.UI.UISymbolButton ui_eload_connect;
        private Sunny.UI.UISymbolButton uibt_power_connect;
        private Sunny.UI.UISymbolButton uibt_osc_connect;
        private Sunny.UI.UISymbolButton uibt_kill;
        private Sunny.UI.UISymbolButton uibt_pause;
        private Sunny.UI.UISymbolButton uiSymbolButton1;
        private Sunny.UI.UISymbolButton uibt_run;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView Eload_DG;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private Sunny.UI.UISymbolButton bt_load_add;
        private Sunny.UI.UISymbolButton bt_load_sub;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cb_item;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_Vin;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox ck_freq1;
        private System.Windows.Forms.CheckBox ck_freq2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.ComboBox tb_chamber;
        private System.Windows.Forms.Label label3;
        private Sunny.UI.UIProcessBar uiProcessBar1;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.CheckBox ck_chaber_en;
        private System.Windows.Forms.NumericUpDown nu_steady;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox tb_templist;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox ck_slave;
        private System.Windows.Forms.CheckBox ck_multi_chamber;
        private Sunny.UI.UISymbolButton uibt_specify;
        private Sunny.UI.UISymbolButton uibt_Wavepath;
        private Sunny.UI.UISymbolButton uibut_binfile;
        private System.Windows.Forms.TextBox tbWave;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

