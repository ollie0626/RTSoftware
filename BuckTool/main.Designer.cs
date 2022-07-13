
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
            this.Eload_DG = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bt_load_add = new Sunny.UI.UISymbolButton();
            this.bt_load_sub = new Sunny.UI.UISymbolButton();
            ((System.ComponentModel.ISupportInitialize)(this.nu_chamber)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_34970A)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_power)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nu_eload)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Eload_DG)).BeginInit();
            this.SuspendLayout();
            // 
            // led_chamber
            // 
            this.led_chamber.Location = new System.Drawing.Point(11, 168);
            this.led_chamber.Name = "led_chamber";
            this.led_chamber.Size = new System.Drawing.Size(16, 14);
            this.led_chamber.TabIndex = 43;
            this.led_chamber.Text = "uiLedBulb4";
            this.led_chamber.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_37940
            // 
            this.led_37940.Location = new System.Drawing.Point(11, 140);
            this.led_37940.Name = "led_37940";
            this.led_37940.Size = new System.Drawing.Size(16, 14);
            this.led_37940.TabIndex = 42;
            this.led_37940.Text = "uiLedBulb4";
            this.led_37940.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_eload
            // 
            this.led_eload.Location = new System.Drawing.Point(11, 111);
            this.led_eload.Name = "led_eload";
            this.led_eload.Size = new System.Drawing.Size(16, 14);
            this.led_eload.TabIndex = 41;
            this.led_eload.Text = "uiLedBulb3";
            this.led_eload.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_power
            // 
            this.led_power.Location = new System.Drawing.Point(11, 85);
            this.led_power.Name = "led_power";
            this.led_power.Size = new System.Drawing.Size(16, 14);
            this.led_power.TabIndex = 40;
            this.led_power.Text = "uiLedBulb2";
            this.led_power.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // led_osc
            // 
            this.led_osc.Location = new System.Drawing.Point(11, 57);
            this.led_osc.Name = "led_osc";
            this.led_osc.Size = new System.Drawing.Size(16, 14);
            this.led_osc.TabIndex = 39;
            this.led_osc.Text = "uiLedBulb1";
            this.led_osc.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // nu_chamber
            // 
            this.nu_chamber.Location = new System.Drawing.Point(33, 163);
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
            this.nu_34970A.Location = new System.Drawing.Point(33, 135);
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
            this.tb_osc.Location = new System.Drawing.Point(33, 48);
            this.tb_osc.Name = "tb_osc";
            this.tb_osc.Size = new System.Drawing.Size(222, 23);
            this.tb_osc.TabIndex = 46;
            this.tb_osc.Text = "TCPIP0::168.254.95.0::hislip0::INSTR";
            // 
            // nu_power
            // 
            this.nu_power.Location = new System.Drawing.Point(33, 77);
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
            this.nu_eload.Location = new System.Drawing.Point(33, 106);
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
            this.uibt_chamber.Location = new System.Drawing.Point(101, 161);
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
            this.uibt_34970.Location = new System.Drawing.Point(101, 133);
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
            this.ui_eload_connect.Location = new System.Drawing.Point(101, 104);
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
            this.uibt_power_connect.Location = new System.Drawing.Point(101, 78);
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
            this.uibt_osc_connect.Location = new System.Drawing.Point(261, 49);
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
            this.uibt_kill.Location = new System.Drawing.Point(261, 161);
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
            this.uibt_pause.Location = new System.Drawing.Point(261, 106);
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
            this.uiSymbolButton1.Location = new System.Drawing.Point(261, 133);
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
            this.uibt_run.Location = new System.Drawing.Point(261, 78);
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
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.bt_load_sub);
            this.groupBox1.Controls.Add(this.bt_load_add);
            this.groupBox1.Controls.Add(this.Eload_DG);
            this.groupBox1.Location = new System.Drawing.Point(23, 226);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(438, 139);
            this.groupBox1.TabIndex = 64;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "ELoad Setting";
            // 
            // Eload_DG
            // 
            this.Eload_DG.AllowUserToAddRows = false;
            this.Eload_DG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Eload_DG.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
            this.Eload_DG.Location = new System.Drawing.Point(66, 22);
            this.Eload_DG.Name = "Eload_DG";
            this.Eload_DG.RowTemplate.Height = 24;
            this.Eload_DG.Size = new System.Drawing.Size(347, 95);
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
            this.bt_load_add.Location = new System.Drawing.Point(10, 22);
            this.bt_load_add.MinimumSize = new System.Drawing.Size(1, 1);
            this.bt_load_add.Name = "bt_load_add";
            this.bt_load_add.Radius = 18;
            this.bt_load_add.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_add.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_add.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_add.RectSize = 2;
            this.bt_load_add.Size = new System.Drawing.Size(41, 39);
            this.bt_load_add.Style = Sunny.UI.UIStyle.Gray;
            this.bt_load_add.Symbol = 61543;
            this.bt_load_add.TabIndex = 65;
            this.bt_load_add.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            this.bt_load_add.Click += new System.EventHandler(this.bt_load_add_Click);
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
            this.bt_load_sub.Location = new System.Drawing.Point(10, 67);
            this.bt_load_sub.MinimumSize = new System.Drawing.Size(1, 1);
            this.bt_load_sub.Name = "bt_load_sub";
            this.bt_load_sub.Radius = 18;
            this.bt_load_sub.RectColor = System.Drawing.Color.FromArgb(((int)(((byte)(140)))), ((int)(((byte)(140)))), ((int)(((byte)(140)))));
            this.bt_load_sub.RectHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
            this.bt_load_sub.RectPressColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.RectSelectedColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(112)))), ((int)(((byte)(112)))));
            this.bt_load_sub.RectSize = 2;
            this.bt_load_sub.Size = new System.Drawing.Size(41, 39);
            this.bt_load_sub.Style = Sunny.UI.UIStyle.Gray;
            this.bt_load_sub.Symbol = 61544;
            this.bt_load_sub.TabIndex = 66;
            this.bt_load_sub.ZoomScaleRect = new System.Drawing.Rectangle(0, 0, 0, 0);
            // 
            // main
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(800, 743);
            this.ControlBoxFillHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(163)))), ((int)(((byte)(163)))), ((int)(((byte)(163)))));
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
    }
}

