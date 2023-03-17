
namespace SimpleTool
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
            this.bt_into_testmode = new System.Windows.Forms.Button();
            this.nuSlave = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.nuAddr = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.numericUpDown3 = new System.Windows.Forms.NumericUpDown();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.nuSlave)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.nuAddr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown3)).BeginInit();
            this.SuspendLayout();
            // 
            // bt_into_testmode
            // 
            this.bt_into_testmode.Location = new System.Drawing.Point(206, 13);
            this.bt_into_testmode.Name = "bt_into_testmode";
            this.bt_into_testmode.Size = new System.Drawing.Size(148, 23);
            this.bt_into_testmode.TabIndex = 0;
            this.bt_into_testmode.Text = "Into Test mode";
            this.bt_into_testmode.UseVisualStyleBackColor = true;
            this.bt_into_testmode.Click += new System.EventHandler(this.bt_into_testmode_Click);
            // 
            // nuSlave
            // 
            this.nuSlave.Hexadecimal = true;
            this.nuSlave.Location = new System.Drawing.Point(109, 13);
            this.nuSlave.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nuSlave.Name = "nuSlave";
            this.nuSlave.Size = new System.Drawing.Size(79, 22);
            this.nuSlave.TabIndex = 1;
            this.nuSlave.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nuSlave.Value = new decimal(new int[] {
            70,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Slave ID (Hex)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Address (Hex)";
            // 
            // nuAddr
            // 
            this.nuAddr.Hexadecimal = true;
            this.nuAddr.Location = new System.Drawing.Point(109, 41);
            this.nuAddr.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.nuAddr.Name = "nuAddr";
            this.nuAddr.Size = new System.Drawing.Size(79, 22);
            this.nuAddr.TabIndex = 3;
            this.nuAddr.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.nuAddr.Value = new decimal(new int[] {
            34,
            0,
            0,
            0});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 73);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "Data (Hex)";
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.Hexadecimal = true;
            this.numericUpDown2.Location = new System.Drawing.Point(109, 69);
            this.numericUpDown2.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(79, 22);
            this.numericUpDown2.TabIndex = 5;
            this.numericUpDown2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(206, 68);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(53, 23);
            this.button1.TabIndex = 7;
            this.button1.Text = "+";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(265, 68);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(53, 23);
            this.button2.TabIndex = 8;
            this.button2.Text = "-";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 101);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "RData (Hex)";
            // 
            // numericUpDown3
            // 
            this.numericUpDown3.Hexadecimal = true;
            this.numericUpDown3.Location = new System.Drawing.Point(109, 97);
            this.numericUpDown3.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.numericUpDown3.Name = "numericUpDown3";
            this.numericUpDown3.Size = new System.Drawing.Size(79, 22);
            this.numericUpDown3.TabIndex = 9;
            this.numericUpDown3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(206, 96);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(112, 23);
            this.button3.TabIndex = 11;
            this.button3.Text = "Read Data";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 140);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.numericUpDown3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.numericUpDown2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.nuAddr);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.nuSlave);
            this.Controls.Add(this.bt_into_testmode);
            this.Font = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "main";
            this.Text = "SimpleTool v1.0";
            ((System.ComponentModel.ISupportInitialize)(this.nuSlave)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.nuAddr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button bt_into_testmode;
        private System.Windows.Forms.NumericUpDown nuSlave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown nuAddr;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numericUpDown3;
        private System.Windows.Forms.Button button3;
    }
}

