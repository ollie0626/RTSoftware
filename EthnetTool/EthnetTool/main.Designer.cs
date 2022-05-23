
namespace EthnetTool
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
            this.CK_Server = new System.Windows.Forms.CheckBox();
            this.bt_connect = new System.Windows.Forms.Button();
            this.bt_close = new System.Windows.Forms.Button();
            this.bt_getfile = new System.Windows.Forms.Button();
            this.tb_save = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tb_IPaddr = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tb_port = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.bt_send = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CK_Server
            // 
            this.CK_Server.AutoSize = true;
            this.CK_Server.Location = new System.Drawing.Point(12, 26);
            this.CK_Server.Name = "CK_Server";
            this.CK_Server.Size = new System.Drawing.Size(89, 16);
            this.CK_Server.TabIndex = 0;
            this.CK_Server.Text = "Server Enable";
            this.CK_Server.UseVisualStyleBackColor = true;
            this.CK_Server.CheckedChanged += new System.EventHandler(this.CK_Server_CheckedChanged);
            // 
            // bt_connect
            // 
            this.bt_connect.Location = new System.Drawing.Point(107, 22);
            this.bt_connect.Name = "bt_connect";
            this.bt_connect.Size = new System.Drawing.Size(88, 23);
            this.bt_connect.TabIndex = 1;
            this.bt_connect.Text = "Connect";
            this.bt_connect.UseVisualStyleBackColor = true;
            this.bt_connect.Click += new System.EventHandler(this.bt_connect_Click);
            // 
            // bt_close
            // 
            this.bt_close.Location = new System.Drawing.Point(201, 22);
            this.bt_close.Name = "bt_close";
            this.bt_close.Size = new System.Drawing.Size(141, 23);
            this.bt_close.TabIndex = 4;
            this.bt_close.Text = "Client Disconnect";
            this.bt_close.UseVisualStyleBackColor = true;
            this.bt_close.Click += new System.EventHandler(this.bt_close_Click);
            // 
            // bt_getfile
            // 
            this.bt_getfile.Location = new System.Drawing.Point(348, 22);
            this.bt_getfile.Name = "bt_getfile";
            this.bt_getfile.Size = new System.Drawing.Size(100, 23);
            this.bt_getfile.TabIndex = 6;
            this.bt_getfile.Text = "Get File";
            this.bt_getfile.UseVisualStyleBackColor = true;
            this.bt_getfile.Click += new System.EventHandler(this.bt_getfile_Click);
            // 
            // tb_save
            // 
            this.tb_save.Location = new System.Drawing.Point(107, 107);
            this.tb_save.Name = "tb_save";
            this.tb_save.Size = new System.Drawing.Size(272, 22);
            this.tb_save.TabIndex = 7;
            this.tb_save.Text = "C:\\Users\\westg\\Desktop\\";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 110);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 12);
            this.label1.TabIndex = 8;
            this.label1.Text = "Save path (Local)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 138);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "File Name (Local)";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(107, 135);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(272, 22);
            this.textBox1.TabIndex = 9;
            this.textBox1.Text = "temp";
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "IP Address";
            // 
            // tb_IPaddr
            // 
            this.tb_IPaddr.Location = new System.Drawing.Point(107, 51);
            this.tb_IPaddr.Name = "tb_IPaddr";
            this.tb_IPaddr.Size = new System.Drawing.Size(272, 22);
            this.tb_IPaddr.TabIndex = 11;
            this.tb_IPaddr.Text = "192.168.150.1";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 82);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(24, 12);
            this.label4.TabIndex = 14;
            this.label4.Text = "Port";
            // 
            // tb_port
            // 
            this.tb_port.Location = new System.Drawing.Point(107, 79);
            this.tb_port.Name = "tb_port";
            this.tb_port.Size = new System.Drawing.Size(272, 22);
            this.tb_port.TabIndex = 13;
            this.tb_port.Text = "1234";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 174);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(85, 12);
            this.label5.TabIndex = 16;
            this.label5.Text = "All of IP Address";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(12, 189);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(342, 281);
            this.textBox2.TabIndex = 15;
            // 
            // bt_send
            // 
            this.bt_send.Location = new System.Drawing.Point(107, 160);
            this.bt_send.Name = "bt_send";
            this.bt_send.Size = new System.Drawing.Size(88, 23);
            this.bt_send.TabIndex = 17;
            this.bt_send.Text = "Send File";
            this.bt_send.UseVisualStyleBackColor = true;
            this.bt_send.Click += new System.EventHandler(this.bt_send_Click);
            // 
            // main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(503, 492);
            this.Controls.Add(this.bt_send);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tb_port);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tb_IPaddr);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_save);
            this.Controls.Add(this.bt_getfile);
            this.Controls.Add(this.bt_close);
            this.Controls.Add(this.bt_connect);
            this.Controls.Add(this.CK_Server);
            this.Name = "main";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.main_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox CK_Server;
        private System.Windows.Forms.Button bt_connect;
        private System.Windows.Forms.Button bt_close;
        private System.Windows.Forms.Button bt_getfile;
        private System.Windows.Forms.TextBox tb_save;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tb_IPaddr;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tb_port;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button bt_send;
    }
}

