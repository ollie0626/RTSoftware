using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using MaterialSkin.Controls;

namespace OLEDLite
{
    public partial class main
    {

        
        private void GUI_Design()
        {
            TDMA_Interface();
        }

        // function gen
        GroupBox func_grop = new GroupBox();
        TextBox tb_hi_level = new TextBox();
        TextBox tb_lo_level = new TextBox();
        NumericUpDown nu_freq = new NumericUpDown();
        NumericUpDown nu_duty = new NumericUpDown();
        NumericUpDown nu_tr = new NumericUpDown();
        NumericUpDown nu_tf = new NumericUpDown();

        private void FuncgenConfig()
        {
            int base_offset = 30;
            Label[] marker_table = new Label[6];
            Control[] table = new Control[] { nu_freq, tb_hi_level, tb_lo_level, nu_duty, nu_tr, nu_tf };

            for (int i = 0; i < marker_table.Length; i++) marker_table[i] = new Label();
            func_grop.Text = "Func Gen Setting";
            func_grop.Size = new System.Drawing.Size(500, 150);
            func_grop.Location = new System.Drawing.Point(10, 10);
            tabPage2.Controls.Add(func_grop);

            marker_table[0].Text = "Freq (KHz)";
            marker_table[1].Text = "HiLevel (V)";
            marker_table[2].Text = "LoLevel (V)";
            marker_table[3].Text = "Duty (%)";
            marker_table[4].Text = "Tr (ns)";
            marker_table[5].Text = "Tf (ns)";

            // label object
            marker_table[0].Location = new System.Drawing.Point(10, base_offset * 1);
            marker_table[1].Location = new System.Drawing.Point(10, base_offset * 2);
            marker_table[2].Location = new System.Drawing.Point(10, base_offset * 3);
            marker_table[3].Location = new System.Drawing.Point(10, base_offset * 4);
            marker_table[4].Location = new System.Drawing.Point(240, base_offset * 1);
            marker_table[5].Location = new System.Drawing.Point(240, base_offset * 2);

            // input object
            nu_freq.Location = new System.Drawing.Point(80, base_offset * 1);
            tb_hi_level.Location = new System.Drawing.Point(80, base_offset * 2);
            tb_lo_level.Location = new System.Drawing.Point(80, base_offset * 3);
            nu_duty.Location = new System.Drawing.Point(80, base_offset * 4);
            nu_tr.Location = new System.Drawing.Point(320, base_offset * 1);
            nu_tf.Location = new System.Drawing.Point(320, base_offset * 2);

            for (int i = 0; i < marker_table.Length; i++)
            {
                func_grop.Controls.Add(marker_table[i]);
                func_grop.Controls.Add(table[i]);
                marker_table[i].Width = 70;
            }

            // initial value part
            nu_tr.Maximum = 100000;
            nu_tf.Maximum = 100000;
            nu_freq.Maximum = 100000;
            nu_duty.Maximum = 100;

            tb_hi_level.Text = "0.25,0.5";
            tb_lo_level.Text = "0";
            nu_freq.Value = 100;
            nu_duty.Value = 50;
            nu_tf.Value = 20;
            nu_tr.Value = 20;


        }

        // interface
        GroupBox inter_grop = new GroupBox();
        CheckBox ck_select = new CheckBox();
        NumericUpDown nu_slave = new NumericUpDown();
        DataGridView dg_swire = new DataGridView();
        TextBox tb_bin = new TextBox();
        NumericUpDown nu_swire_row = new NumericUpDown();
        Button bt_bin = new Button();

        private void InterFaceConfig()
        {
            inter_grop.Text = "Interface Config";
            inter_grop.Size = new System.Drawing.Size(500, 240);
            inter_grop.Location = new System.Drawing.Point(10, 170);
            tabPage2.Controls.Add(inter_grop);

            int base_offset = 20;
            ck_select.Text = "swire enable";
            ck_select.Location = new System.Drawing.Point(10, base_offset);
            ck_select.Width = 90;
            inter_grop.Controls.Add(ck_select);
            
            Label lab_slave = new Label();
            lab_slave.Text = "Slave ID";
            lab_slave.Location = new System.Drawing.Point(100, base_offset + 5);
            lab_slave.Width = 50;
            inter_grop.Controls.Add(lab_slave);

            nu_slave.Location = new System.Drawing.Point(150, base_offset);
            nu_slave.Hexadecimal = true;
            nu_slave.Maximum = 255;
            nu_slave.Width = 60;
            nu_slave.Value = 0x46;
            inter_grop.Controls.Add(nu_slave);
            Label lab_bin = new Label();
            lab_bin.Text = "Bin folder";
            lab_bin.Location = new System.Drawing.Point(220, base_offset + 5);
            lab_bin.Width = 60;
            inter_grop.Controls.Add(lab_bin);
            tb_bin.Location = new System.Drawing.Point(280, base_offset);
            inter_grop.Controls.Add(tb_bin);

            bt_bin = new Button();
            bt_bin.Text = "Open Folder";
            bt_bin.Location = new System.Drawing.Point(380, base_offset);
            bt_bin.Click += bt_bin_Click;
            inter_grop.Controls.Add(bt_bin);

            dg_swire.Location = new System.Drawing.Point(10, base_offset * 3);
            dg_swire.Width = 125;
            dg_swire.ColumnCount = 1;
            dg_swire.RowHeadersVisible = false;
            dg_swire.ColumnHeadersVisible = false;
            dg_swire.Columns[0].Width = 120;
            dg_swire.AllowUserToAddRows = false;
            dg_swire.RowCount = 1;

            nu_swire_row.Location = new System.Drawing.Point(140, base_offset * 3);
            nu_swire_row.Width = 60;
            nu_swire_row.ValueChanged += Nu_swire_row_ValueChanged;
            inter_grop.Controls.Add(nu_swire_row);
            inter_grop.Controls.Add(dg_swire);
            tabPage2.Controls.Add(inter_grop);
        }

        private void Nu_swire_row_ValueChanged(object sender, EventArgs e)
        {
            dg_swire.RowCount = (int)nu_swire_row.Value;
            //throw new NotImplementedException();
        }

        private void bt_bin_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrow = new FolderBrowserDialog();
            if (FolderBrow.ShowDialog() == DialogResult.OK)
            {
                tb_bin.Text = FolderBrow.SelectedPath;
            }
        }


        // Eload 
        GroupBox eload_grop = new GroupBox();
        TextBox tb_eload = new TextBox();
        private void EloadConfig()
        {
            int base_point = 20;
            eload_grop.Text = "Eload Config";
            eload_grop.Location = new System.Drawing.Point(10, 420);
            eload_grop.Width = 500;
            eload_grop.Height = 60;
            
            Label lab_eload = new Label();
            lab_eload.Text = "ELoad (A)";
            lab_eload.Width = 80;
            lab_eload.Location = new System.Drawing.Point(10, base_point + 5);
            eload_grop.Controls.Add(lab_eload);

            tb_eload.Text = "0.2,0.1";
            tb_eload.Location = new System.Drawing.Point(100, base_point);
            tb_eload.Width = 250;
            eload_grop.Controls.Add(tb_eload);

            tabPage2.Controls.Add(eload_grop);
        }

        private void TDMA_Interface()
        {
            FuncgenConfig();
            InterFaceConfig();
            EloadConfig();
        }





    }
}
