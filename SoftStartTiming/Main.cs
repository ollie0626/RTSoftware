using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Forms.VisualStyles;

namespace SoftStartTiming
{
    public partial class Main : Form
    {

        string win_name = "Main v1.14";

        public Main()
        {
            InitializeComponent();
            this.Text = win_name;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CBItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            Form form;
            switch(CBItem.SelectedIndex)
            {
                case 0:
                    form = new SoftStartTiming();
                    form.ShowDialog();
                    break;
                case 1:
                    form = new CrossTalk();
                    form.ShowDialog();
                    break;
                case 2:
                    form = new VIDIO();
                    form.ShowDialog();
                    break;
                case 3:
                    form = new VIDI2C();
                    form.ShowDialog();
                    break;
                case 4:
                    form = new LoadTransient();
                    form.ShowDialog();
                    break;
                case 5:
                    form = new LTLab();
                    form.ShowDialog();
                    break;
                   
            }
        }

        private void textToBinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Info |*.txt";

            if(openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog.FileName;
                List<byte> data = new List<byte>();
                using (StreamReader reader = new StreamReader(path))
                {
                    
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        Console.WriteLine(line);
                        string[] spilt_list = line.Split('\t');

                        data.Add(Convert.ToByte(spilt_list[1], 16));
                    }
                }

                using (BinaryWriter binWriter = new BinaryWriter(File.Open(Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + ".bin", FileMode.Create)))
                {
                    binWriter.Write(data.ToArray());
                }
            }
        }
    }
}
