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
using MetroFramework.Forms;
using MetroFramework.Controls;
using System.Collections;
using System.Runtime.InteropServices;

namespace GenBinTool
{

    public partial class main : MetroForm
    {

        byte[] defaultBuffer;
        byte[] genBuffer;
        ToolStripMenuItem[] _item = new ToolStripMenuItem[2];
        MetroGrid DV;
        List<string> NameList = new List<string>();
        List<string> UnitList = new List<string>();
        List<int> AddList = new List<int>();

        List<int> GDData1 = new List<int>();
        List<string> GDNote1 = new List<string>();

        List<int> GDData2 = new List<int>();
        List<string> GDNote2 = new List<string>();

        List<List<int>> DataList = new List<List<int>>();
        List<List<string>> NoteList = new List<List<string>>();

        List<bool> MixList = new List<bool>();

        int file_index;
        int ColIndex = 0;
        int RowIndex = 0;
        string cp_str = "";
        bool _ctrl = false;

        bool loadBinEn = false;

        string _Addr, _Data, _Note;


        private void InitGUI()
        {
            _item[0] = new ToolStripMenuItem("Delete", null, new EventHandler(ToolStripDelete_Click));
            _item[1] = new ToolStripMenuItem("Add", null, new EventHandler(ToolStripAdd_Click));


            metroGrid1.RowCount = 2;

            metroGrid1[0, 0].Value = "AVDD";
            metroGrid1[1, 0].Value = "00";
            metroGrid1[2, 0].Value = "10,20,30";
            metroGrid1[3, 0].Value = "1v,2v,3v";

            metroGrid1[0, 1].Value = "DVDD";
            metroGrid1[1, 1].Value = "01";
            metroGrid1[2, 1].Value = "11,21,31";
            metroGrid1[3, 1].Value = "6V,7V,8V";
        }

        private void ToolStripDelete_Click(object sender, EventArgs e)
        {
            DV.Rows.RemoveAt(DV.CurrentRow.Index);
        }

        private void ToolStripAdd_Click(object sender, EventArgs e)
        {
            DV.RowCount = DV.RowCount + 1; 
        }

        public main()
        {
            InitializeComponent();
            InitGUI();
        }

        private void metroGrid1_MouseDown(object sender, MouseEventArgs e)
        {
            DV = (MetroGrid)sender;
            if (e.Button == MouseButtons.Right)
            {
                if (DV.RowCount == 0)
                {
                    ContextMenuStrip menu = new ContextMenuStrip();
                    _item[0].Enabled = false;
                    menu.Items.AddRange(_item);
                    menu.Show(DV, new Point(e.X, e.Y));
                }
                else
                {
                    int row = DV.CurrentRow.Index;
                    if (row < 0) return;
                    ContextMenuStrip menu = new ContextMenuStrip();
                    _item[0].Enabled = true;
                    menu.Items.AddRange(_item);
                    menu.Show(DV, new Point(e.X, e.Y));
                }
            }
        }

        private void BTLoad_Click(object sender, EventArgs e)
        {

            openFileDialog1.Filter = "Bin files (*.bin)|*.bin|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream Fio = File.Open(openFileDialog1.FileName, FileMode.Open);
                BinaryReader binRead = new BinaryReader(Fio);
                FileInfo fileInfo = new FileInfo(openFileDialog1.FileName);
                byte[] Binbuffer = binRead.ReadBytes((int)fileInfo.Length);
                defaultBuffer = new byte[fileInfo.Length];
                genBuffer = new byte[fileInfo.Length];

                Array.Copy(Binbuffer, defaultBuffer, fileInfo.Length);
                Array.Copy(Binbuffer, genBuffer, fileInfo.Length);
                Fio.Close();
                binRead.Close();
                loadBinEn = true;
            }
        }

        private double CalculateTimes(double start, double end, double step)
        {
            double res = (end - start) / step;
            return res;
        }

        private void BTGen_Click(object sender, EventArgs e)
        {
            NameList.Clear();
            AddList.Clear();
            NoteList.Clear();
            DataList.Clear();
            MixList.Clear();

            GDData1.Clear();
            GDNote1.Clear();
            file_index = 0;

            if (!loadBinEn)
            {
                MetroFramework.MetroMessageBox.Show(this, "\n\n Please Load Default Bin file !!", "Gen Bin File Tooling");
                return;
            }
            else
            {

                if (metroGrid1.RowCount == 0 && metroGrid2.RowCount == 0)
                {
                    MetroFramework.MetroMessageBox.Show(this, "\n\n Bin File Conditions Fail!!", "Gen Bin File Tooling", MessageBoxButtons.YesNo);
                    return;
                }



                if (metroGrid1.RowCount != 0)
                {
                    int row = metroGrid1.RowCount;
                    int col = metroGrid1.ColumnCount;
                    for (int row_idx = 0; row_idx < row; row_idx++)
                    {
                        for (int col_idx = 0; col_idx < col - 2; col_idx++)
                        {
                            if (metroGrid1[col_idx, row_idx].Value == null)
                            {
                                MetroFramework.MetroMessageBox.Show(this, "\n\n Bin File Conditions Fail!!", "Gen Bin File Tooling", MessageBoxButtons.YesNo);
                                return;
                            }
                        }
                    }
                }

                if (metroGrid2.RowCount != 0)
                {
                    int row = metroGrid2.RowCount;
                    int col = metroGrid2.ColumnCount;
                    for (int row_idx = 0; row_idx < row; row_idx++)
                    {
                        for (int col_idx = 0; col_idx < col - 2; col_idx++)
                        {
                            if (metroGrid2[col_idx, row_idx].Value == null)
                            {
                                MetroFramework.MetroMessageBox.Show(this, "\n\n Bin File Conditions Fail!!", "Gen Bin File Tooling", MessageBoxButtons.YesNo);
                                return;
                            }
                        }
                    }
                }

                GetAllData();
                MetroFramework.MetroMessageBox.Show(this, "\n\n Gen Bin File ok!!", "Gen Bin File Tooling", MessageBoxButtons.YesNo);
            }
        }


        private void GetAllData()
        {
            // GD1
            int row1_len = metroGrid1.RowCount;
            for (int row_idx = 0; row_idx < row1_len; row_idx++)
            {
                if (Convert.ToBoolean(metroGrid1[4, row_idx].Value))
                {
                    NameList.Add((string)metroGrid1[0, row_idx].Value);
                    AddList.Add(Convert.ToInt16((string)metroGrid1[1, row_idx].Value, 16));
                    MixList.Add(Convert.ToBoolean(metroGrid1[5, row_idx].Value));

                    string data = (string)metroGrid1[2, row_idx].Value;
                    string[] tmp = data.Split(',');
                    List<int> step_tmp = new List<int>();
                    foreach (string loop in tmp)
                        step_tmp.Add(Convert.ToInt16(loop.Trim(), 16));
                    DataList.Add(step_tmp);

                    /* get register note */
                    data = (string)metroGrid1[3, row_idx].Value;
                    tmp = data.Split(',');
                    List<string> note_tmp = new List<string>();
                    foreach (string loop in tmp)
                        note_tmp.Add(loop);
                    NoteList.Add(note_tmp);
                }
            }


            int row2_len = metroGrid2.RowCount;
            for (int row_idx = 0; row_idx < row2_len; row_idx++)
            {
                if (Convert.ToBoolean(metroGrid2[8, row_idx].Value))
                {
                    NameList.Add((string)metroGrid2[0, row_idx].Value);
                    AddList.Add(Convert.ToInt16((string)metroGrid2[1, row_idx].Value, 16));
                    MixList.Add(Convert.ToBoolean(metroGrid2[9, row_idx].Value));
                    UnitList.Add((string)metroGrid2[7, row_idx].Value);

                    int start = Convert.ToInt32((string)metroGrid2[2, row_idx].Value, 16);
                    int end = Convert.ToInt32((string)metroGrid2[3, row_idx].Value, 16);
                    int step = Convert.ToInt32((string)metroGrid2[4, row_idx].Value, 16);

                    // for volt calculate
                    double start_val = Convert.ToDouble(metroGrid2[5, row_idx].Value);
                    double step_val = Convert.ToDouble(metroGrid2[6, row_idx].Value);
                    double cnt = CalculateTimes(start, end, step);
                    List<int> tmp = new List<int>();
                    List<string> val = new List<string>();
                    for (int j = 0; j <= cnt; j++)
                    {
                        tmp.Add(start + (step * j));
                        val.Add((start_val + (step_val * j)).ToString() + (string)metroGrid2[7, row_idx].Value);
                    }

                    /* [] --> row, [] --> data */
                    DataList.Add(tmp);
                    NoteList.Add(val);
                }
            }

            int size = 1;
            for (int i = 0; i < MixList.Count; i++)
            {
                if (MixList[i])
                {
                    size = size * DataList[i].Count;
                }
            }
            GenBinFlow(size);
        }

        private void GenBinFlow(int size)
        {
            int cnt = MixList.Count;
            string number = Convert.ToString(size);
            List<string> MixfileList = new List<string>();
            List<string> SumfileList = new List<string>();
            
            System.Collections.Generic.Dictionary<int, byte[]> mix = new Dictionary<int, byte[]>();
            System.Collections.Generic.Dictionary<int, byte[]> sum = new Dictionary<int, byte[]>();

            //int file_idx = 0;

            int sum_idx = 0;
            int mix_idx = 0;
            //bool flag = false;
            for (int i = 0; i < cnt; i++)
            {
                //flag = false;
                if (MixList[i])
                {
                    for (; mix_idx < size;)
                    {
                        int origin = mix_idx;
                        int indexSelect = 0;
                        string fileName = "";
                        //int offset = 1;
                        System.Collections.Generic.Dictionary<int, byte> map = new System.Collections.Generic.Dictionary<int, byte>();

                        for (int data_idx = 0; data_idx < MixList.Count; data_idx++)
                        {
                            if (MixList[data_idx])
                            {
                                indexSelect = origin % DataList[data_idx].Count;
                                origin = origin / DataList[data_idx].Count;
                                if (!map.ContainsKey(AddList[data_idx]))
                                {
                                    map.Add(AddList[data_idx], (byte)DataList[data_idx][indexSelect]);
                                    //flag = true;
                                }
                                else
                                {
                                    map[AddList[data_idx]] |= (byte)DataList[data_idx][indexSelect];
                                }
                            }
                        }
                        origin = mix_idx;
                        byte[] tmp = new byte[genBuffer.Length];
                        Array.Copy(genBuffer, tmp, genBuffer.Length);
                        for (int data_idx = 0; data_idx < MixList.Count; data_idx++)
                        {
                            if (MixList[data_idx])
                            {
                                indexSelect = origin % DataList[data_idx].Count;
                                origin = origin / DataList[data_idx].Count;
                                tmp[AddList[data_idx]] = (byte)map[AddList[data_idx]];
                                fileName += NameList[data_idx] + "_" + NoteList[data_idx][indexSelect] + "_";
                            }
                        }
                        fileName = fileName.Substring(0, fileName.Length - 1);
                        MixfileList.Add(fileName);
                        //Console.WriteLine("{0:X}", tmp[0x2c]);
                        mix.Add(mix_idx, tmp);
                        mix_idx++;
                    }
                }
            }

            for (int i = 0; i < cnt; i++)
            {
                if (!MixList[i])
                {
                    
                    for (int idx = 0; idx < ((mix.Count == 0) ? 1 : mix.Count); idx++)
                    {
                        for (int data_idx = 0; data_idx < DataList[i].Count; data_idx++)
                        {
                            byte[] tmp = new byte[defaultBuffer.Length];
                            string file_name = "";
                            if(mix.Count != 0)
                            {
                                Array.Copy(mix[idx], tmp, mix[idx].Length);
                                file_name = MixfileList[idx] + "_";
                                file_name += NameList[i] + "_" + NoteList[i][data_idx];
                                tmp[AddList[i]] = (byte)DataList[i][data_idx];
                            }
                            else
                            {
                                Array.Copy(defaultBuffer, tmp, defaultBuffer.Length);
                                //file_name = string.Format("{0}_", file_idx++);
                                file_name += NameList[i] + "_" + NoteList[i][data_idx];
                                tmp[AddList[i]] = (byte)DataList[i][data_idx];
                            }
                            //Console.WriteLine(file_name);
                            
                            Console.WriteLine("{0:X}", DataList[i][data_idx]);
                            sum.Add(sum_idx++, tmp);
                            SumfileList.Add(file_name);
                        }
                    }
                }
            }

            if(sum.Count != 0)
            {
                for(int i = 0; i < sum.Count; i++)
                {
                    SumfileList[i] = string.Format("{0}_" + SumfileList[i], i + numericUpDown1.Value);
                    //Console.WriteLine(SumfileList[i]);

                    string path = tbPath.Text + @"\" + SumfileList[i] + ".bin";
                    FileStream Fio = File.Create(path);
                    BinaryWriter bWrite = new BinaryWriter(Fio);
                    bWrite.Write(sum[i]);
                    Fio.Close(); bWrite.Close();
                    Fio.Dispose(); bWrite.Dispose();
                }
                    
            }
            else if(mix.Count != 0)
            {
                for (int i = 0; i < mix.Count; i++)
                {
                    MixfileList[i] = string.Format("{0}_" + MixfileList[i], i + numericUpDown1.Value);
                    //Console.WriteLine(MixfileList[i]);

                    string path = tbPath.Text + @"\" + MixfileList[i] + ".bin";
                    FileStream Fio = File.Create(path);
                    BinaryWriter bWrite = new BinaryWriter(Fio);
                    bWrite.Write(mix[i]);
                    Fio.Close(); bWrite.Close();
                    Fio.Dispose(); bWrite.Dispose();
                }
            }
        }

        private void SaveCSV(MetroGrid dt, FileMode mode, string path)
        {
            FileInfo fio = new FileInfo(path);
            if (dt.RowCount == 0) return;
            if (!fio.Directory.Exists)
            {
                fio.Directory.Create();
            }

            FileStream fs = new FileStream(path, mode, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";

            for (int i = 0; i < dt.ColumnCount; i++)
            {
                data += dt.Columns[i].HeaderText.ToString();
                if (i < dt.Columns.Count - 1) data += ",";
            }
            sw.WriteLine(data);

            for (int i = 0; i < dt.RowCount; i++)
            {
                data = "";
                for (int j = 0; j < dt.ColumnCount; j++)
                {
                    string str = "";
                    try
                    {
                        str = dt[j, i].Value.ToString();
                    }
                    catch
                    {
                        str = Convert.ToBoolean(dt[j, i].Value).ToString();
                    }
                    
                    str = str.Replace(",", "$");
                    str = str.Replace("\"", "\"\"");
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n'))
                    {
                        str = string.Format("\"{0}\"", str);
                    }
                    data += str;
                    if (j < dt.Columns.Count - 1) data += ",";
                }
                sw.WriteLine(data);
            }

            sw.Close();
            fs.Close();
        }

        private void OpenCSV(string path)
        {
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs);
            string strLine = "";
            string[] arLine = null;
            int columnCount = 0;

            while ((strLine = sr.ReadLine()) != null)
            {
                columnCount = strLine.Split(',').Length;
                arLine = strLine.Split(',');
                if (columnCount == 6)
                {
                    arLine[2] = arLine[2].Replace("$", ",");
                    arLine[3] = arLine[3].Replace("$", ",");
                    if (arLine[0] != "Name")
                        metroGrid1.Rows.Add(arLine);
                }
                else
                {
                    if (arLine[0] != "Name")
                        metroGrid2.Rows.Add(arLine);
                }
            }
            sr.Close();
            fs.Close();
        }

        private void BTSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "csv files (*.csv)|*.csv|All files(*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SaveCSV(metroGrid1, FileMode.Create, saveFileDialog1.FileName);
                SaveCSV(metroGrid2, FileMode.Append, saveFileDialog1.FileName);
            }
        }

        private void BTOpen_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "csv files (*.csv)|*.csv|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                metroGrid1.RowCount = 0;
                metroGrid2.RowCount = 0;
                OpenCSV(openFileDialog1.FileName);
            }
        }

        private void BTSet_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            if (path.ShowDialog() == DialogResult.OK)
            {
                tbPath.Text = path.SelectedPath;
            }
        }

        private void metroGrid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ColIndex = e.ColumnIndex;
            RowIndex = e.RowIndex;
            //cp_str = (string)metroGrid1[ColIndex, RowIndex].Value;
            //Console.WriteLine(cp_str);
        }

        private void metroGrid1_CellMouseLeave(object sender, DataGridViewCellMouseEventArgs e)
        {
            DV = (MetroGrid)sender;
            for (int row = 0; row < DV.RowCount - 2; row++)
            {
                if (DV.Name == "metroGrid1")
                {
                    for (int col = 1; col < 3; col++)
                    {
                        if (DV[col, row].Value != null)
                        {
                            DV[col, row].Value = DV[col, row].Value.ToString().ToUpper();
                        }
                    }
                }
                else
                {
                    for (int col = 1; col < 5; col++)
                    {
                        if (DV[col, row].Value != null)
                        {
                            DV[col, row].Value = DV[col, row].Value.ToString().ToUpper();
                        }
                    }
                }
            }
        }
    }
}
