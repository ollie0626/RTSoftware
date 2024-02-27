using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scope_Simple_tool.MyPage
{
    public abstract class FileControl
    {
        public string Name = "";
        public abstract void LoadData(string str);
        public abstract void SaveData();

        static public List<string> SaveString = new List<string>();
    }

    public class MainFile : FileControl
    {
        public MainFile()
        {
            Name = "Main";
        }

        public override void SaveData()
        {
            string TmpStr = "";
            TmpStr += Pages.ShellViewModel._InsName + "&";
            TmpStr += Pages.ShellViewModel._WaveFormPath + "&";
            TmpStr += Pages.ShellViewModel._FileName + "&";
            TmpStr += Pages.ShellViewModel._slave.ToString() + "&";
            TmpStr += Pages.ShellViewModel._BinFile + "&";
            SaveString.Add(Name + ";" + TmpStr);
        }

        public override void LoadData(string str)
        {
            string[] NameArr = str.Split(';');
            string[] UIInfo = str.Split('&');

            Pages.ShellViewModel._InsName = UIInfo[0].Split(';')[1];
            Pages.ShellViewModel._WaveFormPath = UIInfo[1];
            Pages.ShellViewModel._FileName = UIInfo[2];
            Pages.ShellViewModel._slave = Convert.ToInt16(UIInfo[3]);
            Pages.ShellViewModel._BinFile = UIInfo[4];
        }
    }

    public class LoadIni
    {
        static List<FileControl> LoadFiles = new List<FileControl>()
        {
            new MainFile()
        };

        public static void LoadFile(string path = "MainIni.txt")
        {
            if (!System.IO.File.Exists(path)) return;
            string[] lines = System.IO.File.ReadAllLines(path);
            foreach(string line in lines)
            {
                string trimstr = line.Trim();
                string[] TmpStr = trimstr.Split(';');

                for (int i = 0; i < LoadFiles.Count; ++i)
                {
                    if (TmpStr[0].Contains(LoadFiles[i].Name))
                    {
                        try
                        {
                            LoadFiles[i].LoadData(trimstr);
                        }
                        catch { };
                    }
                }
            }
        }

        public static void SaveFile(string path = "MainIni.txt")
        {
            FileControl.SaveString.Clear();
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path))
            {
                for (int i = 0; i < LoadFiles.Count; ++i)
                {
                    try
                    {
                        LoadFiles[i].SaveData();
                    }
                    catch { };
                }

                foreach (string line in FileControl.SaveString)
                {
                    file.WriteLine(line);
                }
            }
        }


    }



}
