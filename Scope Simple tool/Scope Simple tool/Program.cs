
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace Scope_Simple_tool
{
    class Program
    {

        static string[] dllname =
        {

            "PresentationCore",
            "PresentationFramework",

            "ControlzEx",
            "MahApps.Metro",
            "MahApps.Metro.IconPacks.Core",
            "MahApps.Metro.IconPacks",
            "MahApps.Metro.IconPacks.BoxIcons",
            "MahApps.Metro.IconPacks.JamIcons",
            "MahApps.Metro.IconPacks.Material",
            "MahApps.Metro.IconPacks.MaterialDesign",
            "MahApps.Metro.IconPacks.MaterialLight",
            "MahApps.Metro.IconPacks.PicolIcons",
            "InsLibDotNet",
            
            "RTBBLibDotNet",
            "RTBBLib",
            "RTBridgeBoardCore",

            "Stylet",
            "System.Windows.Interactivity",
            "Xceed.Wpf.Toolkit",

            "KeraLua",
            "lua54",
            "NLua"
        };


        [STAThreadAttribute]
        public static void Main()
        {
            //AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
            CreatDll();
            App.Main();
            DeleteDll();
        }

        private static void CreatDll()
        {
            for (int i = 0; i < dllname.Length; ++i)
            {
                string filename = dllname[i] + ".dll";
                string Source = dllname[i].Replace(".", "_");
                if (!File.Exists(filename))
                {

                    File.WriteAllBytes(filename, Scope_Simple_tool.Properties.Resources.ResourceManager.GetObject(Source) as byte[]);
                    FileInfo fileInfo = new FileInfo(filename);
                    fileInfo.Attributes = FileAttributes.Hidden;
                }
            }
        }

        private static void DeleteDll()
        {
            string cmd = "del.bat";
            Execute(cmd);
        }

        public static void Execute(string command)
        {
            var processInfo = new ProcessStartInfo("cmd.exe", "/S /C " + command)
            {
                CreateNoWindow = true,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            Process.Start(processInfo);

        }


        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            AssemblyName assemblyName = new AssemblyName(args.Name);
            string path = assemblyName.Name + ".dll";

            if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
            {
                path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);
            }
            using (Stream stream = executingAssembly.GetManifestResourceStream(@"D:\Desktop\Scope Simple tool\packages\dll\InsLibDotNet.dll"))
            {
                if (stream == null)
                    return null;
                byte[] assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return null;
            }
        }
    }
}
