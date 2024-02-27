using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scope_Simple_tool.MyPage
{
    public class ConsoleControl
    {
        /* call win32 api */
        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        public static extern void FreeConsole();

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetConsoleWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);


        private static IntPtr HWnd;

        public ConsoleControl()
        {
            /* Create Console of application */
            AllocConsole();
            //Console.SetWindowPosition(0, 0);
            HWnd = GetConsoleWindow();
            /* Hiden Console Window */
            ShowWindow(HWnd, 0);

            /* Window setting */
            ConsoleColor oriColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("* Don't close this console window or the application will also close.");
            Console.ForegroundColor = oriColor;
        }


        static public void Show_AppConsole()
        {
            ShowWindow(HWnd, 1);
        }


        static public void Hide_AppConsole()
        {
            ShowWindow(HWnd, 0);
        }

        static public void Clear_AppConsole()
        {
            Console.Clear();
            ConsoleColor oriColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("* Don't close this console window or the application will also close.");
            Console.ForegroundColor = oriColor;
        }


    }
}
