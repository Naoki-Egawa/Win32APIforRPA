using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Win32APIforRPA
{
    class Win32API
    {
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int MK_LBUTTON = 0x0001;
        public static int GWL_STYLE = -16;
        public const int WM_SETTEXT = 0x000C;
        public const int WM_GETTEXT = 0x000D;
        public const Int32 WM_COPYDATA = 0x4A;
        public const int EM_GETSEL = 0x00B0;
        public const int WM_PASTE = 0x302;
        public const Int32 WM_USER = 0x400;
        public const int WM_GETTEXTLENGTH = 0x00E;
        private const int WM_CLOSE = 0x10;
        const int BM_CLICK = 0x00F5;
        public const int SC_CLOSE = 0xF060;
        public const int WM_SYSCOMMAND = 0x0112;

        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_CHAR = 0x0102;
        public const int SC_MAXIMIZE = 0xF030;
        public const int WM_ACTIVATE = 0x0006;

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }


        public static void InputKey(IntPtr hnd, uint number)
        {
            SetFocus(hnd);
            //Thread.Sleep(50);

            SendMessage(hnd, WM_KEYDOWN, number, 0);
            Thread.Sleep(10);
            SendMessage(hnd, WM_CHAR, number, 0);
            Thread.Sleep(10);

            SendMessage(hnd, WM_KEYUP, number, 0);

        }
        public static RECT GetRect(IntPtr hwnd)
        {
            Win32API.RECT rect;
            bool flag = Win32API.GetWindowRect(hwnd, out rect);

            return rect;
        }

        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        //[DllImport("user32.dll", SetLastError = true)]
        //private static extern bool WindowActivate(int wParam, int lParam, IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern int VkKeyScan(char ch);


        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool CloseHandle(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]



        public static extern void keybd_event(
    int bVk,            // 仮想キーコード
    int bScan,          // ハードウェアスキャンコード
    string dwFlags,        // 動作指定フラグ
    IntPtr dwExtraInfo // 追加情報
);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(
    IntPtr hWnd,     // ウィンドウハンドル
    int x,        // x座標
    int y,        // y座標
    int nWitdh,   // 幅
    int nHeight,  // 高さ
    bool bRepaint  // 再描画フラグ
            );

        public static void WindowSizeMax(IntPtr Whnd)
        {
            Console.WriteLine(Convert.ToString((int)Whnd));
            SendMessage(Whnd, WM_SETTEXT, SC_CLOSE, 0);

        }
        public uint GetNumCode(int number)
        {
            if (number == 0)
            {
                return 0x30;
            }
            else if (number == 1)
            {
                return 0x31;
            }
            else if (number == 2)
            {
                return 0x32;
            }
            else if (number == 3)
            {
                return 0x33;
            }
            else if (number == 4)
            {
                return 0x34;
            }
            else if (number == 5)
            {
                return 0x35;
            }
            else if (number == 6)
            {
                return 0x36;
            }
            else if (number == 7)
            {
                return 0x37;
            }
            else if (number == 8)
            {
                return 0x38;
            }
            else if (number == 9)
            {
                return 0x39;
            }
            else
            {
                return 0x30;
            }
        }
        public static void WindowActivate(IntPtr Whnd)
        {
            SendMessage(Whnd, WM_ACTIVATE, 0, 0);

        }

        int interval = 5000;
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;
        private const int MOUSEEVENTF_RIGHTTDOWN = 0x8;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        [DllImport("USER32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        int[] cursorLogX = new int[10] { -1, -1, -1, -1, -1, -1, -1, -1, -1, -1 };
        int[] cursorLogY = new int[10] { -1, -1, -1, -1, -1, -1, -1, -1, -1, -1 };
        int n = 0;


        public static void MouseRClick(int x, int y, int sleepTime = 100)
        {

            Console.WriteLine($"{x.ToString()}-{y.ToString()}");
            //マウスポインタの位置を画面左上基準の座標(x,y)にする
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_RIGHTTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(sleepTime);
        }
        public static void MouseLClick(int x, int y, int sleepTime = 100)
        {

            Console.WriteLine($"{x.ToString()}-{y.ToString()}");
            //マウスポインタの位置を画面左上基準の座標(x,y)にする
            Cursor.Position = new System.Drawing.Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(sleepTime);
        }

        /// <summary>
        /// ターゲットのハンドル（テキストボックス）へtextをsetする
        /// </summary>
        /// <param name="Whnd"></param>
        /// <param name="setText"></param>
        public static async void SetText(Window Whnd, string WindowName, string setText)
        {

            IntPtr mainWindowHandle = GetDesktopWindow();
            var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);
            var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

            var temp = all.Where(x => x.Title == WindowName).FirstOrDefault();
            //MessageBox.Show($"{temp.Title}_{temp.ClassName}");

            var targetwHnd = Win32API.GetTargetForm(temp.Title, temp.ClassName);
            Win32API.ShowWindow(targetwHnd.hWnd, 1);
            Win32API.SetForegroundWindow(targetwHnd.hWnd);

            await Task.Delay(500);
            Console.WriteLine($"Handle :{Convert.ToString((int)Whnd.hWnd, 16)} Caption:{Whnd.Title} Class:{Whnd.ClassName} Get Complete\r\n");

            //SetFocus(targetwHnd.hWnd);
            //await Task.Delay(500);
            SetFocus(Whnd.hWnd);
            await Task.Delay(500);

            SendMessage(Whnd.hWnd, WM_SETTEXT, 0, setText);
            await Task.Delay(500);

        }

        public static async void ObjectCenterClick(Window Whnd, int x, int y)
        {
            var rect = Win32API.GetRect(Whnd.hWnd);

            MouseLClick((rect.left + x), (rect.top + y), 200);
        }
        public static async void InputKeyNessesaryMouseClick(Window Whnd, string word, bool delKey = false)
        {
            var rect = Win32API.GetRect(Whnd.hWnd);

            MouseLClick((rect.left + 7), ((rect.bottom + rect.top) / 2), 500);

            if (delKey)
            {

                Win32API.InputKey(Whnd.hWnd, 0x2E);
                Win32API.InputKey(Whnd.hWnd, 0x2E);
            }

            foreach (var item in word)
            {
                var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                Win32API.InputKey(Whnd.hWnd, code);
            }
        }
        public static void ClickWhenShowForm(string WindowName, string classname, string FormClassName, int skipCount = 0)
        {

            bool isValue = true;
            int loogCount = 1;

            while (isValue)
            {
                if (loogCount == 30)
                {
                    throw new InvalidOperationException();
                }
                else
                {
                    isValue = !Win32API.IsShowForm(WindowName, FormClassName);
                    Thread.Sleep(500);
                    loogCount++;
                }
            }

            var bt = Win32API.GetTargetControl(WindowName, classname, skipCount);
            Win32API.ButtonClick(bt);
            Thread.Sleep(500);

        }

        public static bool IsShowTargetObject(string WindowName, string classname, string captionname, int skipCount = 0)
        {

            try
            {

                var hnd = Win32API.GetTargetControl(WindowName, classname, captionname);

                if (hnd == null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception)
            {

                return false;
            }
        }
        public static bool IsShowForm(string WindowName, string classname, int skipCount = 0)
        {
            try
            {

                IntPtr mainWindowHandle = GetDesktopWindow();
                var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);

                var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

                var temp = all.Where(x => x.ClassName == classname & x.Title == WindowName);

                int i = 0;
                string a = "";
                foreach (var item in temp)
                {
                    a = a + $" Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title})\r\n";
                    Console.WriteLine(a);
                    i++;
                }

                if (temp.Count() >= 1)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }

        }

        public static async Task ExecuteButtonClick(string WindowName, string CaptionName, int indexNumber = 0)
        {
            IntPtr mainWindowHandle = GetDesktopWindow();
            var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);
            var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

            foreach (var item in all)
            {
                Console.WriteLine($"{item.Title}-{item.Style}-{item.hWnd}");
            }

            Thread.Sleep(3000);

            var temp = all.Where(x => x.Title == WindowName).FirstOrDefault();

            var targetwHnd = Win32API.GetTargetForm(temp.Title, temp.ClassName);

            Win32API.ShowWindow(targetwHnd.hWnd, 1);
            Win32API.SetForegroundWindow(targetwHnd.hWnd);

            var wHnd = GetWHnd(GetWindow(WindowhWnd), CaptionName, indexNumber);
            Thread.Sleep(500);


            SetFocus(targetwHnd.hWnd);

            Thread.Sleep(1000);
            SendMessage(wHnd.hWnd, BM_CLICK, 0, 0);
            Thread.Sleep(500);
            //await Task.Delay(1000);
            Thread.Sleep(1000);

        }

        public static async Task ButtonClick(Window wHnd)
        {


            await Task.Delay(500);

            //SetFocus(wHnd.hWnd);
            //SendMessage(wHnd.hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0);
            //await Task.Delay(500);


            //SendMessage(wHnd.hWnd, WM_LBUTTONUP, 0x00000000, 0x000A000A);
            //await Task.Delay(500);

            SendMessage(wHnd.hWnd, BM_CLICK, 0, 0);

            //await Task.Delay(1000);
            Thread.Sleep(500);

        }
        public static string GetText(Window Whnd)
        {
            try
            {

                var sb = new StringBuilder();
                var length = SendMessage(Whnd.hWnd, WM_GETTEXTLENGTH, 0, 0);
                sb.Capacity = length + 20; //追記2
                Task.Delay(500);  //無くても同じ
                length = SendMessage(Whnd.hWnd, WM_GETTEXT, length + 20, sb);

                return sb.ToString();
            }
            catch (Exception)
            {

                return "";
            }

        }

        //★★★★★★★★★★★★★★★★★★★★★★★
        //★★これより下はサブのメソッド★★★★★★★★
        //★★★★★★★★★★★★★★★★★★★★★★★

        public static Window GetWHnd(Window top, string captionName, int skipCount = 0)
        {
            var all = GetAllChildWindows(top, new List<Window>());
            var temp = all.Where(x => x.Title == captionName);

            int i = 0;

            string a = "";

            foreach (var item in temp)
            {
                a = a + $" Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title})\r\n";
                i++;
            }
            Console.WriteLine(a);

            return temp.Skip(skipCount).First();
        }


        // 指定したウィンドウの全ての子孫ウィンドウを取得し、リストに追加する
        public static List<Window> GetAllChildWindows(Window parent, List<Window> dest)
        {
            dest.Add(parent);
            EnumChildWindows(parent.hWnd).ToList().ForEach(x => GetAllChildWindows(x, dest));
            return dest;
        }

        // 与えた親ウィンドウの直下にある子ウィンドウを列挙する（孫ウィンドウは見つけてくれない）
        public static IEnumerable<Window> EnumChildWindows(IntPtr hParentWindow)
        {
            IntPtr hWnd = IntPtr.Zero;
            while ((hWnd = FindWindowEx(hParentWindow, hWnd, null, null)) != IntPtr.Zero) { yield return GetWindow(hWnd); }
        }

        // ウィンドウハンドルを渡すと、ウィンドウテキスト（ラベルなど）、クラス、スタイルを取得してWindowsクラスに格納して返す
        public static Window GetWindow(IntPtr hWnd)
        {
            int textLen = GetWindowTextLength(hWnd);
            string windowText = null;
            if (0 < textLen)
            {
                //ウィンドウのタイトルを取得する
                StringBuilder windowTextBuffer = new StringBuilder(textLen + 1);
                GetWindowText(hWnd, windowTextBuffer, windowTextBuffer.Capacity);
                windowText = windowTextBuffer.ToString();
            }

            //ウィンドウのクラス名を取得する
            StringBuilder classNameBuffer = new StringBuilder(256);
            GetClassName(hWnd, classNameBuffer, classNameBuffer.Capacity);

            // スタイルを取得する
            int style = GetWindowLong(hWnd, GWL_STYLE);
            return new Window() { hWnd = hWnd, Title = windowText, ClassName = classNameBuffer.ToString(), Style = style };
        }
        public static Window GetTargetControl(string WindowName, string classname, string captionname, int skipCount = 0)
        {

            IntPtr mainWindowHandle = GetDesktopWindow();
            var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);

            var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

            var temp = all.Where(x => x.Title == WindowName);

            int i = 0;
            string a = "";
            foreach (var item in temp)
            {
                a = a + $"Form Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title} Class:{item.ClassName}\r\n";
                i++;
            }
            //Console.WriteLine(a);

            var b = temp.Skip(0).First();
            Win32API.ShowWindow(b.hWnd, 1);
            Win32API.SetForegroundWindow(b.hWnd);

            SetFocus(b.hWnd);

            WindowhWnd = FindWindowEx(temp.Skip(0).First().hWnd, IntPtr.Zero, WindowName, null);

            var ControlObject = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

            var temp2 = ControlObject.Where(x => x.ClassName == classname & x.Title == captionname);

            i = 0;

            foreach (var item in temp2)
            {
                a = a + $"Control Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title} Class:{item.ClassName}\r\n";
                i++;
            }
            Console.WriteLine(a);
            //MessageBox.Show("正常");
            return temp2.Skip(skipCount).First();

        }
        public static Window GetTargetControl(string WindowName, string classname, int skipCount = 0)
        {
            for (int h = 0; h < 10; h++)
            {

                try
                {
                    IntPtr mainWindowHandle = GetDesktopWindow();
                    var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);

                    var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

                    var temp = all.Where(x => x.Title == WindowName);

                    int i = 0;
                    string a = "";
                    foreach (var item in temp)
                    {
                        a = a + $"Form Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title} Class:{item.ClassName}\r\n";
                        i++;
                    }
                    //Console.WriteLine(a);

                    var b = temp.Skip(0).First();
                    Win32API.ShowWindow(b.hWnd, 1);
                    Win32API.SetForegroundWindow(b.hWnd);

                    SetFocus(b.hWnd);

                    WindowhWnd = FindWindowEx(temp.Skip(0).First().hWnd, IntPtr.Zero, WindowName, null);

                    var ControlObject = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

                    var temp2 = ControlObject.Where(x => x.ClassName == classname);

                    i = 0;

                    foreach (var item in temp2)
                    {
                        a = a + $"Control Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title} Class:{item.ClassName}\r\n";
                        i++;
                    }
                    Console.WriteLine(a);
                    //MessageBox.Show("正常");
                    return temp2.Skip(skipCount).First();

                }
                catch (Exception)
                {
                    Thread.Sleep(500);
                    continue;
                }
            }

            throw new Exception();

        }
        public static Window GetTargetForm(string WindowName, string classname, int skipCount = 0)
        {
            IntPtr mainWindowHandle = GetDesktopWindow();
            var WindowhWnd = FindWindowEx(mainWindowHandle, IntPtr.Zero, WindowName, null);

            var all = GetAllChildWindows(GetWindow(WindowhWnd), new List<Window>());

            var temp = all.Where(x => x.ClassName == classname & x.Title == WindowName);

            int i = 0;
            string a = "";
            foreach (var item in temp)
            {
                a = a + $" Index:{i} Handle :{Convert.ToString((int)item.hWnd, 16)} Caption:{item.Title} Class:{item.ClassName}\r\n";
                i++;
            }
            Console.WriteLine(a);

            return temp.Skip(skipCount).First();
        }
    }
}
