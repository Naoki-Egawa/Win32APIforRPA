using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Win32APIforRPA
{
    public static class ExpantionWin32API
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

        //int interval = 5000;
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;
        private const int MOUSEEVENTF_RIGHTTDOWN = 0x8;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;
        [DllImport("USER32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        //int[] cursorLogX = new int[10] { -1, -1, -1, -1, -1, -1, -1, -1, -1, -1 };
        //int[] cursorLogY = new int[10] { -1, -1, -1, -1, -1, -1, -1, -1, -1, -1 };
        //int n = 0;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, StringBuilder lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        public static async Task Click(this Window window, int sleetTime = 1500)
        {
            var a = $"選択したオブジェクトハンドル Handle :{Convert.ToString((int)window.hWnd, 16)} Caption:{window.Title} Class:{window.ClassName}\r\n";

            Console.WriteLine(a);

            //SetFocus(wHnd.hWnd);
            //SendMessage(wHnd.hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0);
            //await Task.Delay(sleetTime);

            //SendMessage(wHnd.hWnd, WM_LBUTTONUP, 0x00000000, 0x000A000A);
            //await Task.Delay(500);

            SendMessage(window.hWnd, BM_CLICK, 0, 0);

        }

        public static void ObjectCenterClick(this Window Whnd, int sleeptime = 1000)
        {
            var rect = Win32API.GetRect(Whnd.hWnd);

            Thread.Sleep(500);
            MouseLClick(((rect.left + rect.right) / 2), ((rect.bottom + rect.top) / 2), sleeptime);
        }

        public static Window ObjectLeftEndClick(this Window Whnd, int sleeptime = 1000)
        {
            var rect = Win32API.GetRect(Whnd.hWnd);

            Thread.Sleep(500);
            MouseLClick(((rect.left + 5)), ((rect.bottom + rect.top) / 2), sleeptime);
            return Whnd;
        }

        public static void MouseLClick(int x, int y, int sleepTime = 500)
        {

            Console.WriteLine($"{x.ToString()}-{y.ToString()}");
            //マウスポインタの位置を画面左上基準の座標(x,y)にする
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(sleepTime);
        }
    }
}
