using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Win32APIforRPA
{
    class RunSaisinsaEntry
    {
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

        public string TitleName = "給付－過去レセプトエラー訂正・削除";

        public bool ChouseizumiFlag = false;

        public void ShowOldRece(ReceSearchValue t, bool kako = false, bool Validation = false)
        {
            var kakoTitle = "";
            var winN = "";
            Window hnd;
            Win32API.RECT rect;
            string gengo = "9";

            if (!kako)
            {
                var KosmoPass = @"C:\Program Files\dirapp\kosmo21\v6\ABE203X.exe";

                var isExsist = File.Exists(KosmoPass);

                if (!isExsist)
                {
                    KosmoPass = @"C:\Program Files (x86)\dirapp\kosmo21\v6\ABE203X.exe";
                }

                System.Diagnostics.Process.Start(KosmoPass);
            }
            else
            {
                kakoTitle = "過去";
                var KosmoPass = @"C:\Program Files\dirapp\kosmo21\v6\ABE206X.exe";

                var isExsist = File.Exists(KosmoPass);

                if (!isExsist)
                {
                    KosmoPass = @"C:\Program Files (x86)\dirapp\kosmo21\v6\ABE206X.exe";
                }

                System.Diagnostics.Process.Start(KosmoPass);
                winN = "過去のレセプトの登録・訂正";


                hnd = Win32API.GetTargetControl(winN, "ThunderRT6CommandButton", 19);

                hnd = Win32API.GetTargetControl(winN, "MaskEditX");
                rect = Win32API.GetRect(hnd.hWnd);

                runMouseCursorMoveOnLeftClick((rect.left + 3), ((rect.bottom + rect.top) / 2), 500);

                if (t.seikyu.Substring(0, 2) == "30" | t.seikyu.Substring(0, 2) == "31")
                {
                    gengo = "7";
                }

                foreach (var item in gengo)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(hnd.hWnd, code);
                }

                hnd = Win32API.GetTargetControl(winN, "DateEditX");

                rect = Win32API.GetRect(hnd.hWnd);

                runMouseCursorMoveOnLeftClick((rect.left + 3), ((rect.bottom + rect.top) / 2), 500);

                foreach (var item in t.seikyu)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(hnd.hWnd, code);
                }



                //2が「訂正・削除 3が　エラー修正
                hnd = Win32API.GetTargetControl(winN, "ThunderRT6CommandButton", 2);

                Win32API.ButtonClick(hnd).ConfigureAwait(false);
            }


            //★★★★★★★★ここでWindowが切り替わる★★★★★★★★★★
            winN = $"給付－{kakoTitle}レセプト修正";
            //★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★


            hnd = Win32API.GetTargetControl(winN, "TextEditX", 6);

            //Win32API.ButtonClick(hnd).ConfigureAwait(false);
            Win32API.SetText(hnd, winN, t.itirenNo);
            Thread.Sleep(500);

            hnd = Win32API.GetTargetControl(winN, "TextEditX", 2);
            Win32API.SetText(hnd, winN, t.kigo);
            Thread.Sleep(500);

            hnd = Win32API.GetTargetControl(winN, "TextEditX", 3);
            Win32API.SetText(hnd, winN, t.bango);
            Thread.Sleep(500);

            hnd = Win32API.GetTargetControl(winN, "MaskEditX", 4);
            rect = Win32API.GetRect(hnd.hWnd);
            Thread.Sleep(500);
            runMouseCursorMoveOnLeftClick(((rect.left + 5)), ((rect.bottom + rect.top) / 2), 500);
            foreach (var item in t.honke)
            {
                var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                Win32API.InputKey(hnd.hWnd, code);
            }
            //Thread.Sleep(500);

            hnd = Win32API.GetTargetControl(winN, "MaskEditX", 5);
            rect = Win32API.GetRect(hnd.hWnd);
            // Thread.Sleep(500);
            runMouseCursorMoveOnLeftClick(((rect.left + 5)), ((rect.bottom + rect.top) / 2), 500);
            foreach (var item in t.fuken)
            {
                var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                Win32API.InputKey(hnd.hWnd, code);
            }
            //   Thread.Sleep(500);

            hnd = Win32API.GetTargetControl(winN, "MaskEditX", 6);
            rect = Win32API.GetRect(hnd.hWnd);
            //  Thread.Sleep(500);
            runMouseCursorMoveOnLeftClick(((rect.left + 5)), ((rect.bottom + rect.top) / 2), 500);

            foreach (var item in t.tensuhyo)
            {
                var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                Win32API.InputKey(hnd.hWnd, code);
            }
            Thread.Sleep(800);

            hnd = Win32API.GetTargetControl(winN, "TextEditX", 8);
            rect = Win32API.GetRect(hnd.hWnd);
            runMouseCursorMoveOnLeftClick(((rect.left + 5)), ((rect.bottom + rect.top) / 2), 500);
            foreach (var item in t.iryoukikan)
            {
                var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                Win32API.InputKey(hnd.hWnd, code);
            }
            Thread.Sleep(800);

            //次へ
            hnd = Win32API.GetTargetControl(winN, "ThunderRT6CommandButton", 1);
            Win32API.ButtonClick(hnd).ConfigureAwait(false);

            winN = $"給付－{kakoTitle}レセプト訂正・削除 一覧";
            Thread.Sleep(2000);

            hnd = Win32API.GetTargetControl(winN, "SPRJ32X60_SpreadSheet");
            rect = Win32API.GetRect(hnd.hWnd);
            Thread.Sleep(500);
            runMouseCursorMoveOnLeftClick((rect.left + 10 + 620), (rect.top + 40 + 350), 500);


            Thread.Sleep(500);
            //画像同時表示
            runMouseCursorMoveOnLeftClick((rect.left + 690), (rect.top + 40), 500);

            Thread.Sleep(500);

            //////////////////選択のレセプトの照合/////////////////////////////////
            if (Validation)
            {

                //int rowCount = 1;
                //var info = new HenpuIrai().CopyTest(rowCount);
                //if (t.KetteiTensu != info.決定点数)
                //{
                //    MessageBox.Show($"{info.決定点数} {t.KetteiTensu}");
                //}

            }

            else if (!Validation)
            {
                runMouseCursorMoveOnLeftClick((rect.left + 10), (rect.top + 40), 500);

            }

            //レセプト選択して再審査画面へ


            //MessageBox.Show($"{t.KetteiTensu.ToString()} = {Tensu}");

            //var intTensu = int.Parse(Tensu);

            //if ( (t.KetteiTensu == intTensu.ToString()))
            //{
            //    //選択チェックボックス
            //    runMouseCursorMoveOnLeftClick((rect.left + 10), (rect.top + 40), 500);
            //}
            //else if ( (t.KetteiTensu != intTensu.ToString()))
            //{
            //    //画像表示
            //    Win32API.ObjectCenterClick(hnd, 30, 25 + rowInterval * rowCount);

            //    //３つとも一致しない
            //    var md = MessageBox.Show("選択した情報（１行目）が返付依頼書の情報と一致しない部分があります。このレセプトで返付依頼を行ってもよい場合は「はい」を、２行目以降に処理を遷移する場合は「いいえ」を押してください。", "", MessageBoxButtons.YesNoCancel);

            //    if (md == DialogResult.Yes)
            //    {
            //        //選択チェックボックス
            //        runMouseCursorMoveOnLeftClick((rect.left + 10), (rect.top + 40), 500);

            //    }
            //    else if (md == DialogResult.No)
            //    {
            //         var a = GetSelectedInfo(hnd, 2);
            //        MessageBox.Show("この処理をスキップします。（２行目）");
            //        return;
            //    }

            //    else if (md == DialogResult.Cancel)
            //    {
            //        MessageBox.Show("この処理をスキップします。");
            //        return;
            //    }
            //}

        }

        Info GetSelectedInfo(Window hnd, int rowCount)
        {
            //////////////////選択のレセプトの照合/////////////////////////////////
            const int rowInterval = 17;
            var info = new Info();

            Win32API.ObjectCenterClick(hnd, 690, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.決定点数 = Clipboard.GetText();

            Win32API.ObjectCenterClick(hnd, 400, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.診療年月 = Clipboard.GetText();

            Win32API.ObjectCenterClick(hnd, 243, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.整理番号 = Clipboard.GetText();

            Win32API.ObjectCenterClick(hnd, 460, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.受診者名 = Clipboard.GetText();

            Win32API.ObjectCenterClick(hnd, 300, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.記号 = Clipboard.GetText();

            Win32API.ObjectCenterClick(hnd, 334, 25 + rowInterval * rowCount);
            SendKeys.SendWait("^c");
            Thread.Sleep(100);
            info.番号 = Clipboard.GetText();
            ///////////////////////////////////////////////////////////////////////
            ///
            return info;
        }

        private void runMouseCursorMoveOnLeftClick(int x, int y, int sleepTime)
        {

            Console.WriteLine($"{x.ToString()}-{y.ToString()}");
            //マウスポインタの位置を画面左上基準の座標(x,y)にする
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(sleepTime);
        }

        public void SaisinsaTouroku(string seikyuDate, string mousideCode = "100018", string hosoku = "", string seiriNo = "", bool chosei = false)
        {

            var ChouseizumiFlag = chosei;

            var ConfirmWindow = "再審査－再審査等レセプト登録（紙レセプト請求分）";

            const string but = "ThunderRT6CommandButton";
            Thread.Sleep(1000);
            //レセ訂正削除画面の再審査登録ボタンを押す
            //var saisinsa = Win32API.GetTargetControl(TitleName, but).Click(1000).ConfigureAwait(false);
            var saisinsa = Win32API.GetTargetControl(TitleName, but);
            Win32API.ButtonClick(saisinsa);

            try
            {
                //string TitleName = "給付－レセプトエラー訂正・削除";

                Thread.Sleep(3000);

                //20210720　なぜか押せなくなった

                //Win32API.GetTargetControl(TitleName, "Button", "はい(&Y)").ObjectCenterClick(1500);

                Win32API.MouseLClick(660, 573);
                //レせ電だとここでダイアログボックスがでる。　紙レせはでない
                try
                {


                    //ボタンを物理的にクリックする一連処理
                    Thread.Sleep(4000);

                    Win32API.GetTargetControl(ConfirmWindow, "Button", "はい(&Y)").ObjectCenterClick(1500);

                    ConfirmWindow = "再審査－再審査等レセプト登録（オンライン請求分）";
                }
                catch
                {
                    throw new ArgumentNullException();

                }

                var bt = Win32API.GetTargetControl(ConfirmWindow, "MaskEditX", 3);
                Thread.Sleep(500);
                //rect = Win32API.GetRect(bt.hWnd);
                //Thread.Sleep(500);
                //runMouseCursorMoveOnLeftClick((rect.left + 5), ((rect.bottom + rect.top) / 2), 1000); 

                foreach (var item in mousideCode)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(bt.hWnd, code);
                }

                if (mousideCode == "100049")
                {
                    bt = Win32API.GetTargetControl(ConfirmWindow, "MaskEditX", 4);
                    Thread.Sleep(500);
                    var rect2 = Win32API.GetRect(bt.hWnd);
                    Thread.Sleep(500);
                    runMouseCursorMoveOnLeftClick((rect2.left + 5), ((rect2.bottom + rect2.top) / 2), 1000);

                    foreach (var item in "190001")
                    {
                        var code = new Win32API().GetNumCode(int.Parse(item.ToString()));


                        Win32API.InputKey(bt.hWnd, code);
                    }
                }

                Thread.Sleep(500);
                ConfirmWindow = "再審査－再審査等レセプト登録（オンライン請求分）";

                //補足
                var hnd = Win32API.GetTargetControl(ConfirmWindow, "TextEditX", 6);
                Thread.Sleep(500);
                Win32API.SetText(hnd, ConfirmWindow, hosoku);
                Thread.Sleep(500);

                hnd = Win32API.GetTargetControl(ConfirmWindow, "TextEditX", 2);
                Thread.Sleep(500);
                Win32API.SetText(hnd, ConfirmWindow, seiriNo);
                Thread.Sleep(500);

                //請求年月   
                hnd = Win32API.GetTargetControl(ConfirmWindow, "DateEditX", 2);
                Thread.Sleep(500);
                var rect = Win32API.GetRect(hnd.hWnd);

                hnd.ObjectLeftEndClick();

                foreach (var item in seikyuDate)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(hnd.hWnd, code);
                }
                Thread.Sleep(500);

                //連絡済みチェック
                if (ChouseizumiFlag)
                {

                    runMouseCursorMoveOnLeftClick(((rect.left + 226)), ((rect.bottom + rect.top) / 2) - 22, 800);
                }


                //実際はここを再審査ボタンにする indexは6
                Win32API.GetTargetControl(ConfirmWindow, "ThunderRT6CommandButton", 6).Click().ConfigureAwait(false);

                try
                {
                    Thread.Sleep(1000);
                    Win32API.GetTargetControl("再審査－再審査等レセプト登録（オンライン請求分）", "Button", "はい(&Y)").ObjectCenterClick();
                    Thread.Sleep(1000);

                }
                catch (Exception)
                {

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "再審査のクラスで");
            }
        }

        public void SendSaisinsa(string seikyuDate = "", string saisinsaDate = "", string mousideCode = "100018", string hosoku = "", string seiriNo = "", string ItirenNo = "", bool chosei = false, bool secondFlg = false)
        {
            try
            {
                var ChouseizumiFlag = chosei;


                var ConfirmWindow = "再審査－再審査等レセプト登録（紙レセプト請求分）";

                var ConfirmWindow2 = "再審査－再審査等レセプト登録（紙レセプト請求分）";

                if (secondFlg)
                {
                    ConfirmWindow2 = "再審査－再審査等レセプト登録（オンライン請求分）";
                }
                Thread.Sleep(4000);
                string TitleName = "給付－レセプトエラー訂正・削除";
                const string but = "ThunderRT6CommandButton";

                Win32API.MouseLClick(38, 101, 100);

                Win32API.MouseLClick(38, 132, 100);

                var obj = Win32API.GetTargetControl(ConfirmWindow2, "MaskEditX", 11);
                Win32API.InputKeyNessesaryMouseClick(obj, "9");

                obj = Win32API.GetTargetControl(ConfirmWindow2, "DateEditX", 9);
                Win32API.InputKeyNessesaryMouseClick(obj, seikyuDate);

                obj = Win32API.GetTargetControl(ConfirmWindow2, "TextEditX", 18);
                Win32API.InputKeyNessesaryMouseClick(obj, ItirenNo);

                //await Win32API.GetTargetControl(ConfirmWindow2, but, 3).Click();          


                var saisinsa = Win32API.GetTargetControl(ConfirmWindow2, but, 3);
                Win32API.ButtonClick(saisinsa);
                //Console.ReadKey();
                //レせ電だとここでダイアログボックスがでる。　紙レせはでない


                try
                {
                    //ボタンを物理的にクリックする一連処理
                    Thread.Sleep(15000);

                    Win32API.GetTargetControl(ConfirmWindow, "Button", "はい(&Y)").ObjectCenterClick(1500);

                    ConfirmWindow = "再審査－再審査等レセプト登録（オンライン請求分）";
                }
                catch
                {
                    Thread.Sleep(1000);
                    throw new ArgumentNullException();
                }


                var bt = Win32API.GetTargetControl(ConfirmWindow, "MaskEditX", 3);
                Thread.Sleep(500);

                foreach (var item in mousideCode)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(bt.hWnd, code);
                }

                Thread.Sleep(500);
                ConfirmWindow = "再審査－再審査等レセプト登録（オンライン請求分）";

                //補足
                var hnd = Win32API.GetTargetControl(ConfirmWindow, "TextEditX", 6);
                //Thread.Sleep(500);
                Win32API.SetText(hnd, ConfirmWindow, hosoku);
                //Thread.Sleep(500);

                hnd = Win32API.GetTargetControl(ConfirmWindow, "TextEditX", 2);
                //Thread.Sleep(500);
                Win32API.SetText(hnd, ConfirmWindow, seiriNo);
                Thread.Sleep(500);

                //請求年月
                hnd = Win32API.GetTargetControl(ConfirmWindow, "DateEditX", 2);
                Thread.Sleep(500);
                var rect = Win32API.GetRect(hnd.hWnd);

                hnd.ObjectLeftEndClick();

                foreach (var item in saisinsaDate)
                {
                    var code = new Win32API().GetNumCode(int.Parse(item.ToString()));

                    Win32API.InputKey(hnd.hWnd, code);
                }
                Thread.Sleep(500);

                //連絡済みチェック
                if (ChouseizumiFlag)
                {

                    Win32API.MouseLClick(((rect.left + 226)), ((rect.bottom + rect.top) / 2) - 22, 800);
                }


                //実際はここを再審査ボタンにする indexは6
                Win32API.GetTargetControl(ConfirmWindow, "ThunderRT6CommandButton", 6).Click().ConfigureAwait(false);

                try
                {
                    Thread.Sleep(3000);

                    Win32API.GetTargetControl("再審査－再審査等レセプト登録（オンライン請求分）", "Button", "はい(&Y)").ObjectCenterClick();
                    Thread.Sleep(1000);

                }
                catch (Exception)
                {

                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ClickNext()
        {

            var next = Win32API.GetTargetControl(TitleName, "ThunderRT6CommandButton", 2);
            Thread.Sleep(1000);
            Win32API.ButtonClick(next).ConfigureAwait(false);
            Thread.Sleep(1000);
        }

        public void SetZeroItibuFutankin()
        {
            var TitleName = "給付－レセプトエラー訂正・削除";
            Thread.Sleep(1000);

            var button = Win32API.GetTargetControl(TitleName, "NumEditX", 25);
            Thread.Sleep(1000);
            var rect = Win32API.GetRect(button.hWnd);
            // 選択した文字列を置き換える
            Thread.Sleep(500);
            runMouseCursorMoveOnLeftClick(((rect.left + rect.right) / 2), ((rect.bottom + rect.top) / 2), 1000);
            Thread.Sleep(500);


            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);
            Win32API.InputKey(button.hWnd, 0x2E);

            Win32API.InputKey(button.hWnd, 0x30);
            Thread.Sleep(1000);

            //////強制訂正にすること
            //var next = Win32API.GetTargetControl(TitleName, "ThunderRT6CommandButton", 2);
            //Console.ReadKey();
            //Thread.Sleep(500);
            //Win32API.ButtonClick(next).ConfigureAwait(false);
            //Thread.Sleep(500);

        }
    }
}
