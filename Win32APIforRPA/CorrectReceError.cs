using KiiS.dbContext;
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Win32APIforRPA
{
    class CorrectReceError

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

        private void runMouseCursorMoveOnLeftClick(int x, int y, int sleepTime)
        {

            Console.WriteLine($"{x.ToString()}-{y.ToString()}");
            //マウスポインタの位置を画面左上基準の座標(x,y)にする
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(sleepTime);
        }


        public async Task<string> Execute(IProgress<string[]> p, int skipCount = 1)
        {

            try
            {
                string[] pgdata = { "", "" };

                var dResult = MessageBox.Show($"レセプトエラーリストの自働訂正を実行します。以下の項目を確認してから実行してください。\r\n\r\n" +
                    $"（１）KOSMOはメインウィンドウとエラー訂正ウィンドウ（処理する先頭の訂正画面を開いた状態にしてください）のみ開いてください\r\n" +
                    $"（２）「画像同時表示」オプションのチェックは外してください\r\n" +
                    $"（３）実行中はWindowsの操作は控えてください\r\n" +
                    $"\r\n　以上、準備ができましたらＯＫボタンをおしてください。", "実行前確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);


                if (dResult == DialogResult.Cancel) { MessageBox.Show("実行をキャンセルしました。"); return ""; }

                int skipNumber = skipCount - 1;
                int SleetTime = 2000;

                string WindowTitle = "給付－レセプトエラー訂正・削除";
                string ButtonClass = "ThunderRT6CommandButton";

                var data = //訂正するエラーデータのコレクション

                foreach (var item in data)
                {

                    var pMessage = "";


                    //await Task.Delay(2000);
                    Thread.Sleep(SleetTime);
                    pMessage = $"【{item.エラー内容}(整理番号：{item.整理ＮＯ})】{item.記号} - {item.番号} {item.氏名_漢字}を自動処理中です......";

                    var SeiriNo = Win32API.GetText(Win32API.GetTargetControl(WindowTitle, "TextEditX", 2));


                    //現在表示されている整理番号と処理中の整理番号をチェックする
                    if (SeiriNo != item.整理ＮＯ)
                    {
                        var mes = $"処理対象の整理番号 :【{item.エラー内容}(整理番号：{item.整理ＮＯ})】{item.記号} - {item.番号} {item.氏名_漢字}" +
                            $"\r\n" +
                            $"表示中の整理番号：{SeiriNo}" +
                            $"処理を続ける場合は「はい」を処理を中断する場合は「いいえ」を選択してください。" +
                            $"\r\n" +
                            $"（検索番号重複エラーは「はい」を選択してください）";

                        var result = MessageBox.Show(mes, "処理を続けてもよいですか", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.No)
                        {
                            MessageBox.Show("処理を中断しました。");
                            return "エラーが発生しました";
                        }
                    }

                    pgdata = new string[] { "", pMessage };
                    p.Report(pgdata);

                    if (item.エラー内容 == "資格無しＥＲＲ")
                    {
                        if (item.再審査申出理由コード == "エラーなし")
                        {


                            //★強制訂正
                            var bt = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 16);
                            Win32API.ButtonClick(bt).ConfigureAwait(false);



                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　エラーなしのため強制訂正しました";
                        }
                        else if (item.再審査申出理由コード.Contains("返納金"))
                        {

                            //★エラーとして訂正
                            var bt = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 15);
                            Win32API.ButtonClick(bt).ConfigureAwait(false);

                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　喪失後受診が含まれているため、返納金（疑）としてエラーとして訂正しました";

                        }
                        else if (item.再審査申出理由コード.Contains("再審査対象"))
                        {

                            if (item.再審査申出理由コード.Contains("100018"))
                            {
                                var c = new RunSaisinsaEntry();
                                c.TitleName = "給付－レセプトエラー訂正・削除";
                                c.SaisinsaTouroku("");

                                pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　申出理由コード 100018 で再審査登録を行いました。";

                            }
                            else if (item.再審査申出理由コード.Contains("100016"))
                            {
                                var c = new RunSaisinsaEntry();
                                c.TitleName = "給付－レセプトエラー訂正・削除";
                                c.SaisinsaTouroku("", "100016", "資格取得（認定）日以前の受診分です。");

                                pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　申出理由コード 100016 で再審査登録を行いました。";

                            }
                            else
                            {

                                var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                                Win32API.ButtonClick(next).ConfigureAwait(false);
                                pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　確定喪失後受診（医療機関証未確認）です。再審査登録をしてください";
                            }

                        }
                        else
                        {
                            var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                            Win32API.ButtonClick(next).ConfigureAwait(false);
                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　紙レセプト等のため資格確認ができませんでした、目視で訂正ください";

                        }
                    }


                    else if (item.エラー内容.Contains("重複ＥＲＲ"))
                    {
                        if (item.再審査申出理由コード == "エラーなし")
                        {

                            //★強制訂正
                            var bt = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 16);
                            Win32API.ButtonClick(bt).ConfigureAwait(false);

                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　重複請求ではないため強制訂正しました";
                        }
                        else if (item.再審査申出理由コード == "エラー判定不可能" | string.IsNullOrEmpty(item.再審査申出理由コード))
                        {
                            var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                            Win32API.ButtonClick(next).ConfigureAwait(false);
                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　エラーが判別できないためスキップします。";

                        }
                        else if (item.再審査申出理由コード == "100019")
                        {
                            var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                            Win32API.ButtonClick(next).ConfigureAwait(false);

                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　再審査対象（100019）です。";

                        }
                        else
                        {
                            var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                            Win32API.ButtonClick(next).ConfigureAwait(false);
                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　エラーが判別できないためスキップします。";

                        }
                    }

                    else if (item.エラー内容 == "特定疾患等一部負担額設定エラー" | item.エラー内容 == "公費併用の負担金入力です")
                    {
                        new RunSaisinsaEntry().SetZeroItibuFutankin();

                        var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 16);
                        Win32API.ButtonClick(next).ConfigureAwait(false);

                        pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　社保　一部負担金を０円に設定しました";



                    }

                    else if (item.エラー内容 == "検索番号が重複しています")
                    {
                        pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　　処理不要のためスキップします";

                    }

                    else if (item.エラー内容 == "未登録ＥＲＲ")
                    {

                        var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                        Win32API.ButtonClick(next).ConfigureAwait(false);
                        pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　　加入者の特定ができないためスキップします、目視訂正してください";

                    }

                    else if (item.エラー内容 == "氏名不照合" | item.エラー内容 == "生年月日不照合")
                    {
                        try
                        {

                            pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　訂正が不安定のため処理をスキップします。手動でエラー訂正をお願いいたします。";
                            await Task.Delay(4000);

                            var bt = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "Button", "OK");
                            Thread.Sleep(500);

                            var rect = Win32API.GetRect(bt.hWnd);
                            Thread.Sleep(500);

                            runMouseCursorMoveOnLeftClick(((rect.left + rect.right) / 2), ((rect.bottom + rect.top) / 2), 1000);

                            Thread.Sleep(1000);
                            var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                            Win32API.ButtonClick(next).ConfigureAwait(false);


                        }
                        catch (InvalidOperationException)
                        {

                        }

                    }

                    else //どれにも該当しないエラー
                    {
                        var next = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "ThunderRT6CommandButton", 2);
                        Win32API.ButtonClick(next).ConfigureAwait(false);
                        pMessage = $"【{item.エラー内容}】{item.記号} - {item.番号} {item.氏名_漢字}　⇒　エラー判別不能のためスキップします";

                    }

                    Thread.Sleep(SleetTime + 1000);

                    try
                    {
                        var bt = Win32API.GetTargetControl("給付－レセプトエラー訂正・削除", "Button", "OK");
                        Thread.Sleep(500);

                        var rect = Win32API.GetRect(bt.hWnd);
                        Thread.Sleep(500);

                        runMouseCursorMoveOnLeftClick(((rect.left + rect.right) / 2), ((rect.bottom + rect.top) / 2), 1000);

                        Thread.Sleep(1000);

                    }
                    catch (InvalidOperationException)
                    {

                    }

                    pgdata = new string[] { "", pMessage };
                    p.Report(pgdata);

                }

                pgdata = new string[] { "", "レセプトエラーの自動訂正は正常に完了しました。" };
                p.Report(pgdata);
                Thread.Sleep(SleetTime);


                return "";
            }
            catch (Exception ex)
            {

                string[] pgdata = { "", $"エラーが発生しました。自動処理を停止します。\r\n {ex.Message}" };
                p.Report(pgdata);
                await Task.Delay(2000);
                return "";
            }
        }
    }
}
