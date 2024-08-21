using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace MD_Explorer
{
    public partial class MainForm : Form
    {
        private void listBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            ListBox listBox = (ListBox)sender;
            if (e.Index >= 0 && e.Index < listBox.Items.Count)
            {
                e.DrawBackground();
                string item = listBox.Items[e.Index].ToString();
                Brush brush = GetBrushForItem(item);
                e.Graphics.DrawString(item, e.Font, brush, e.Bounds);
                e.DrawFocusRectangle();
            }
        }

        private Brush GetBrushForItem(string item)
        {
            if (item.StartsWith("[Dir]: "))
            {
                return Brushes.Red;
            }
            else if (item.StartsWith("[File]: "))
            {
                return Brushes.Gray;
            }
            else if (item.StartsWith("[Link]: "))
            {
                return Brushes.Green;
            }
            else
            {
                return Brushes.White;
            }
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            // タブの背景色を設定
            Color[] tabColors = { Color.Black, Color.Red, Color.Blue, Color.Green, Color.Brown };
            TabData tabData = (TabData)tabControl1.TabPages[e.Index].Tag; // 修正箇所
            int networkIndex = tabData.NetworkIndex; // 修正箇所
            e.Graphics.FillRectangle(new SolidBrush(tabColors[networkIndex]), e.Bounds);

            // タブのテキストを描画
            string tabText = tabControl1.TabPages[e.Index].Text;
            SizeF textSize = e.Graphics.MeasureString(tabText, e.Font);
            e.Graphics.DrawString(tabText, e.Font, Brushes.White, e.Bounds.Left + 5, e.Bounds.Top + 4);

            // タブの縁を黒に設定
            e.Graphics.DrawRectangle(Pens.Black, e.Bounds);

            // 太字のフォントを作成
            Font boldFont = new Font(e.Font, FontStyle.Bold);

            // タブの閉じるボタンの表示領域
            int closeButtonWidth = 12;
            int closeButtonHeight = 12;
            int closeButtonMargin = 5; // タブ名と閉じるボタンの間のマージン
            int closeButtonX = e.Bounds.Right - closeButtonWidth - closeButtonMargin;
            if (closeButtonX < e.Bounds.Left + textSize.Width + closeButtonMargin)
            {
                closeButtonX = e.Bounds.Left + (int)textSize.Width + closeButtonMargin;
            }
            Rectangle closeButtonRect = new Rectangle(closeButtonX, e.Bounds.Top + 4, closeButtonWidth, closeButtonHeight);
            e.Graphics.DrawString("X", boldFont, Brushes.Red, closeButtonRect); // 閉じるボタンを描画

            tabData.CloseButton = closeButtonRect; // 閉じるボタンの矩形を更新
        }

        private void tabControl1_MouseDown(object sender, MouseEventArgs e)
        {
            for (int i = 0; i < tabControl1.TabPages.Count; i++) // 各タブをチェック
            {
                Rectangle r = tabControl1.GetTabRect(i); // タブの矩形を取得
                TabData tabData = (TabData)tabControl1.TabPages[i].Tag; // 修正箇所
                Rectangle closeButton = tabData.CloseButton; // 閉じるボタンの矩形を取得
                if (closeButton.Contains(e.Location)) // マウスクリックが閉じるボタン内か確認
                {
                    tabControl1.TabPages.RemoveAt(i); // タブを削除
                    break;
                }
            }
        }
    }
}