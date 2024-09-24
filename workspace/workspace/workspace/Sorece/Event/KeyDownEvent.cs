using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Management.Automation;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.ComponentModel;

namespace MD_Explorer
{
    public partial class MainForm : Form
    {
        private void btn_KeyDown(object sender, KeyEventArgs e)
        {
            Button button = sender as Button; // senderをButtonにキャスト
            TextBox textBox = sender as TextBox;

            if (textBox != null)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    // エンターキーが押されたときにボタンのクリックメソッドを呼び出す
                    button_Click(textBox, new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0));
                }
                else if (e.Control && e.KeyCode == Keys.Enter)
                {
                    // エンターキーが押されたときにボタンのクリックメソッドを呼び出す
                    button_Click(textBox, new MouseEventArgs(MouseButtons.Middle, 1, 0, 0, 0));
                }
            }
            else if (button != null) // キャストが成功した場合
            {
                if (e.Control && e.KeyCode == Keys.Enter)
                {
                    // エンターキーが押されたときにボタンのクリックメソッドを呼び出す
                    button_Click(button, new MouseEventArgs(MouseButtons.Middle, 1, 0, 0, 0));
                }
            }
        }

        private void HandleKeyDown(object sender, KeyEventArgs e)
        {
            ListBox listBox = (ListBox)sender;
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.K)
            {
                // 上矢印キーが押されたときの処理
                if (listBox.SelectedIndex > 0)
                {
                    int newIndex = listBox.SelectedIndex - 1;
                    listBox.ClearSelected();  // すべての選択を解除
                    listBox.SelectedIndex = newIndex;  // 新しい項目を選択
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.J)
            {
                // 下矢印キーが押されたときの処理
                if (listBox.SelectedIndex < listBox.Items.Count - 1)
                {
                    int newIndex = listBox.SelectedIndex + 1;
                    listBox.ClearSelected();  // すべての選択を解除
                    listBox.SelectedIndex = newIndex;  // 新しい項目を選択
                }
                e.Handled = true;
            }
            else if ((e.Alt && e.KeyCode == Keys.H))
            {
                // ALT+Hキーが押されたときの処理
                if (tabControl1.SelectedIndex > 0) // タブが最初でなければ
                {
                    tabControl1.SelectedIndex--; // タブを一つ左に切り替える
                }
                e.Handled = true;
            }
            else if (e.Alt && e.KeyCode == Keys.L)
            {
                // ALT+Lキーが押されたときの処理
                if (tabControl1.SelectedIndex < tabControl1.TabPages.Count - 1) // タブが最後でなければ
                {
                    tabControl1.SelectedIndex++; // タブを一つ右に切り替える
                }
                e.Handled = true;
            }
            else if (e.Shift && e.KeyCode == Keys.X)
            {
                eventCopyName_Click(null, null);
            }
            else if (e.Shift && e.KeyCode == Keys.C)
            {
                eventCopyFullPath_Click(null, null);
            }
            else if (e.KeyCode == Keys.Enter)
            {
                // Enterキーが押されたときの処理
                if (listBox.SelectedItem != null)
                {
                    string selectedPath = listBox.SelectedItem.ToString();
                    string path = listBox.Tag.ToString();
                    TabPage tabPage = (TabPage)listBox.Parent;
                    HandleDoubleClick(selectedPath, path, listBox, tabPage);
                }
            }
        // 選択個数を表示するラベルを更新
        labelSelectedCount.Text = String.Format("選択された項目数: {0}", listBox.SelectedItems.Count);
        }
        private void listBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            ListBox listBox = (ListBox)sender;
            // 'D' キーの入力を無視します。
            if (e.KeyChar == 'D' || e.KeyChar == 'd')
            {
                e.Handled = true;
            }
            else if (e.KeyChar == 'F' || e.KeyChar == 'f')
            {
                e.Handled = true;
            }
            else if (e.KeyChar == 'L' || e.KeyChar == 'l')
            {
                dropDownMenu.Show(listBox, new Point(0,0));
                // コンテキストメニューにフォーカスを設定します。
                dropDownMenu.Focus();

                // 最初の項目を選択します。
                if (dropDownMenu.Items.Count > 0)
                {
                    ((ToolStripMenuItem)dropDownMenu.Items[0]).Select();
                }
                e.Handled = true;
            }
        }

        private void dropDownMenu_KeyDown(object sender, KeyEventArgs e)
        {
            var dropDownMenu = (ToolStripDropDown)sender;
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.K)
            {
                // 上矢印キーが押されたときの処理
                // 選択を一つ上の項目に移動します。
                MoveSelection(dropDownMenu, -1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.J)
            {
                // 下矢印キーが押されたときの処理
                // 選択を一つ下の項目に移動します。
                MoveSelection(dropDownMenu, 1);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.H)
            {
                // Hキーが押されたときの処理
                // ドロップダウンメニューを閉じます。
                dropDownMenu.Close();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                // Enterキーが押されたときの処理
                // 現在選択されている項目をクリックします。
                PerformClickOnSelectedItem(dropDownMenu);
                e.Handled = true;
            }
        }

        private void TxtPowerShellInput_KeyDown(object sender, KeyEventArgs e)
        {
            // このメソッドは、PowerShell入力テキストボックスでキーが押されたときに呼び出されます。
            // エンターキーが押された場合、テキストボックスの内容を取得し、PowerShellプロセスの標準入力に書き込みます。
            // その後、テキストボックスの内容をクリアします。
            if (e.KeyCode == Keys.Enter)
            {
                string input = txtPowerShellInput.Text.Trim();
                if (input.ToLower() == "home")
                {
                    // 「home」が入力されたときにディレクトリを変更する
                    string home = GetCurrentTabDirectory();
                    powerShellProcess.StandardInput.WriteLine("cd \'" + home + "\';Write-Host pwd \'" + home +"\'"); 
                }
                else
                {
                    // それ以外の入力はそのままパワーシェルに送る
                    powerShellProcess.StandardInput.WriteLine(input);
                }
                txtPowerShellInput.Clear();
            }
        }
    }
}