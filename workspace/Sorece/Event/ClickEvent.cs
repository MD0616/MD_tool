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
        private void button_Click(object sender, MouseEventArgs e)
        {
            Button button = sender as Button;
            TextBox textBox = sender as TextBox;

            if (textBox != null)
            {
                if (textBox.Name == "SearchBar")
                {
                    // btnSearchがクリックされたときの処理をここに書く
                    btnSearch_Click(null,e);
                }
            }
            else if (button != null)
            {
                if (button.Name == "Search")
                {
                    // btnSearchがクリックされたときの処理をここに書く
                    btnSearch_Click(null,e);
                }
                else if (button.Name == "Home")
                {
                    // btnHomeがクリックされたときの処理をここに書く
                    btnHome_Click(null,e);
                }
            }
        }

        private void contextMenuShortcut_Click(object sender, MouseEventArgs e)
        {
            var menuItem = contextMenu.GetItemAt(e.Location) as ToolStripMenuItem; // クリックされたメニューアイテムを取得
            if (menuItem != null)
            {
                var fullPath = menuItem.Tag.ToString();
                if (e.Button == MouseButtons.Left) // シングルクリック
                {
                    if (Directory.Exists(fullPath))
                    {
                        UpdateActiveTab(fullPath);
                    }
                }
                else if (e.Button == MouseButtons.Middle) // ホイールクリック
                {
                    if (Directory.Exists(fullPath))
                    {
                        OpenNewTab(fullPath);
                    }
                }
            }
        }

        // ショートカットボタンの位置にコンテキストメニューを表示
        private void btnShortcut_Click(object sender, EventArgs e)
        {
            contextMenu.Show(btnShortcut, new Point(0, btnShortcut.Height));
        }

        private void ExecuteScript(string scriptPath)
        {
            // 現在のアクティブタブのディレクトリでスクリプトを実行
            var psi = new ProcessStartInfo();
            psi.FileName = "powershell";
            psi.UseShellExecute = false;
            string directoryPath = GetCurrentTabDirectory();
            psi.WorkingDirectory = directoryPath;
            psi.Arguments = string.Format("-File \"{0}\"", scriptPath);

            if (Directory.Exists(directoryPath))
            {
                var result = MessageBox.Show(string.Format("スクリプト: {0}\nディレクトリ: {1}\n\nこれらの設定でスクリプトを実行しますか？", Path.GetFileName(scriptPath), directoryPath), "確認", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Process.Start(psi);
                }
            }
            else
            {
                MessageBox.Show("The directory or the script file does not exist.");
            }
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            if (scriptComboBox != null && scriptComboBox.SelectedItem != null)
            {
                var scriptInfo = (ScriptInfo)scriptComboBox.SelectedItem;
                var scriptPath = scriptInfo.FullPath;
                ExecuteScript(scriptPath);
            }
            else
            {
                MessageBox.Show("スクリプトが選択されていません");
            }
        }

        private void btnTerminal_MouseDown(object sender, MouseEventArgs e)
        {
            // このメソッドは、ターミナルボタンがクリックされたときに呼び出されます。
            // 中クリックが行われた場合、新しいターミナルを管理者権限で開きます。
            OpenTerminal(e.Button == MouseButtons.Middle);
        }

        private void listBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // このメソッドは、リストボックスのアイテムがダブルクリックされたときに呼び出されます。
            // 現在選択されているアイテムのパスを取得し、そのパスをHandleDoubleClickメソッドに渡します。
            ListBox listBox = (ListBox)sender;
            for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
            {
                string selectedPath = listBox.SelectedItems[i].ToString();
                string path = listBox.Tag.ToString();
                HandleDoubleClick(selectedPath, path, listBox, (TabPage)listBox.Parent);
            }
        }

        private void HandleDoubleClick(string selectedPath, string path, ListBox listBox, TabPage tabPage)
        {
            // このメソッドは、リストボックスのアイテムがダブルクリックされたときに呼び出されます。
            // 選択されたアイテムがディレクトリを指している場合、そのディレクトリを開きます。
            // 選択されたアイテムがファイルを指している場合、そのファイルをVSCodeで開きます。
            if (selectedPath == "ひとつ前に戻る")
            {
                DirectoryInfo parentDir = Directory.GetParent(path);
                if (parentDir != null)
                {
                    string parentPath = parentDir.FullName;
                    UpdateListBox(listBox, parentPath, tabPage);
                }
            }
            else
            {
                string[] parts = selectedPath.Split(new[] { ": " }, StringSplitOptions.None);
                if (parts.Length > 1)
                {
                    // アイテム名のみを取得
                    string itemName = parts[1].Split(new[] { " サイズ: " }, StringSplitOptions.None)[0];
                    itemName = itemName.Replace(" サイズ", "").TrimEnd();
                    // フルパスを取得
                    selectedPath = Path.Combine(path, itemName);
                    if (Directory.Exists(selectedPath))
                    {
                        UpdateListBox(listBox, selectedPath, tabPage);
                    }
                    else if (selectedPath.EndsWith(".lnk"))
                    {
                        // PowerShellを使用してショートカットのリンク先を取得
                        var psi = new ProcessStartInfo();
                        psi.FileName = "powershell";
                        psi.UseShellExecute = false;
                        psi.RedirectStandardOutput = true;
                        psi.Arguments = string.Format("-Command \"$sh = New-Object -COM WScript.Shell; $sc = $sh.CreateShortcut('{0}'); $sc.TargetPath\"", selectedPath);
                        var process = Process.Start(psi);
                        var output = process.StandardOutput.ReadToEnd();
                        process.WaitForExit();

                        string targetPath = output.Trim();
                        if (Directory.Exists(targetPath))
                        {
                            // リンク先がディレクトリの場合、そのディレクトリを開く
                            UpdateListBox(listBox, targetPath, tabPage);
                        }
                        else if (File.Exists(targetPath))
                        {
                            try
                            {
                                // リンク先がファイルの場合、そのファイルを開く
                                Process.Start(targetPath);
                            }
                            catch (System.ComponentModel.Win32Exception)
                            {
                                // 関連付けられたアプリケーションが存在しない場合、codeで開く。
                                Process.Start(GlobalSettings.fileOpenExe, targetPath);
                            }
                        }
                    }
                    else
                    {
                        Process.Start(selectedPath);
                    }
                }
            }
        }

        private void listBox_MouseDown(object sender, MouseEventArgs e)
        {
            // このメソッドは、リストボックスのアイテムがマウスでクリックされたときに呼び出されます。
            ListBox listBox = (ListBox)sender;
            if (e.Button == MouseButtons.Right)
            {
                // 右クリックされたとき、dropDownMenuを表示します。
                dropDownMenu.Show(listBox, new Point(e.X, e.Y));

                // コンテキストメニューにフォーカスを設定します。
                dropDownMenu.Focus();

                // 最初の項目を選択します。
                if (dropDownMenu.Items.Count > 0)
                {
                    ((ToolStripMenuItem)dropDownMenu.Items[0]).Select();
                }
            }
            else if (e.Button == MouseButtons.Middle)
            {
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string selectedPath = listBox.SelectedItems[i].ToString();
                    string path = listBox.Tag.ToString();
                    if (selectedPath == "ひとつ前に戻る")
                    {
                        DirectoryInfo parentDir = Directory.GetParent(path);
                        if (parentDir != null)
                        {
                            string parentPath = parentDir.FullName;
                            OpenNewTab(parentPath);
                        }
                    }
                    else
                    {
                        string[] parts = selectedPath.Split(new[] { ": " }, StringSplitOptions.None);
                        if (parts.Length > 1)
                        {
                            // アイテム名のみを取得
                            string itemName = parts[1].Split(new[] { " サイズ: " }, StringSplitOptions.None)[0];
                            itemName = itemName.Replace(" サイズ", "").TrimEnd();
                            // フルパスを取得
                            selectedPath = Path.Combine(path, itemName);
                            if (Directory.Exists(selectedPath))
                            {
                                OpenNewTab(selectedPath);
                            }
                            else if (selectedPath.EndsWith(".lnk"))
                            {
                                // PowerShellを使用してショートカットのリンク先を取得
                                var psi = new ProcessStartInfo();
                                psi.FileName = "powershell";
                                psi.UseShellExecute = false;
                                psi.RedirectStandardOutput = true;
                                psi.Arguments = string.Format("-Command \"$sh = New-Object -COM WScript.Shell; $sc = $sh.CreateShortcut('{0}'); $sc.TargetPath\"", selectedPath);
                                var process = Process.Start(psi);
                                var output = process.StandardOutput.ReadToEnd();
                                process.WaitForExit();

                                string targetPath = output.Trim();
                                if (Directory.Exists(targetPath))
                                {
                                    // リンク先がディレクトリの場合、そのディレクトリを開く
                                    OpenNewTab(targetPath);
                                }
                                else if (File.Exists(targetPath))
                                {
                                    try
                                    {
                                        // リンク先がファイルの場合、そのファイルを開く
                                        Process.Start(targetPath);
                                    }
                                    catch (System.ComponentModel.Win32Exception)
                                    {
                                        // 関連付けられたアプリケーションが存在しない場合、codeで開く。
                                        Process.Start(GlobalSettings.fileOpenExe, targetPath);
                                    }
                                }
                            }
                            else
                            {
                                Process.Start(selectedPath);
                            }
                        }
                    }
                }
            }
        // 選択個数を表示するラベルを更新
        labelSelectedCount.Text = String.Format("選択された項目数: {0}", listBox.SelectedItems.Count);
        }
        private void BtnCopyMove_Click(object sender, EventArgs e)
        {
            using (CopyMoveForm copyMoveForm = new CopyMoveForm())
            {
                copyMoveForm.ShowDialog();
            }
        }
    }
}