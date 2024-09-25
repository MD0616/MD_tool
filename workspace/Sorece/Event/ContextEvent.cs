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
using System.Runtime.InteropServices;

namespace MD_Explorer
{
    public partial class MainForm : Form
    {
        // 削除ボタンがクリックされたときのイベントハンドラを追加
        private void eventDelete_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                List<string> itemNames = new List<string>();
                string absoluteDustBoxPath = Path.GetFullPath(GlobalSettings.dustBoxPath);

                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0];
                    string fullPath = Path.Combine(path, itemName);

                    if (!Directory.Exists(absoluteDustBoxPath))
                    {
                        Directory.CreateDirectory(absoluteDustBoxPath);
                    }

                    // `fullPath` が `dustbox` 内にあるかを確認
                    if (fullPath.StartsWith(absoluteDustBoxPath, StringComparison.OrdinalIgnoreCase))
                    {
                        var deleteResult = MessageBox.Show(string.Format("'{0}'を完全に削除してもよろしいですか？", itemName), "完全削除確認", MessageBoxButtons.YesNo);

                        if (deleteResult == DialogResult.Yes)
                        {
                            if (File.Exists(fullPath))
                            {
                                File.Delete(fullPath);
                            }
                            else if (Directory.Exists(fullPath))
                            {
                                Directory.Delete(fullPath, true);
                            }
                        }
                    }
                    else
                    {
                        // 削除確認ダイアログを表示
                        var result = MessageBox.Show(string.Format("'{0}'を削除してもよろしいですか？", itemName), "削除確認", MessageBoxButtons.YesNo);

                        if (result == DialogResult.Yes)
                        {
                            // ファイルまたはディレクトリを移動
                            string destinationPath = Path.Combine(absoluteDustBoxPath, itemName);

                            if (File.Exists(fullPath))
                            {
                                CommonLibrary.CopyFile(fullPath, destinationPath);
                                File.Delete(fullPath);
                            }
                            else if (Directory.Exists(fullPath))
                            {
                                // 移動先が `dustbox` 内で無限ループしないように、既存のフォルダを確認
                                if (!destinationPath.StartsWith(absoluteDustBoxPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    CommonLibrary.MoveDirectory(fullPath, destinationPath);
                                }
                            }
                        }
                    }
                }
            }
            // リストを更新する。
            RefreshActiveTab();
        }

        private void eventRename_Click(object sender, EventArgs e)
        {
            // 現在選択されているタブがある場合、そのタブの現在選択されているアイテムの名前を取得します。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    string fullPath = Path.Combine(path, itemName);

                    // 新しい名前を入力するためのダイアログを表示
                    string newFileName = Prompt.ShowDialog(string.Format("変更前'{0}'\n新しい名前を記載してください", itemName), "リネーム", itemName);

                    if (!string.IsNullOrEmpty(newFileName))
                    {
                        string newPath = Path.Combine(path, newFileName);

                        // ファイルまたはディレクトリの名前を変更
                        if (File.Exists(fullPath))
                        {
                            CommonLibrary.CopyFile(fullPath, newPath);
                            File.Delete(fullPath);
                        }
                        else if (Directory.Exists(fullPath))
                        {
                            CommonLibrary.MoveDirectory(fullPath, newPath);
                        }
                    }
                }
                // リストを更新する。
                RefreshActiveTab();
            }
        }

        private void eventCopyName_Click(object sender, EventArgs e)
        {
            // このメソッドは、名前コピーボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブの現在選択されているアイテムの名前を取得し、その名前をクリップボードにコピーします。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                List<string> itemNames = new List<string>();
                foreach (var item in listBox.SelectedItems)
                {
                    string selectedItem = item.ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    itemNames.Add(itemName);
                }
                Clipboard.SetText(string.Join(Environment.NewLine, itemNames)); // クリップボードにコピー
            }
        }

        private void eventCopyFullPath_Click(object sender, EventArgs e)
        {
            // このメソッドは、フルパスコピーボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブの現在選択されているアイテムのフルパスを取得し、そのフルパスをクリップボードにコピーします。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                List<string> fullPaths = new List<string>();
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    string fullPath = Path.Combine(path, itemName);
                    fullPaths.Add(fullPath);
                }
                Clipboard.SetText(string.Join(Environment.NewLine, fullPaths)); // クリップボードにコピー
            }
        }

        private void eventSafeOpen_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                List<string> itemNames = new List<string>();
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0];
                    string fullPath = Path.Combine(path, itemName);

                    if (!Directory.Exists(GlobalSettings.safeBoxPath))
                    {
                        Directory.CreateDirectory(GlobalSettings.safeBoxPath);
                    }

                    string absoluteSafeBoxPath = Path.Combine(GlobalSettings.safeBoxPath, itemName);
                    MessageBox.Show(string.Format("設定ファイル '{0}' {1} ",absoluteSafeBoxPath,fullPath));

                    if (File.Exists(fullPath))
                    {
                        CommonLibrary.CopyFile(fullPath, absoluteSafeBoxPath);
                        Process process = Process.Start(absoluteSafeBoxPath);
                        process.EnableRaisingEvents = true;
                        process.Exited += (s, args) =>
                        {
                            try
                            {
                                File.Delete(absoluteSafeBoxPath);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ファイルの削除中にエラーが発生しました: " + ex.Message);
                            }
                        };
                    }
                }
            }
        }

        private void eventVSCode_Click(object sender, EventArgs e)
        {
            // 現在選択されているアイテムをVSCodeで開きます。
            OpenSelectedItemWith("code");
        }

        private void eventDuplicate_Click(object sender, EventArgs e)
        {
            // 現在選択されているタブがある場合、そのタブの現在選択されているアイテムの名前を取得します。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    string fullPath = Path.Combine(path, itemName);

                    // 新しい名前を入力するためのダイアログを表示
                    string newFileName = Prompt.ShowDialog(string.Format("複製元'{0}'\n新しい名前を記載してください", itemName), "複製", itemName);

                    if (!string.IsNullOrEmpty(newFileName))
                    {
                        string newPath = Path.Combine(path, newFileName);

                        // ファイルまたはディレクトリを複製
                        if (File.Exists(fullPath))
                        {
                            // ファイルの複製処理
                            CommonLibrary.CopyFile(fullPath, newPath);
                        }
                        else if (Directory.Exists(fullPath))
                        {
                            // フォルダの複製処理（再帰的にコピー）
                            CopyDirectoryRecursive(fullPath, newPath);
                        }
                    }
                }
                // リストを更新する。
                RefreshActiveTab();
            }
        }

        // 再帰的にディレクトリをコピーするメソッド
        private void CopyDirectoryRecursive(string sourceDir, string targetDir)
        {
            // ターゲットディレクトリが存在しない場合は作成
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // ファイルをコピー
            foreach (string filePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(filePath);
                string targetFilePath = Path.Combine(targetDir, fileName);
                CommonLibrary.CopyFile(filePath, targetFilePath); // ファイルのコピー処理
            }

            // サブディレクトリを再帰的にコピー
            foreach (string dirPath in Directory.GetDirectories(sourceDir))
            {
                string dirName = Path.GetFileName(dirPath);
                string targetDirPath = Path.Combine(targetDir, dirName);
                CopyDirectoryRecursive(dirPath, targetDirPath);
            }
        }
    }
}