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
        // 削除ボタンがクリックされたときのイベントハンドラを追加
        private void eventDelete_Click(object sender, EventArgs e)
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

                    if (!Directory.Exists(GlobalSettings.dustBoxPath))
                    {
                        Directory.CreateDirectory(GlobalSettings.dustBoxPath);
                    }

                    string absoluteDustBoxPath = Path.GetFullPath(GlobalSettings.dustBoxPath);
                    
                    // 移動したアイテムが指定ディレクトリ内のものであれば完全削除
                    if (absoluteDustBoxPath.Equals(path, StringComparison.OrdinalIgnoreCase))
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
                            if (File.Exists(fullPath))
                            {
                                CommonLibrary.CopyFile(fullPath, Path.Combine(absoluteDustBoxPath, itemName));
                                File.Delete(fullPath);
                            }
                            else if (Directory.Exists(fullPath))
                            {
                                CommonLibrary.MoveDirectory(fullPath, Path.Combine(absoluteDustBoxPath, itemName));
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
    }
}