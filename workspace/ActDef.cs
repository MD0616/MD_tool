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
        private bool HasAccessPermission(string path)
        {
            // このメソッドは、指定されたパスに対するアクセス権限があるかどうかを確認します。
            // ディレクトリの内容を取得できる場合は、アクセス権限があると判断し、trueを返します。
            // UnauthorizedAccessExceptionがスローされた場合、またはその他の例外がスローされた場合は、アクセス権限がないと判断し、falseを返します。
            try
            {
                var directories = Directory.GetDirectories(path);
                var files = Directory.GetFiles(path);
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void dropDownMenu_Opening(object sender, CancelEventArgs e)
        {
            // コンテキストメニューが開かれたとき、フォーカスを設定します。
            dropDownMenu.Focus();
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

        private void MoveSelection(ToolStripDropDown menu, int direction)
        {
            int maxIndex = menu.Items.Count - 1;
            for (int i = 0; i <= maxIndex; i++)
            {
                if (menu.Items[i].Selected)
                {
                    int newIndex = i + direction;
                    if (newIndex < 0) newIndex = 0;
                    if (newIndex > maxIndex) newIndex = maxIndex;
                    menu.Items[newIndex].Select();
                    break;
                }
            }
        }

        private void PerformClickOnSelectedItem(ToolStripDropDown menu)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                if (item.Selected && item.Enabled)
                {
                    ((ToolStripMenuItem)item).PerformClick();
                    break;
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
        private string GetCurrentTabDirectory()
        {
            // 現在のアクティブタブのディレクトリを取得
            var tabData = (TabData)tabControl1.SelectedTab.Tag;
            string directoryPath = tabData.Path;
            return directoryPath;
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            var scriptInfo = (ScriptInfo)scriptComboBox.SelectedItem;
            var scriptPath = scriptInfo.FullPath;
            ExecuteScript(scriptPath);
        }

        // 削除ボタンがクリックされたときのイベントハンドラを追加
        private void btnDelete_Click(object sender, EventArgs e)
        {
            // 現在選択されているタブがある場合、そのタブの現在選択されているアイテムの名前を取得し、その名前をクリップボードにコピーします。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                List<string> itemNames = new List<string>();
                for (int i = listBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItems[i].ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    string fullPath = Path.Combine(path, itemName);
                    // 削除確認ダイアログを表示
                    var result = MessageBox.Show(string.Format("'{0}'を削除してもよろしいですか？", itemName), "削除確認", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        // ファイルまたはディレクトリを削除
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
            }
            // リストを更新する。
            RefreshActiveTab();
        }

        private void btnRename_Click(object sender, EventArgs e)
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
                    string newFileName= Prompt.ShowDialog(string.Format("変更前'{0}'\n新しい名前を記載してください",itemName), "リネーム");

                    if (!string.IsNullOrEmpty(newFileName))
                    {
                        string newPath = Path.Combine(path, newFileName);

                        // ファイルまたはディレクトリの名前を変更
                        if (File.Exists(fullPath))
                        {
                            File.Move(fullPath, newPath);
                        }
                        else if (Directory.Exists(fullPath))
                        {
                            Directory.Move(fullPath, newPath);
                        }
                    }
                }
                // リストを更新する。
                RefreshActiveTab();
            }
        }

        private void btnSearch_Click(object sender, MouseEventArgs e)
        {
            // このメソッドは、検索ボタンがクリックされたときに呼び出されます。
            // テキストボックスからパスを取得し、そのパスがディレクトリを指しているか、ファイルを指しているかを確認します。
            // パスがディレクトリを指している場合、左クリックで現在のタブを更新し、中クリックで新しいタブを開きます。
            // パスがファイルを指している場合、そのファイルを開きます。
            // パスが存在しない場合、エラーメッセージを表示します。
            string path = txtSearchBar.Text;

            if (Directory.Exists(path))
            {
                if (e.Button == MouseButtons.Left)
                {
                    if (tabControl1.SelectedTab != null)
                    {
                        UpdateActiveTab(path);
                    }
                    else
                    {
                        OpenNewTab(path);
                    }
                }
                else if (e.Button == MouseButtons.Middle)
                {
                    OpenNewTab(path);
                }
                txtSearchBar.Text = string.Empty;
            }
            else if (File.Exists(path))
            {
                try
                {
                    // ファイルが存在する場合、そのファイルを開く
                    Process.Start(path);
                    txtSearchBar.Text = string.Empty;
                }
                catch (System.ComponentModel.Win32Exception)
                {
                    // 関連付けられたアプリケーションが存在しない場合、codeで開く。
                    Process.Start(fileOpenExe, path);
                    txtSearchBar.Text = string.Empty;
                }
            }
            else
            {
                MessageBox.Show("指定されたパスが存在しません。");
                txtSearchBar.Text = string.Empty;
            }
        }

        private void btnHome_Click(object sender, MouseEventArgs e)
        {
            // 左クリックで現在のタブをホームディレクトリに更新し、中クリックでホームディレクトリを開く新しいタブを開きます。
            if (e.Button == MouseButtons.Left)
            {
                if (tabControl1.TabPages.Count == 0)
                {
                    OpenNewTab(homeDirectory);
                }
                else
                {
                    UpdateActiveTab(homeDirectory);
                }
            }
            else if (e.Button == MouseButtons.Middle)
            {
                OpenNewTab(homeDirectory);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            // このメソッドは、更新ボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブを更新します。
            if (tabControl1.SelectedTab != null)
            {
                RefreshActiveTab();
            }
        }

        private void btnVSCode_Click(object sender, EventArgs e)
        {
            // 現在選択されているアイテムをVSCodeで開きます。
            OpenSelectedItemWith("code");
        }

        private void btnExplorer_Click(object sender, EventArgs e)
        {
            // このメソッドは、エクスプローラボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブの現在のパスを取得し、そのパスをエクスプローラで開きます。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString();
                    Process.Start("explorer.exe", path);
                }
            }
        }

        private void btnCopyPath_Click(object sender, EventArgs e)
        {
            // このメソッドは、パスコピーボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブの現在のパスを取得し、そのパスをクリップボードにコピーします。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString();
                    Clipboard.SetText(path);
                }
            }
        }

        private void btnCopyName_Click(object sender, EventArgs e)
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

        private void btnCopyFullPath_Click(object sender, EventArgs e)
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

        private void btnNewFile_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString(); // 現在のパスを取得
                    string fileName = Prompt.ShowDialog("新規ファイル名を入力してください", "新規ファイル作成");
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        string fullPath = Path.Combine(path, fileName);
                        File.Create(fullPath).Close();
                        UpdateListBox(listBox, path, activeTab);
                    }
                }
            }
        }

        private void btnNewFolder_Click(object sender, EventArgs e)
        {
            // このメソッドは、新規フォルダ作成ボタンがクリックされたときに呼び出されます。
            // 現在選択されているタブがある場合、そのタブの現在のパスを取得し、新規フォルダ名をユーザーに入力させます。
            // ユーザーが有効なフォルダ名を入力した場合、その名前の新しいフォルダを作成し、リストボックスを更新します。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString(); // 現在のパスを取得
                    string folderName = Prompt.ShowDialog("新規フォルダ名を入力してください", "新規フォルダ作成");
                    if (!string.IsNullOrEmpty(folderName))
                    {
                        string fullPath = Path.Combine(path, folderName);
                        Directory.CreateDirectory(fullPath);
                        UpdateListBox(listBox, path, activeTab);
                    }
                }
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
                    powerShellProcess.StandardInput.WriteLine("cd " + home + ";Write-Host pwd " + home ); 
                }
                else
                {
                    // それ以外の入力はそのままパワーシェルに送る
                    powerShellProcess.StandardInput.WriteLine(input);
                }
                txtPowerShellInput.Clear();
            }
        }

        private void btnTerminal_MouseDown(object sender, MouseEventArgs e)
        {
            // このメソッドは、ターミナルボタンがクリックされたときに呼び出されます。
            // 中クリックが行われた場合、新しいターミナルを管理者権限で開きます。
            OpenTerminal(e.Button == MouseButtons.Middle);
        }

        private void UpdateActiveTab(string path)
        {
            // このメソッドは、現在アクティブなタブを更新します。
            // 指定されたパスでリストボックスを更新します。
            TabPage activeTab = tabControl1.SelectedTab;
            ListBox listBox = (ListBox)activeTab.Controls[0];
            UpdateListBox(listBox, path, activeTab);
        }

        private void RefreshActiveTab()
        {
            // このメソッドは、現在アクティブなタブを更新します。
            // リストボックスの現在のパスでリストボックスを更新します。
            TabPage activeTab = tabControl1.SelectedTab;
            ListBox listBox = (ListBox)activeTab.Controls[0];
            if (listBox.Tag != null)
            {
                string path = listBox.Tag.ToString();
                UpdateListBox(listBox, path, activeTab);
            }
        }

        private void OpenSelectedItemWith(string application)
        {
            // このメソッドは、指定されたアプリケーションで現在選択されているアイテムを開きます。
            // 現在選択されているアイテムのパスを取得し、そのパスを指定されたアプリケーションで開きます。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.SelectedItem != null)
                {
                    string path = listBox.Tag.ToString();
                    string selectedItem = listBox.SelectedItem.ToString().Split(new[] { ": " }, StringSplitOptions.None)[1];
                    string itemName = selectedItem.Split(new[] { "  " }, StringSplitOptions.None)[0]; // アイテム名のみを取得
                    string fullPath = Path.Combine(path, itemName);
                    Process.Start(application, fullPath);
                }
            }
        }

        private void OpenTerminal(bool runAsAdmin)
        {
            // このメソッドは、新しいターミナルを開きます。
            // 現在選択されているタブの現在のパスを取得し、そのパスで新しいPowerShellターミナルを開きます。
            // runAsAdminパラメータがtrueの場合、ターミナルは管理者権限で開かれます。
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString();
                    string arguments = String.Format("-NoExit -Command Set-Location -Path '{0}'", path);
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = arguments,
                        Verb = runAsAdmin ? "runas" : ""
                    };
                    Process.Start(startInfo);
                }
            }
        }

        private void UpdateListBox(ListBox listBox, string path, TabPage tabPage)
        {
            // このメソッドは、指定されたパスでリストボックスを更新します。
            // パスのアクセス権限を確認し、ディレクトリ、ファイル、リンクをリストボックスに追加します。
            // リストボックスのタグに現在のパスを設定し、タブのテキストを現在のディレクトリ名に設定します。
            if (!HasAccessPermission(path))
            {
                MessageBox.Show("指定されたパスにアクセスする権限がありません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            listBox.Items.Clear();
            DirectoryInfo parentDir = Directory.GetParent(path);
            if (parentDir != null)
            {
                listBox.Items.Add("ひとつ前に戻る");
            }
            foreach (string dir in Directory.GetDirectories(path))
            {
                DirectoryInfo info = new DirectoryInfo(dir);
                string lastModified = info.LastWriteTime.ToString("yyyy/MM/dd HH:mm");
                listBox.Items.Add(FormatString("Dir ", Path.GetFileName(dir), "-", lastModified));
            }
            foreach (string file in Directory.GetFiles(path))
            {
                if (!file.EndsWith(".lnk")) // .lnkファイルは除外
                {
                    FileInfo info = new FileInfo(file);
                    string size = (info.Length / 1024).ToString() + " KB";
                    string lastModified = info.LastWriteTime.ToString("yyyy/MM/dd HH:mm");
                    listBox.Items.Add(FormatString("File", Path.GetFileName(file), size, lastModified));
                }
            }
            foreach (string link in Directory.GetFiles(path, "*.lnk"))
            {
                listBox.Items.Add(FormatString("Link", Path.GetFileName(link), "-", "-"));
            }

            // ネットワークのインデックスを再計算し、タブのTagプロパティに保存
            tabPage.Tag = new TabData
            {
                Path = path,
                NetworkIndex = path[0] % 4, // パスの先頭文字をASCII値に変換し、その値を4で割った余りを使用
                CloseButton = new Rectangle() // 閉じるボタンの矩形を初期化
            };

            tabPage.Text = new DirectoryInfo(path).Name;
            listBox.Tag = path;
        }

        private string FormatString(string type, string name, string size, string lastModified)
        {
            int nameLength = Encoding.GetEncoding("Shift_JIS").GetByteCount(name);
            int padding = 50 - nameLength;
            if (padding < 0) // itemNameの長さが40を超えている場合
            {
                padding = 0; // paddingを0に設定
            }
            return string.Format("[{0}]: {1}{2} サイズ: {3,-12} 更新日時: {4}", type, name, new string(' ', padding), size, lastModified);
        }


        public void OpenNewTab(string path)
        {
            // このメソッドは、指定されたパスで新しいタブを開きます。
            // パスのアクセス権限を確認し、新しいタブとリストボックスを作成します。
            // リストボックスのタグに現在のパスを設定し、タブをタブコントロールに追加します。
            if (!HasAccessPermission(path))
            {
                MessageBox.Show("指定されたパスにアクセスする権限がありません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string dirName = new DirectoryInfo(path).Name;
            TabPage tabPage = new TabPage(dirName)
            {
                BackColor = Color.Black,
                ForeColor = Color.White,
            };

            ListBox listBox = CreateListBox(path, tabPage);
            listBox.Tag = path;

            tabPage.Controls.Add(listBox);
            // ネットワークのインデックスをタブのTagプロパティに保存
            tabPage.Tag = new TabData
            {
                Path = path,
                NetworkIndex = path[0] % 4, // パスの先頭文字をASCII値に変換し、その値を4で割った余りを使用
                CloseButton = new Rectangle() // 閉じるボタンの矩形を初期化
            };
            tabControl1.TabPages.Add(tabPage);
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
            else if ((e.Alt && e.KeyCode == Keys.L))
            {
                // ALT+Lキーが押されたときの処理
                if (tabControl1.SelectedIndex < tabControl1.TabPages.Count - 1) // タブが最後でなければ
                {
                    tabControl1.SelectedIndex++; // タブを一つ右に切り替える
                }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.L)
            {
                dropDownMenu.Show(listBox, new Point(0,0));
                // コンテキストメニューにフォーカスを設定します。
                dropDownMenu.Focus();

                // 最初の項目を選択します。
                if (dropDownMenu.Items.Count > 0)
                {
                    ((ToolStripMenuItem)dropDownMenu.Items[0]).Select();
                }
            }
            else if (e.KeyCode == Keys.C)
            {
                btnCopyName_Click(null, null);
            }
            else if (e.KeyCode == Keys.X)
            {
                btnCopyFullPath_Click(null, null);
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
        }

        private ListBox CreateListBox(string path, TabPage tabPage)
        {
            // このメソッドは、新しいリストボックスを作成します。
            // リストボックスのプロパティを設定し、イベントハンドラを追加します。
            // 指定されたパスでリストボックスを更新します。
            // 最後に、新しく作成したリストボックスを返します。
            ListBox listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Black,
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font(myFont, 11),
                ItemHeight = 15,
                DrawMode = DrawMode.OwnerDrawFixed,
                SelectionMode = SelectionMode.MultiExtended,
            };

            listBox.DrawItem += listBox_DrawItem;
            listBox.MouseDoubleClick += listBox_MouseDoubleClick;
            listBox.MouseDown += listBox_MouseDown;
            listBox.KeyDown += HandleKeyDown;
            UpdateListBox(listBox, path, tabPage);

            return listBox;
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
                                Process.Start(fileOpenExe, targetPath);
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
                                        Process.Start(fileOpenExe, targetPath);
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
        }
    }
}