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

        private string GetCurrentTabDirectory()
        {
            if (tabControl1.SelectedTab == null || tabControl1.SelectedTab.Tag == null)
            {
                // タブが選択されていない場合やデータがない場合にデフォルトメッセージを返す
                return "ディレクトリが選択されていません";
            }

            var tabData = tabControl1.SelectedTab.Tag as TabData;
            if (tabData != null)
            {
                return tabData.Path;
            }

            // タブデータがnullの場合にもデフォルトメッセージを返す
            return "ディレクトリが選択されていません";
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
        public void OpenNewTab(string path)
        {
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

            // リストボックス作成
            ListBox listBox = CreateListBox(path, tabPage);
            listBox.Tag = path;
            tabPage.Controls.Add(listBox);

            // TabDataに履歴を初期化して設定
            tabPage.Tag = new TabData
            {
                Path = path,
                CloseButton = new Rectangle(),
                History = new Stack<string>()  // 履歴の初期化
            };

            // タブを追加
            tabControl1.TabPages.Add(tabPage);
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
                BackColor = GlobalSettings.listBoxBackColor,
                ForeColor = GlobalSettings.listBoxTextColor,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.listSizeFont),
                ItemHeight = 14,
                DrawMode = DrawMode.OwnerDrawFixed,
                SelectionMode = SelectionMode.MultiExtended,
            };
            listBox.KeyPress += new KeyPressEventHandler(listBox_KeyPress);
            listBox.DrawItem += listBox_DrawItem;
            listBox.MouseDoubleClick += listBox_MouseDoubleClick;
            listBox.MouseDown += listBox_MouseDown;
            listBox.KeyDown += HandleKeyDown;
            UpdateListBox(listBox, path, tabPage);

            return listBox;
        }

        private void UpdateListBox(ListBox listBox, string path, TabPage tabPage)
        {
            if (!HasAccessPermission(path))
            {
                MessageBox.Show("指定されたパスにアクセスする権限がありません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // タブにTagが存在しない場合は新しく作成
            if (tabPage.Tag == null)
            {
                tabPage.Tag = new TabData();
            }

            // タブのTagからTabDataを取得
            var tabData = (TabData)tabPage.Tag;

            // 現在のパスを履歴に保存
            if (listBox.Tag != null && path != (string)listBox.Tag && tabData.History != null)
            {
                tabData.History.Push(listBox.Tag.ToString());
            }

            // 新しいパスをTabDataに更新
            tabData.Path = path;

            listBox.Items.Clear();

            // 戻るアイテムを追加 (履歴が存在する場合のみ)
            if (tabData.History != null && tabData.History.Count > 0)
            {
                listBox.Items.Add("← 戻る"); // リストの一番上に戻るボタンを追加
            }

            DirectoryInfo parentDir = Directory.GetParent(path);
            if (parentDir != null)
            {
                listBox.Items.Add("親パスへ");
            }

            foreach (string dir in Directory.GetDirectories(path))
            {
                DirectoryInfo info = new DirectoryInfo(dir);
                string lastModified = info.LastWriteTime.ToString("yyyy/MM/dd HH:mm");
                listBox.Items.Add(FormatString("Dir ", Path.GetFileName(dir), "-", lastModified));
            }

            foreach (string file in Directory.GetFiles(path))
            {
                if (!file.EndsWith(".lnk"))
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

            tabPage.Text = new DirectoryInfo(path).Name;
            listBox.Tag = path;

            // フォルダ変更後にプレースホルダーを更新
            SetSearchBarPlaceholder();
        }

        private string FormatString(string type, string name, string size, string lastModified)
        {
            int nameLength = Encoding.GetEncoding("Shift_JIS").GetByteCount(name);
            int padding = 70 - nameLength;
            if (padding < 0) // itemNameの長さが40を超えている場合
            {
                padding = 0; // paddingを0に設定
            }
            return string.Format("[{0}]: {1}{2} サイズ: {3,-12} 更新日時: {4}", type, name, new string(' ', padding), size, lastModified);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                var tabData = (TabData)activeTab.Tag;

                // 戻る履歴が存在する場合
                if (tabData.History != null && tabData.History.Count > 0)
                {
                    string previousPath = tabData.History.Pop(); // 履歴から前のパスを取得
                    ListBox listBox = (ListBox)activeTab.Controls[0];
                    UpdateListBox(listBox, previousPath, activeTab); // リストを更新
                }
                else
                {
                    MessageBox.Show("これ以上戻れません。", "戻る", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // プレースホルダー機能の実装
        private void SetSearchBarPlaceholder()
        {
            if (txtSearchBar == null) return;

            string currentDir = GetCurrentTabDirectory();

            // プレースホルダーの設定（入力がない場合のみ）
            if (string.IsNullOrEmpty(txtSearchBar.Text) || txtSearchBar.ForeColor == Color.Gray)
            {
                txtSearchBar.ForeColor = Color.Gray; // プレースホルダーの色
                txtSearchBar.Text = string.Format("PWD: {0}", currentDir);
            }
        }

        private void ClearSearchBarPlaceholder()
        {
            if (txtSearchBar.ForeColor == Color.Gray)
            {
                txtSearchBar.Text = "";
                txtSearchBar.ForeColor = GlobalSettings.btnTextColor; // 元の色に戻す
            }
        }

        // タブが変更されたときに呼び出されるイベント
        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetSearchBarPlaceholder();
        }
    }
}
