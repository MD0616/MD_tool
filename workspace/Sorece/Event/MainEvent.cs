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
            // 現在のアクティブタブのディレクトリを取得
            var tabData = (TabData)tabControl1.SelectedTab.Tag;
            string directoryPath = tabData.Path;
            return directoryPath;
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
                CloseButton = new Rectangle() // 閉じるボタンの矩形を初期化
            };
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
                CloseButton = new Rectangle() // 閉じるボタンの矩形を初期化
            };

            tabPage.Text = new DirectoryInfo(path).Name;
            listBox.Tag = path;
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

    }
}