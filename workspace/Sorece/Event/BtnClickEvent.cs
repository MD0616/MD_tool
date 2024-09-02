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
                    Process.Start(GlobalSettings.fileOpenExe, path);
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
                    OpenNewTab(GlobalSettings.homeDirectory);
                }
                else
                {
                    UpdateActiveTab(GlobalSettings.homeDirectory);
                }
            }
            else if (e.Button == MouseButtons.Middle)
            {
                OpenNewTab(GlobalSettings.homeDirectory);
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

        private void btnNewFile_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab != null)
            {
                TabPage activeTab = tabControl1.SelectedTab;
                ListBox listBox = (ListBox)activeTab.Controls[0];
                if (listBox.Tag != null)
                {
                    string path = listBox.Tag.ToString(); // 現在のパスを取得
                    string fileName = Prompt.ShowDialog("新規ファイル名を入力してください", "新規ファイル作成", "");
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
                    string folderName = Prompt.ShowDialog("新規フォルダ名を入力してください", "新規フォルダ作成", "");
                    if (!string.IsNullOrEmpty(folderName))
                    {
                        string fullPath = Path.Combine(path, folderName);
                        Directory.CreateDirectory(fullPath);
                        UpdateListBox(listBox, path, activeTab);
                    }
                }
            }
        }

        // Copy&Moveボタンのクリックイベントハンドラ
        private void btnCopyMove_Click(object sender, EventArgs e)
        {
            // コピー&移動ウィンドウを表示
            CopyMoveForm copyMoveForm = new CopyMoveForm();
            copyMoveForm.Show();
        }
    }
}