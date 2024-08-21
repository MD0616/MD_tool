using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace MD_Explorer
{
    public partial class MainForm
    {
        private void InitializePowerShell()
        {
            powerShellProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoExit -Command \"& {Write-Host 'PowerShell Ready'; }\"",
                    RedirectStandardOutput = true,
                    RedirectStandardInput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            powerShellProcess.OutputDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    txtPowerShellOutput.Invoke(new Action(() =>
                    {
                        txtPowerShellOutput.AppendText(e.Data + Environment.NewLine);
                    }));
                }
            };

            powerShellProcess.Start();
            powerShellProcess.BeginOutputReadLine();
        }

        private void InitializeComponent()
        {
            // ダークモードの設定
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White; // 文字色を白に設定

            // PowerShellの出力を表示するTextBoxの初期化
            txtPowerShellOutput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Bottom,
                BackColor = Color.Black,
                Height = 1000,
                ForeColor = Color.White,
                Font = new Font(myFont, 10),
                ReadOnly = true
            };

            // PowerShellの入力を受け取るTextBoxの初期化
            txtPowerShellInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                BackColor = Color.Black,
                ForeColor = Color.White,
                Font = new Font(myFont, 10)
            };
            txtPowerShellInput.KeyDown += TxtPowerShellInput_KeyDown;


            // 検索バーの初期化
            txtSearchBar = new TextBox
            {
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                BorderStyle = BorderStyle.FixedSingle, // 枠線スタイルを固定に設定
                Font = new Font(myFont, 13), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            // 実行ボタンの初期化
            btnExecute = new Button
            {
                Text = "実行",
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                FlatStyle = FlatStyle.Flat, // フラットスタイルに設定
                Width = 80,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnExecute.FlatAppearance.BorderColor = Color.Gray; // 縁を灰色に設定
            btnExecute.Click += new EventHandler(btnExecute_Click);

            // コンテキストメニューの初期化
            contextMenu = new ContextMenuStrip();
            {
                contextMenu.BackColor = Color.Black; // 背景色を黒に設定
                contextMenu.ForeColor = Color.White; // 文字色を白に設定
                contextMenu.Font = new Font(myFont, 13); // 文字サイズを大きくする

                // CSVファイルが存在する場合のみ処理を行う
                if (File.Exists(csvPath))
                {
                    using(var reader = new StreamReader(csvPath)){
                        // ヘッダ行をスキップ
                        reader.ReadLine();

                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');

                            // 値を取得
                            string displayName = values[0];
                            string className = values[1];
                            string fullPath = values[2];

                            // メニューアイテムの作成
                            ToolStripMenuItem menuItem = new ToolStripMenuItem(displayName);
                            menuItem.Tag = fullPath; // フルパスをTagプロパティに保存
                            menuItem.BackColor = Color.Black; // メニューアイテムの背景色を黒に設定
                            menuItem.ForeColor = Color.White; // メニューアイテムの文字色を白に設定

                            // クラス名がある場合は、その下の階層のコンテキストメニューを作成
                            if (!string.IsNullOrEmpty(className))
                            {
                                ToolStripMenuItem subMenuItem = new ToolStripMenuItem(className);
                                subMenuItem.DropDownItems.Add(menuItem);
                                subMenuItem.BackColor = Color.Black; // サブメニューアイテムの背景色を黒に設定
                                subMenuItem.ForeColor = Color.White; // サブメニューアイテムの文字色を白に設定
                                contextMenu.Items.Add(subMenuItem);
                            }
                            else
                            {
                                contextMenu.Items.Add(menuItem);
                            }
                        }
                    }
                }
            };
            contextMenu.MouseClick += new MouseEventHandler(contextMenuShortcut_Click);

            // プルダウンメニューの初期化
            scriptComboBox = new ComboBox();
            {
                if (Directory.Exists(csvPath))
                {
                    // 指定のフォルダのps1スクリプトをプルダウンで表示
                    var scripts = Directory.GetFiles(@myToolPath,"*.ps1");
                    foreach (var script in scripts)
                    {
                        var scriptInfo = new ScriptInfo
                        {
                            FullPath = script,
                            FileName = Path.GetFileName(script)
                        };
                        scriptComboBox.Items.Add(scriptInfo);
                    }
                }
                scriptComboBox.BackColor = Color.Black; // 背景色を黒に設定
                scriptComboBox.ForeColor = Color.White; // 文字色を白に設定
                scriptComboBox.Width = 250;
                scriptComboBox.Font = new Font(myFont, 13); // 文字サイズを大きくする
                scriptComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right; // ウィンドウのサイズに合わせて伸縮
            };

            // ショートカットボタンの初期化
            btnShortcut = new Button
            {
                Text = "リンク",
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                FlatStyle = FlatStyle.Flat, // フラットスタイルに設定
                Width = 80,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnShortcut.FlatAppearance.BorderColor = Color.Gray; // 縁を灰色に設定
            btnShortcut.Click += new EventHandler(btnShortcut_Click);

            // 削除ボタンの初期化
            btnDelete = new Button
            {
                Text = "削除",
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                FlatStyle = FlatStyle.Flat, // フラットスタイルに設定
                Width = 80,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnDelete.FlatAppearance.BorderColor = Color.Gray; // 縁を灰色に設定
            btnDelete.Click += new EventHandler(btnDelete_Click);

            // 決定ボタンの初期化
            btnSearch = new Button
            {
                Text = "決定",
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                FlatStyle = FlatStyle.Flat, // フラットスタイルに設定
                Width = 80,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnSearch.FlatAppearance.BorderColor = Color.Gray; // 縁を灰色に設定
            btnSearch.MouseDown += new MouseEventHandler(btnSearch_Click);

            // ホームボタンの初期化
            btnHome = new Button
            {
                Text = "ホーム",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnHome.FlatAppearance.BorderColor = Color.Gray;
            btnHome.MouseDown += new MouseEventHandler(btnHome_Click);

            // 更新ボタンの初期化
            btnRefresh = new Button
            {
                Text = "更新",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnRefresh.FlatAppearance.BorderColor = Color.Gray;
            btnRefresh.Click += new EventHandler(btnRefresh_Click);

            // VScodeボタンの初期化
            btnVSCode = new Button
            {
                Text = "VSCode",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top| AnchorStyles.Left
            };
            btnVSCode.FlatAppearance.BorderColor = Color.Gray;
            btnVSCode.Click += new EventHandler(btnVSCode_Click);

            // エクスプローラボタンの初期化
            btnExplorer = new Button
            {
                Text = "Explorer",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top| AnchorStyles.Left
            };
            btnExplorer.FlatAppearance.BorderColor = Color.Gray;
            btnExplorer.Click += new EventHandler(btnExplorer_Click);

            // ターミナルボタンの初期化
            btnTerminal = new Button
            {
                Text = "Terminal",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top| AnchorStyles.Left
            };
            btnTerminal.FlatAppearance.BorderColor = Color.Gray;
            btnTerminal.MouseDown += new MouseEventHandler(btnTerminal_MouseDown);

            // タブコントロールの初期化
            tabControl1 = new TabControl
            {
                BackColor = Color.Black, // 背景色を黒に設定
                ForeColor = Color.White, // 文字色を白に設定
                Size = new Size(700, 400), // サイズを指定
                Font = new Font(myFont, 11), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
                DrawMode = TabDrawMode.OwnerDrawFixed, // タブの描画モードを設定
                Padding = new Point(10, 3), // タブの幅を設定
            };
            // イベントハンドラの設定
            tabControl1.DrawItem += new DrawItemEventHandler(tabControl1_DrawItem);
            tabControl1.MouseDown += new MouseEventHandler(tabControl1_MouseDown);

            // 現在ディレクトリのフルパスをクリップボードに追加するボタン
            btnCopyPath = new Button
            {
                Text = "Get PWD",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnCopyPath.FlatAppearance.BorderColor = Color.Gray;
            btnCopyPath.Click += new EventHandler(btnCopyPath_Click);

            // 選択中のリストのファイル、フォルダ名をクリップボードに追加するボタン
            btnCopyName = new Button
            {
                Text = "ファイル名",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnCopyName.FlatAppearance.BorderColor = Color.Gray;
            btnCopyName.Click += new EventHandler(btnCopyName_Click);

            // 選択中のリストのフルパスをクリップボードに追加するボタン
            btnCopyFullPath = new Button
            {
                Text = "フルパス",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnCopyFullPath.FlatAppearance.BorderColor = Color.Gray;
            btnCopyFullPath.Click += new EventHandler(btnCopyFullPath_Click);

            // 新規ファイルを作成するボタン
            btnNewFile = new Button
            {
                Text = "New File",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnNewFile.FlatAppearance.BorderColor = Color.Gray;
            btnNewFile.Click += new EventHandler(btnNewFile_Click);

            // 新規フォルダを作成するボタン
            btnNewFolder = new Button
            {
                Text = "New Dir",
                BackColor = Color.Black,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Width = 80,
                Height = 25,
                Font = new Font(myFont, 8),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnNewFolder.FlatAppearance.BorderColor = Color.Gray;
            btnNewFolder.Click += new EventHandler(btnNewFolder_Click);

            // FlowLayoutPanelの初期化（ホーム、更新ボタン用）
            FlowLayoutPanel topPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10),
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Top,
            };

            // FlowLayoutPanelの初期化（2段目ボタン行）
            FlowLayoutPanel bottomPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10),
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Top,
            };

            // TableLayoutPanelを使用してレイアウトを調整
            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, // フォーム全体にフィットするように設定
                ColumnCount = 2, // 列数を設定
                RowCount = 6 // 行数を設定
            };

            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // 1列目の幅を100%に設定
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 2列目の幅を自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 4行目の高さを100%に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 5行目の高さを100%に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 6行目の高さを自動調整に設定

            // ボタンをFlowLayoutPanelに追加
            topPanel.Controls.Add(btnHome);
            topPanel.Controls.Add(btnRefresh);
            topPanel.Controls.Add(btnCopyPath);
            topPanel.Controls.Add(btnCopyName);
            topPanel.Controls.Add(btnCopyFullPath);
            topPanel.Controls.Add(scriptComboBox);
            topPanel.Controls.Add(btnExecute);

            // ボタンをFlowLayoutPanelに追加
            bottomPanel.Controls.Add(btnVSCode);
            bottomPanel.Controls.Add(btnExplorer);
            bottomPanel.Controls.Add(btnTerminal);
            bottomPanel.Controls.Add(btnNewFile);
            bottomPanel.Controls.Add(btnNewFolder);
            bottomPanel.Controls.Add(btnDelete);
            bottomPanel.Controls.Add(btnShortcut);

            // 検索バーと決定ボタンをTableLayoutPanelに追加
            tableLayoutPanel.Controls.Add(txtSearchBar, 0, 0);
            tableLayoutPanel.Controls.Add(btnSearch, 1, 0);
            tableLayoutPanel.Controls.Add(topPanel, 0, 1);
            tableLayoutPanel.Controls.Add(bottomPanel, 0, 2);
            tableLayoutPanel.Controls.Add(tabControl1, 0, 3);
            tableLayoutPanel.Controls.Add(txtPowerShellOutput, 0, 4);
            tableLayoutPanel.Controls.Add(txtPowerShellInput, 0, 5);
            tableLayoutPanel.SetColumnSpan(topPanel, 2);
            tableLayoutPanel.SetColumnSpan(bottomPanel, 2);
            tableLayoutPanel.SetColumnSpan(tabControl1, 2);
            tableLayoutPanel.SetColumnSpan(txtPowerShellOutput, 2);
            tableLayoutPanel.SetColumnSpan(txtPowerShellInput, 2);

            // フォームにコンポーネントを追加
            Controls.Add(tableLayoutPanel);

            // フォームの設定
            Text = "MD_Explorer";
            Width = 900;
            Height = 1000;

            // PowerShellの初期化と実行
            InitializePowerShell();
        }
    }
}
