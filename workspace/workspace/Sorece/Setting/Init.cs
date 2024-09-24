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
                    Arguments = "-NoExit -Command \"& {Write-Host 'PowerShell Ready'; pwd;}\"",
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
            this.BackColor = GlobalSettings.formBackColor;
            this.ForeColor = GlobalSettings.formTextColor;

            // PowerShellの出力を表示するTextBoxの初期化
            txtPowerShellOutput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Bottom,
                BackColor = GlobalSettings.txtBackColor,
                ForeColor = GlobalSettings.txtTextColor,
                Height = 1000,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ReadOnly = true
            };

            // PowerShellの入力を受け取るTextBoxの初期化
            txtPowerShellInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                BackColor = GlobalSettings.txtBackColor,
                ForeColor = GlobalSettings.txtTextColor,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont)
            };
            txtPowerShellInput.KeyDown += TxtPowerShellInput_KeyDown;

            // コンテキストメニューの初期化
            contextMenu = new ContextMenuStrip();
            {
                contextMenu.BackColor = GlobalSettings.contxtBackColor; // 背景色を黒に設定
                contextMenu.ForeColor = GlobalSettings.contxtTextColor; // 文字色を白に設定
                contextMenu.Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont); // 文字サイズを大きくする
                contextMenu.Renderer = new MyRenderer();

                // CSVファイルが存在する場合のみ処理を行う
                if (File.Exists(GlobalSettings.csvPath))
                {
                    using(var reader = new StreamReader(GlobalSettings.csvPath)){
                        // ヘッダ行をスキップ
                        reader.ReadLine();

                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(',');

                            // 値を取得
                            string displayName = values[0];
                            string fullPath = values[1];

                            // メニューアイテムの作成
                            ToolStripMenuItem menuItem1 = new ToolStripMenuItem(displayName);
                            menuItem1.Tag = fullPath; // フルパスをTagプロパティに保存
                            menuItem1.BackColor = GlobalSettings.menuBackColor; // メニューアイテムの背景色を黒に設定
                            menuItem1.ForeColor = GlobalSettings.menuTextColor; // メニューアイテムの文字色を白に設定

                            contextMenu.Items.Add(menuItem1);
                        }
                    }
                }
            };
            contextMenu.MouseClick += new MouseEventHandler(contextMenuShortcut_Click);

            // 新しいコンテキストメニューの初期化
            dropDownMenu = new ContextMenuStrip();
            {
                dropDownMenu.BackColor = GlobalSettings.dropDownBackColor;
                dropDownMenu.ForeColor = GlobalSettings.dropDownTextColor;
                dropDownMenu.Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont); // 文字サイズを大きくする
                dropDownMenu.Renderer = new MyRenderer();

                // アイテムの作成
                ToolStripMenuItem menuItemRename = new ToolStripMenuItem("リネーム");
                menuItemRename.BackColor = GlobalSettings.menuBackColor;
                menuItemRename.ForeColor = GlobalSettings.menuTextColor;
                menuItemRename.Click += eventRename_Click;

                ToolStripMenuItem menuItemDelete = new ToolStripMenuItem("削除");
                menuItemDelete.BackColor = GlobalSettings.menuBackColor;
                menuItemDelete.ForeColor = GlobalSettings.menuTextColor;
                menuItemDelete.Click += eventDelete_Click;

                ToolStripMenuItem menuItemGetFileName = new ToolStripMenuItem("ファイル名取得");
                menuItemGetFileName.BackColor = GlobalSettings.menuBackColor;
                menuItemGetFileName.ForeColor = GlobalSettings.menuTextColor;
                menuItemGetFileName.Click += eventCopyName_Click;

                ToolStripMenuItem menuItemGetPath = new ToolStripMenuItem("パス取得");
                menuItemGetPath.BackColor = GlobalSettings.menuBackColor;
                menuItemGetPath.ForeColor = GlobalSettings.menuTextColor;
                menuItemGetPath.Click += eventCopyFullPath_Click;

                ToolStripMenuItem menuItemSafeOpen = new ToolStripMenuItem("SafeOpen");
                menuItemSafeOpen.BackColor = GlobalSettings.menuBackColor;
                menuItemSafeOpen.ForeColor = GlobalSettings.menuTextColor;
                menuItemSafeOpen.Click += eventSafeOpen_Click;

                ToolStripMenuItem menuItemOpenVScode = new ToolStripMenuItem("VScodeで開く");
                menuItemOpenVScode.BackColor = GlobalSettings.menuBackColor;
                menuItemOpenVScode.ForeColor = GlobalSettings.menuTextColor;
                menuItemOpenVScode.Click += eventVSCode_Click;

                ToolStripMenuItem menuItemDuplicate = new ToolStripMenuItem("複製");
                menuItemDuplicate.BackColor = GlobalSettings.menuBackColor;
                menuItemDuplicate.ForeColor = GlobalSettings.menuTextColor;
                menuItemDuplicate.Click += eventDuplicate_Click;

                // アイテムをコンテキストメニューに追加
                dropDownMenu.Items.Add(menuItemRename);
                dropDownMenu.Items.Add(menuItemDelete);
                dropDownMenu.Items.Add(menuItemGetFileName);
                dropDownMenu.Items.Add(menuItemGetPath);
                dropDownMenu.Items.Add(menuItemSafeOpen);
                dropDownMenu.Items.Add(menuItemOpenVScode);
                dropDownMenu.Items.Add(menuItemDuplicate);

                // KeyDownイベントハンドラを設定
                dropDownMenu.KeyDown += dropDownMenu_KeyDown;
            };

            // プルダウンメニューの初期化
            scriptComboBox = new ComboBox();
            {
                if (Directory.Exists(GlobalSettings.myToolPath))
                {
                    // 指定のフォルダのps1スクリプトをプルダウンで表示
                    var scripts = Directory.GetFiles(@GlobalSettings.myToolPath,"*.ps1");
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
                scriptComboBox.BackColor = GlobalSettings.comboBoxBackColor; // 背景色を黒に設定
                scriptComboBox.ForeColor = GlobalSettings.comboBoxTextColor; // 文字色を白に設定
                scriptComboBox.Width = 250;
                scriptComboBox.Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont); // 文字サイズを大きくする
                scriptComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right; // ウィンドウのサイズに合わせて伸縮
            };

            // 検索バーの初期化
            txtSearchBar = new TextBox
            {
                Name = "SearchBar",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                BorderStyle = BorderStyle.FixedSingle, // 枠線スタイルを固定に設定
                Font = new Font(GlobalSettings.myFont, GlobalSettings.textSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            // プレースホルダー設定
            txtSearchBar.Enter += (sender, e) => ClearSearchBarPlaceholder();
            txtSearchBar.Leave += (sender, e) => SetSearchBarPlaceholder();

            txtSearchBar.KeyDown += new KeyEventHandler(btn_KeyDown);

            // 実行ボタンの初期化
            btnExecute = new Button
            {
                Text = "実行",
                Name = "ScriptDo",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnExecute.Click += new EventHandler(btnExecute_Click);

            // ショートカットボタンの初期化
            btnShortcut = new Button
            {
                Text = "リンク",
                Name = "Link",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnShortcut.Click += new EventHandler(btnShortcut_Click);

            // 決定ボタンの初期化
            btnSearch = new Button
            {
                Text = "決定",
                Name = "Search",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnSearch.KeyDown += new KeyEventHandler(btn_KeyDown);
            btnSearch.MouseDown += new MouseEventHandler(btnSearch_Click);

            // ホームボタンの初期化
            btnHome = new Button
            {
                Text = "ホーム",
                Name = "Home",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnHome.KeyDown += new KeyEventHandler(btn_KeyDown);
            btnHome.MouseDown += new MouseEventHandler(btnHome_Click);

            // 更新ボタンの初期化
            btnRefresh = new Button
            {
                Text = "更新",
                Name = "Refresh",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnRefresh.Click += new EventHandler(btnRefresh_Click);

            // エクスプローラボタンの初期化
            btnExplorer = new Button
            {
                Text = "Explorer",
                Name = "Explorer",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top| AnchorStyles.Left
            };
            btnExplorer.Click += new EventHandler(btnExplorer_Click);

            // ターミナルボタンの初期化
            btnTerminal = new Button
            {
                Text = "Terminal",
                Name = "Terminal",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Bottom| AnchorStyles.Right
            };
            btnTerminal.MouseDown += new MouseEventHandler(btnTerminal_MouseDown);

            // Labelコントロールの初期化
            labelSelectedCount = new Label
            {
                Text = "選択された数: 0",
                Name = "ListCount",
                BackColor = GlobalSettings.labelBackColor,
                ForeColor = GlobalSettings.labelTextColor,
                AutoSize = true, // サイズを自動調整
                Font = new Font(GlobalSettings.myFont, GlobalSettings.labelSizeFont), // フォントを設定
                Anchor = AnchorStyles.Top | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            // タブコントロールの初期化
            tabControl1 = new TabControl
            {
                Name = "TabCtrl",
                BackColor = GlobalSettings.tabBackColor,
                ForeColor = GlobalSettings.tabTextColor,
                Size = new Size(700, 400), // サイズを指定
                Font = new Font(GlobalSettings.myFont, GlobalSettings.tabSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
                DrawMode = TabDrawMode.OwnerDrawFixed, // タブの描画モードを設定
                Padding = new Point(15, 3), // タブの幅を設定
            };
            // イベントハンドラの設定
            tabControl1.DrawItem += new DrawItemEventHandler(tabControl1_DrawItem);
            tabControl1.MouseDown += new MouseEventHandler(tabControl1_MouseDown);
            // タブ切り替えイベントを設定
            tabControl1.SelectedIndexChanged += TabControl1_SelectedIndexChanged;

            // 現在ディレクトリのフルパスをクリップボードに追加するボタン
            btnCopyPath = new Button
            {
                Text = "Get PWD",
                Name = "GetPWD",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnCopyPath.Click += new EventHandler(btnCopyPath_Click);

            // 新規ファイルを作成するボタン
            btnNewFile = new Button
            {
                Text = "New File",
                Name = "NewFile",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnNewFile.Click += new EventHandler(btnNewFile_Click);

            // 新規フォルダを作成するボタン
            btnNewFolder = new Button
            {
                Text = "New Dir",
                Name = "NewDir",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnNewFolder.Click += new EventHandler(btnNewFolder_Click);

            btnCopyMove = new Button
            {
                Text = "Copy&Move",
                Name = "Copy&Move",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnCopyMove.Click += btnCopyMove_Click;

            btnSearchFiles = new Button
            {
                Text = "SearchFile",
                Name = "SearchFile",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnSearchFiles.Click += btnSearchFiles_Click;

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
                RowCount = 7 // 行数を設定
            };

            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F)); // 1列目の幅を100%に設定
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // 2列目の幅を自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 4行目の高さを100%に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3行目の高さを自動調整に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 5行目の高さを100%に設定
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 6行目の高さを自動調整に設定

            // ボタンをFlowLayoutPanelに追加
            topPanel.Controls.Add(btnHome);
            topPanel.Controls.Add(btnRefresh);
            topPanel.Controls.Add(btnCopyPath);
            topPanel.Controls.Add(scriptComboBox);
            topPanel.Controls.Add(btnExecute);

            // ボタンをFlowLayoutPanelに追加
            bottomPanel.Controls.Add(btnExplorer);
            bottomPanel.Controls.Add(btnTerminal);
            bottomPanel.Controls.Add(btnNewFile);
            bottomPanel.Controls.Add(btnNewFolder);
            bottomPanel.Controls.Add(btnCopyMove);
            bottomPanel.Controls.Add(btnShortcut);
            bottomPanel.Controls.Add(btnSearchFiles);

            // 検索バーと決定ボタンをTableLayoutPanelに追加
            tableLayoutPanel.Controls.Add(txtSearchBar, 0, 0);
            tableLayoutPanel.Controls.Add(btnSearch, 1, 0);
            tableLayoutPanel.Controls.Add(topPanel, 0, 1);
            tableLayoutPanel.Controls.Add(bottomPanel, 0, 2);
            tableLayoutPanel.Controls.Add(tabControl1, 0, 3);
            tableLayoutPanel.Controls.Add(labelSelectedCount, 0, 4);
            tableLayoutPanel.Controls.Add(txtPowerShellOutput, 0, 5);
            tableLayoutPanel.Controls.Add(txtPowerShellInput, 0, 6);
            tableLayoutPanel.SetColumnSpan(topPanel, 2);
            tableLayoutPanel.SetColumnSpan(bottomPanel, 2);
            tableLayoutPanel.SetColumnSpan(tabControl1, 2);
            tableLayoutPanel.SetColumnSpan(tableLayoutPanel, 2);
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