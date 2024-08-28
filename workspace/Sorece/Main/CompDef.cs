using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32.SafeHandles;
using System.Collections.Generic;

namespace MD_Explorer
{
    public static class Prompt
    {
        public static string ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White
            };

            Label textLabel = new Label() { Left = 50, Top = 20, Text = text, AutoSize = true };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400, BackColor = Color.Black, ForeColor = Color.White };

            Button confirmation = new Button() { Text = "OK", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK, BackColor = Color.Black, ForeColor = Color.White };
            confirmation.FlatAppearance.BorderColor = Color.Gray;
            confirmation.FlatStyle = FlatStyle.Flat;

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }
    }

    public partial class CopyMoveForm : Form
    {
        public string[] SourcePaths { get; set; }
        public string[] DestinationPaths { get; set; }
        public string SourceText
        {
            get { return txtSource.Text; }
            set { txtSource.Text = value; }
        }

        public string DestinationText
        {
            get { return txtDestination.Text; }
            set { txtDestination.Text = value; }
        }
        public OperationTypeEnum CurrentOperationType { get; set; } 

        public enum OperationTypeEnum
        {
            Copy,
            Move
        }
        private TextBox txtSource;
        private TextBox txtDestination;
        private Button btnCancel;
        private Button btnCopy;
        private Button btnMove;
        private Label lblDestination;
        private Label lblSource;

        public CopyMoveForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // コントロールの初期化と配置
            lblSource = new Label
            {
                Text = "コピーまたは移動するファイル、ディレクトリをフルパスで記載してください。\n複数アイテムの場合は改行で分けてください。",
                Dock = DockStyle.Top,
                AutoSize = true
            };
            txtSource = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Top,
                Height = 100, // 高さを適切に設定
                BackColor = Color.FromArgb(30, 30, 30), // ダークモードの背景色
                ForeColor = Color.White // テキストの色
            };
            lblDestination = new Label
            {
                Text = "コピー、移動先を記載してください。\n複数個所へコピー場合は改行で分けてください。",
                Dock = DockStyle.Top,
                AutoSize = true
            };
            txtDestination = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Top,
                Height = 100, // 高さを適切に設定
                BackColor = Color.FromArgb(30, 30, 30), // ダークモードの背景色
                ForeColor = Color.White // テキストの色
            };

            // ボタンの初期化
            FlowLayoutPanel panelButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight
            };
            btnCancel = new Button
            {
                Text = "キャンセル",
                BackColor = Color.FromArgb(45, 45, 45), // ダークモードの背景色
                ForeColor = Color.White // テキストの色
            };
            btnCopy = new Button
            {
                Text = "コピー",
                BackColor = Color.FromArgb(45, 45, 45), // ダークモードの背景色
                ForeColor = Color.White // テキストの色
            };
            btnMove = new Button
            {
                Text = "移動",
                BackColor = Color.FromArgb(45, 45, 45), // ダークモードの背景色
                ForeColor = Color.White // テキストの色
            };

            // ボタンをパネルに追加
            panelButtons.Controls.Add(btnCancel);
            panelButtons.Controls.Add(btnCopy);
            panelButtons.Controls.Add(btnMove);

            // イベントハンドラの設定
            btnCancel.Click += (sender, e) => { this.Close(); };
            btnCopy.Click += (sender, e) => {
                this.SourcePaths = txtSource.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                this.DestinationPaths = txtDestination.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                ProcessFiles(this.SourcePaths, this.DestinationPaths, OperationTypeEnum.Copy); 
            };
            btnMove.Click += (sender, e) => {
                this.SourcePaths = txtSource.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                this.DestinationPaths = txtDestination.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                ProcessFiles(this.SourcePaths, this.DestinationPaths, OperationTypeEnum.Move);
            };

            // フォームにコントロールを追加
            Controls.Add(panelButtons); // ボタンパネルを最初に追加
            Controls.Add(txtDestination); // 宛先テキストボックスを追加
            Controls.Add(lblDestination); // 宛先ラベルを追加
            Controls.Add(txtSource); // ソーステキストボックスを追加
            Controls.Add(lblSource); // ソースラベルを追加

            // フォームのプロパティ設定
            this.BackColor = Color.FromArgb(20, 20, 20); // ダークモードの背景色
            this.ForeColor = Color.White; // テキストの色
            this.AutoSize = true; // フォームのサイズをコントロールに合わせて自動調整
            this.Text = "コピーと移動"; // フォームのタイトルを設定
        }
        // ディレクトリを再帰的にコピーするメソッド

        private void CopyDirectory(string sourceDirPath, string destinationDirPath)
        {
            // サブディレクトリを含むすべてのファイルとディレクトリのリストを取得
            foreach (string dirPath in Directory.GetDirectories(sourceDirPath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceDirPath, destinationDirPath));
            }

            // ファイルをコピー
            foreach (string filePath in Directory.GetFiles(sourceDirPath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(filePath, filePath.Replace(sourceDirPath, destinationDirPath), true);
            }
        }

        public void ProcessFiles(string[] sourcePaths, string[] destinationPaths, OperationTypeEnum operationType)
        {
            if (sourcePaths == null || sourcePaths.Length == 0)
            {
                MessageBox.Show("コピーまたは移動するファイル、ディレクトリが指定されていません。");
                return;
            }
            if (destinationPaths == null || destinationPaths.Length == 0)
            {
                MessageBox.Show("コピー、移動先が指定されていません。");
                return;
            }

            List<string> successfulOperations = new List<string>();
            List<string> failedOperations = new List<string>();

            foreach (string sourcePath in sourcePaths)
            {
                string quotedSourcePath = QuotePath(sourcePath);

                if (!File.Exists(sourcePath) && !Directory.Exists(sourcePath))
                {
                    failedOperations.Add(sourcePath);
                    continue;
                }

                // 移動操作の場合、destinationPathsが複数あれば失敗扱いにして進む
                if (operationType == OperationTypeEnum.Move && destinationPaths.Length > 1)
                {
                    failedOperations.Add(sourcePath + " => 複数の移動先が指定されています。");
                    continue; // 次のsourcePathへ
                }

                foreach (string destinationPath in destinationPaths)
                {
                    string quotedDestinationPath = QuotePath(destinationPath);

                    try
                    {
                        switch (operationType)
                        {
                            case OperationTypeEnum.Copy: // 修正された列挙型の名前を使用
                                if (File.Exists(sourcePath))
                                {
                                    File.Copy(sourcePath, Path.Combine(destinationPath, Path.GetFileName(sourcePath)), true);
                                }
                                else if (Directory.Exists(sourcePath))
                                {
                                    CopyDirectory(sourcePath, Path.Combine(destinationPath, new DirectoryInfo(sourcePath).Name));
                                }
                                break;
                            case OperationTypeEnum.Move: // 修正された列挙型の名前を使用
                                if (File.Exists(sourcePath))
                                {
                                    File.Move(sourcePath, Path.Combine(destinationPath, Path.GetFileName(sourcePath)));
                                }
                                else if (Directory.Exists(sourcePath))
                                {
                                    Directory.Move(sourcePath, destinationPath);
                                }
                                break;
                        }
                        successfulOperations.Add(string.Format("{0} <= {1}",destinationPath ,sourcePath));
                    }
                    catch (Exception ex)
                    {
                        // エラーログを出力または表示
                        failedOperations.Add(string.Format("{0} <= {1}: {2}", destinationPath,sourcePath, ex.Message));
                    }
                }
            }
            // 処理結果を表示
            ShowOperationResults(successfulOperations, failedOperations);
        }

        private string QuotePath(string path)
        {
            // パスにスペースが含まれている場合、または引用符が含まれていない場合は引用符で囲む
            if (path.Contains(" ") && !path.StartsWith("\"") && !path.EndsWith("\""))
            {
                return "\"" + path + "\"";
            }
            return path;
        }

        private void ShowOperationResults(List<string> successfulOperations, List<string> failedOperations)
        {
            // 結果を表示するためのテキストボックスまたはラベルを作成
            TextBox txtResults = new TextBox
            {
                Multiline = true,
                BackColor = Color.FromArgb(45, 45, 45), // ダークモードの背景色
                ForeColor = Color.White, // テキストの色
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                ReadOnly = true
            };

            StringBuilder results = new StringBuilder();

            // 成功した操作のリストを追加
            if (successfulOperations.Count > 0)
            {
                results.AppendLine("成功した操作:");
                foreach (string success in successfulOperations)
                {
                    results.AppendLine(success);
                }
            }

            // 失敗した操作のリストを追加
            if (failedOperations.Count > 0)
            {
                results.AppendLine("\n失敗した操作:");
                foreach (string failure in failedOperations)
                {
                    results.AppendLine(failure);
                }
            }

            // 結果をテキストボックスに設定
            txtResults.Text = results.ToString();

            // テキストボックスをフォームまたはダイアログに追加
            this.Controls.Add(txtResults);

            // 必要に応じて、結果を表示するための新しいウィンドウを開く
            Form resultsForm = new Form
            {
                Text = "操作結果",
                Size = new Size(500, 300)
            };
            resultsForm.Controls.Add(txtResults);
            resultsForm.ShowDialog();
        }

    }

    public class TabData
    {
        public string Path { get; set; }
        public int NetworkIndex { get; set; }
        public Rectangle CloseButton { get; set; } // 閉じるボタンの矩形を追加
    }

    // スクリプトの情報を保持するクラス
    public class ScriptInfo
    {
        public string FullPath { get; set; }
        public string FileName { get; set; }

        // ComboBoxに表示する文字列を制御します
        public override string ToString()
        {
            return FileName;
        }
    }
}