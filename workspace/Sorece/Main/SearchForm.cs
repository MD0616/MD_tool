using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

// NPOIを使ったExcel検索に必要
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

// DocXを使ったWord検索に必要
using Xceed.Words.NET;

// iTextSharpを使ったPDF検索に必要
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;


namespace MD_Explorer
{
    public class OutputForm : Form
    {
        private ListBox lstResults;
        private TextBox txtDetails;
        private Label lblSearchType;
        public OutputForm()
        {
            // フォームの設定
            this.Text = "検索結果";
            this.Size = new Size(800, 600);
            // ダークモードの設定
            this.BackColor = GlobalSettings.formBackColor;
            this.ForeColor = GlobalSettings.formTextColor;

            // ラベルの初期化
            lblSearchType = new Label
            {
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Height = 30 // ラベルの高さを設定してスペースを確保
            };

            // リストボックス上部の説明ラベルの初期化
            var lblResultsDescription = new Label
            {
                Text = "検索結果一覧",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Height = 30 // ラベルの高さを設定
            };

            // テキストボックス上部の説明ラベルの初期化
            var lblDetailsDescription = new Label
            {
                Text = "選択した項目の詳細",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Height = 30 // ラベルの高さを設定
            };

            // リストボックスの初期化
            lstResults = new ListBox
            {
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.txtBackColor,
                Dock = DockStyle.Fill
            };
            lstResults.SelectedIndexChanged += LstResults_SelectedIndexChanged;

            // テキストボックスの初期化
            txtDetails = new TextBox
            {
                Multiline = true,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.txtBackColor,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill
            };

            // TableLayoutPanelを使用してレイアウトを調整
            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4
            };
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // ラベルの高さ
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30)); // 説明ラベルの高さ
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // リストボックスとテキストボックスの高さ
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); // リストボックスの幅
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); // テキストボックスの幅

            tableLayoutPanel.Controls.Add(lblSearchType, 0, 0);
            tableLayoutPanel.SetColumnSpan(lblSearchType, 2); // ラベルを2列にまたがるように設定
            tableLayoutPanel.Controls.Add(lblResultsDescription, 0, 1);
            tableLayoutPanel.Controls.Add(lblDetailsDescription, 1, 1);
            tableLayoutPanel.Controls.Add(lstResults, 0, 2);
            tableLayoutPanel.Controls.Add(txtDetails, 1, 2);

            Controls.Add(tableLayoutPanel);
        }

        // リストボックスにアイテムを追加するメソッド
        public void AddResult(string fileName, DateTime lastWriteTime, string filePath, string grepContent, string searchType)
        {
            if (lblSearchType.Text != searchType)
            {
                lblSearchType.Text = searchType;
            }
            lstResults.Items.Add(new SearchResult { FileName = fileName, LastWriteTime = lastWriteTime, FilePath = filePath, GrepContent = grepContent });
        }

        // リストボックスの選択が変更されたときのイベントハンドラ
        private void LstResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedItem = lstResults.SelectedItem as SearchResult;
            if (selectedItem != null)
            {
                txtDetails.Text = String.Format("ファイルパス: {0}\r\n\r\n", selectedItem.FilePath);
                if (!string.IsNullOrEmpty(selectedItem.GrepContent))
                {
                    txtDetails.Text += String.Format("grep検索結果:\r\n{0}", selectedItem.GrepContent);
                }
            }
        }

        // 検索結果を格納するクラス
        public class SearchResult
        {
            public string FileName { get; set; }
            public DateTime LastWriteTime { get; set; }
            public string FilePath { get; set; }
            public string GrepContent { get; set; }
            public override string ToString()
            {
                return String.Format("{0} (更新日時: {1})", FileName, LastWriteTime);
            }
        }
    }
    public class SearchForm : Form
    {
        private TextBox txtSearchPath;
        private TextBox txtExcludeExtensions;
        private TextBox txtIncludeExtensions;
        private TextBox txtSearchWord; // 検索ワードのテキストボックスを追加
        private RadioButton rdoFileNameSearch;
        private RadioButton rdoGrepSearch;
        private CheckBox chkRecursive;
        private Button btnSearch;

        public SearchForm()
        {
            // ダークモードの設定
            this.BackColor = GlobalSettings.formBackColor;
            this.ForeColor = GlobalSettings.formTextColor;

            // 各コントロールの初期化
            txtSearchPath = new TextBox
            {
                Text = "検索対象のパス",
                BackColor = GlobalSettings.txtBackColor,
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            txtExcludeExtensions = new TextBox
            {
                Text = "除外する拡張子 (カンマ区切り)",
                BackColor = GlobalSettings.txtBackColor,
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            txtIncludeExtensions = new TextBox
            {
                Text = "検索対象の拡張子 (カンマ区切り)",
                BackColor = GlobalSettings.txtBackColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            // 検索ワードのテキストボックスを初期化
            txtSearchWord = new TextBox
            {
                Text = "検索ワード",
                BackColor = GlobalSettings.txtBackColor,
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };

            rdoFileNameSearch = new RadioButton
            {
                Text = "ファイル名検索",
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
            };

            rdoGrepSearch = new RadioButton
            {
                Text = "grep検索",
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
            };

            chkRecursive = new CheckBox
            {
                Text = "再帰的に検索",
                ForeColor = GlobalSettings.txtTextColor,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
            };

            // 検索実行ボタンの初期化
            btnSearch = new Button
            {
                Text = "検索実行",
                BackColor = GlobalSettings.btnBackColor,
                ForeColor = GlobalSettings.btnTextColor,
                Width = GlobalSettings.btnSizeWidth,
                Height = GlobalSettings.btnSizeHeight,
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.btnSizeFont), // 文字サイズを大きくする
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right, // ウィンドウのサイズに合わせて伸縮
            };
            btnSearch.Click += new EventHandler(BtnSearchFiles_Click);

            // プレースホルダーテキストの設定
            SetPlaceholder(txtSearchPath, "検索対象のパス");
            SetPlaceholder(txtExcludeExtensions, "除外する拡張子 (カンマ区切り)");
            SetPlaceholder(txtIncludeExtensions, "検索対象の拡張子 (カンマ区切り)");
            SetPlaceholder(txtSearchWord, "検索ワード");

            // 説明ラベルを追加
            var lblSearchPathDescription = new Label
            {
                Text = "検索対象のパス",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };

            var lblSearchWordDescription = new Label
            {
                Text = "検索ワード",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };

            var lblExcludeExtensionsDescription = new Label
            {
                Text = "除外する拡張子",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };

            var lblIncludeExtensionsDescription = new Label
            {
                Text = "検索対象の拡張子",
                Font = new System.Drawing.Font(GlobalSettings.myFont, GlobalSettings.textSizeFont),
                ForeColor = GlobalSettings.txtTextColor,
                BackColor = GlobalSettings.formBackColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill
            };

            // レイアウト設定
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 9 };
            
            // 各説明ラベルとテキストボックスをレイアウトに追加
            layout.Controls.Add(lblSearchPathDescription, 0, 0);
            layout.Controls.Add(txtSearchPath, 0, 1);
            layout.Controls.Add(lblSearchWordDescription, 0, 2);
            layout.Controls.Add(txtSearchWord, 0, 3);
            layout.Controls.Add(lblExcludeExtensionsDescription, 0, 4);
            layout.Controls.Add(txtExcludeExtensions, 0, 5);
            layout.Controls.Add(lblIncludeExtensionsDescription, 0, 6);
            layout.Controls.Add(txtIncludeExtensions, 0, 7);

            // ラジオボタンとチェックボックスのレイアウト
            var radioButtonLayout = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            radioButtonLayout.Controls.Add(rdoFileNameSearch);
            radioButtonLayout.Controls.Add(rdoGrepSearch);
            radioButtonLayout.Controls.Add(chkRecursive); // チェックボックスを追加
            layout.Controls.Add(radioButtonLayout, 0, 8);

            layout.Controls.Add(btnSearch, 0, 9);
            Controls.Add(layout);

            // フォームの設定
            Text = "MD_SearchFiles";
            Width = 900;
            Height = 400;
        }

        private void SetPlaceholder(TextBox textBox, string placeholder)
        {
            textBox.ForeColor = Color.Gray;
            textBox.Text = placeholder;
            textBox.Enter += (sender, e) =>
            {
                if (textBox.Text == placeholder)
                {
                    textBox.Text = "";
                    textBox.ForeColor = Color.White;
                }
            };
            textBox.Leave += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    textBox.Text = placeholder;
                    textBox.ForeColor = Color.Gray;
                }
            };
        }

        private void BtnSearchFiles_Click(object sender, EventArgs e)
        {
            string searchPath = txtSearchPath.Text;
            string[] excludeExtensions = string.IsNullOrWhiteSpace(txtExcludeExtensions.Text) ? new string[0] : txtExcludeExtensions.Text.Split(',');
            string[] includeExtensions = string.IsNullOrWhiteSpace(txtIncludeExtensions.Text) ? new string[0] : txtIncludeExtensions.Text.Split(',');
            string searchWord = txtSearchWord.Text; // 検索ワードを取得
            bool isFileNameSearch = rdoFileNameSearch.Checked;
            bool isGrepSearch = rdoGrepSearch.Checked;
            bool isRecursive = chkRecursive.Checked;

            // ラジオボタンがどちらも選択されていない場合はエラーメッセージを表示して検索を中止
            if (!isFileNameSearch && !isGrepSearch)
            {
                MessageBox.Show("ファイル名検索またはgrep検索のいずれかを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 出力フォームの初期化
            var outputForm = new OutputForm();
            outputForm.Show();

            try
            {
                // 検索処理の実装
                SearchFiles(searchPath, excludeExtensions, includeExtensions, searchWord, isFileNameSearch, isGrepSearch, isRecursive, outputForm);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("正規表現のパターンが無効です: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SearchFiles(string path, string[] excludeExtensions, string[] includeExtensions, string searchWord, bool isFileNameSearch, bool isGrepSearch, bool isRecursive, OutputForm outputForm)
        {
            // パスが存在するかどうかを確認
            if (!Directory.Exists(path))
            {
                MessageBox.Show("指定されたパスが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 正規表現オブジェクトを作成
            Regex regex = null;
            if (!string.IsNullOrEmpty(searchWord))
            {
                try
                {
                    regex = new Regex(searchWord);
                }
                catch (ArgumentException ex)
                {
                    throw new ArgumentException("正規表現のパターンが無効です: " + ex.Message);
                }
            }

            // ファイル検索ロジックの実装
            var files = Directory.EnumerateFiles(path, "*.*", isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);

            foreach (var file in files)
            {
                string extension = System.IO.Path.GetExtension(file).TrimStart('.');

                // 除外する拡張子が指定されている場合、その拡張子を持つファイルを除外
                if (excludeExtensions.Length > 0 && excludeExtensions.Contains(extension))
                {
                    continue;
                }

                // 検索対象の拡張子が指定されている場合、その拡張子を持つファイルのみを対象
                if (includeExtensions == null &&includeExtensions.Length > 0 && !includeExtensions.Contains(extension))
                {
                    continue;
                }

                // ファイル名検索
                string grepContent = null;
                if (isFileNameSearch && (regex == null || regex.IsMatch(System.IO.Path.GetFileName(file))))
                {
                    // ファイル名検索ロジック
                    outputForm.AddResult(System.IO.Path.GetFileName(file), File.GetLastWriteTime(file), System.IO.Path.GetDirectoryName(file), grepContent, "ファイル名検索結果");
                }

                // grep検索
                if (isGrepSearch)
                {
                    // Excelファイル検索
                    if (extension == "xls" || extension == "xlsx")
                    {
                        SearchExcelFile(file, searchWord, regex, outputForm);
                        continue; // 処理を次のファイルへ
                    }

                    // Wordファイル検索
                    if (extension == "doc" || extension == "docx")
                    {
                        SearchWordFile(file, searchWord, regex, outputForm);
                        continue; // 処理を次のファイルへ
                    }

                    // PDFファイル検索
                    if (extension == "pdf")
                    {
                        SearchPdfFile(file, searchWord, regex, outputForm);
                        continue; // 処理を次のファイルへ
                    }

                    // grep検索ロジック
                    var lines = File.ReadLines(file);
                    var matchedLines = new System.Text.StringBuilder();

                    int lineNumber = 0;
                    foreach (var line in lines)
                    {
                        lineNumber++;
                        if (regex == null || regex.IsMatch(line))
                        {
                            matchedLines.AppendLine(String.Format("{0}：{1}", lineNumber, line));
                        }
                    }

                    if (matchedLines.Length > 0)
                    {
                        grepContent = matchedLines.ToString();
                        outputForm.AddResult(System.IO.Path.GetFileName(file), File.GetLastWriteTime(file), System.IO.Path.GetDirectoryName(file), grepContent, "grep検索結果");
                    }
                }
            }
        }

        private void SearchExcelFile(string filePath, string searchWord, Regex regex, OutputForm outputForm)
        {
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (filePath.EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook(fs); // .xlsx ファイルを読み込む
                }
                else if (filePath.EndsWith(".xls"))
                {
                    workbook = new HSSFWorkbook(fs); // .xls ファイルを読み込む
                }
                else
                {
                    return;
                }
            }

            StringBuilder excelText = new StringBuilder();
            excelText.AppendLine("Excelファイル: " + System.IO.Path.GetFileName(filePath));  // ファイル名を最初に出力
            bool hasMatches = false;

            // シート名も検索対象に含めて処理
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                StringBuilder sheetText = new StringBuilder();
                string sheetName = sheet.SheetName;
                bool sheetHasMatches = false;

                // シート名に対する検索
                if (regex == null || regex.IsMatch(sheetName))
                {
                    sheetText.AppendLine("シート名: " + sheetName);
                    sheetHasMatches = true;
                }

                // シート内のセルに対する検索
                for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row == null) continue;

                    for (int colIndex = row.FirstCellNum; colIndex < row.LastCellNum; colIndex++)
                    {
                        ICell cell = row.GetCell(colIndex);
                        if (cell == null) continue;

                        string cellValue = cell.ToString();
                        if (regex == null || regex.IsMatch(cellValue))
                        {
                            sheetText.AppendLine("セル[" + (rowIndex + 1) + ", " + (colIndex + 1) + "]: " + cellValue);
                            sheetHasMatches = true;
                        }
                    }
                }

                // シート名またはセルがヒットした場合に出力
                if (sheetHasMatches)
                {
                    excelText.AppendLine("シート: " + sheetName);
                    excelText.Append(sheetText.ToString());
                    hasMatches = true;
                }
            }

            // ヒットした内容があれば出力
            if (hasMatches)
            {
                string result = excelText.ToString();
                outputForm.AddResult(System.IO.Path.GetFileName(filePath), File.GetLastWriteTime(filePath), filePath, result, "grep検索結果");
            }
        }

        private void SearchWordFile(string filePath, string searchWord, Regex regex, OutputForm outputForm)
        {
            using (var document = DocX.Load(filePath))
            {
                StringBuilder wordText = new StringBuilder();
                wordText.AppendLine("Wordファイル: " + System.IO.Path.GetFileName(filePath));  // ファイル名を最初に出力

                int pageIndex = 1;
                bool hasMatches = false;

                foreach (Paragraph paragraph in document.Paragraphs)
                {
                    string text = paragraph.Text;
                    if (regex == null || regex.IsMatch(text))
                    {
                        wordText.AppendLine("ページ " + pageIndex + ":\r\n" + text);
                        hasMatches = true;
                    }
                    pageIndex++;  // 段落ごとに仮想的にページを増やす
                }

                if (hasMatches)
                {
                    string result = wordText.ToString();
                    outputForm.AddResult(System.IO.Path.GetFileName(filePath), File.GetLastWriteTime(filePath), filePath, result, "grep検索結果");
                }
            }
        }

        private void SearchPdfFile(string filePath, string searchWord, Regex regex, OutputForm outputForm)
        {
            using (PdfReader reader = new PdfReader(filePath))
            {
                StringBuilder pdfText = new StringBuilder();
                pdfText.AppendLine("PDFファイル: " + System.IO.Path.GetFileName(filePath));  // System.IO.Path を明示的に指定
                
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, i);
                    if (regex == null || regex.IsMatch(pageText))
                    {
                        pdfText.AppendLine("ページ " + i + ":\r\n" + pageText);
                    }
                }

                // 検索結果を1つにまとめて出力
                string result = pdfText.ToString();
                if (!string.IsNullOrEmpty(result))
                {
                    outputForm.AddResult(System.IO.Path.GetFileName(filePath), File.GetLastWriteTime(filePath), filePath, result, "grep検索結果");
                }
            }
        }
    }
}
