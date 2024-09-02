using System;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Xml; // XmlDocument と XmlNodeList のため
using System.Collections.Specialized; // NameValueCollection のため
using System.Drawing;
using System.Collections.Generic;
// Color[] tabColors = { Color.Black, Color.Red, Color.Blue, Color.Green, Color.Brown, Color.Purple, Color.Pink, Color.Turquoise, Color.Orange };
namespace MD_Explorer
{
    // Static class to hold global-like variables
    static class GlobalSettings
    {
        // パスと色のマッピングを定義
        private static Dictionary<string, Color> pathColorMapping = new Dictionary<string, Color>();

        // パスに基づいて色を取得するメソッド
        public static Color GetColorFromPath(string path)
        {
            foreach (var pair in pathColorMapping)
            {
                if (path.StartsWith(pair.Key)) // パスが指定したディレクトリを含むかどうかをチェック
                {
                    return pair.Value;
                }
            }
            // マッピングが見つからない場合のデフォルト色
            return Color.Black;
        }

        public static string homeDirectory = @"C:"; // Home directory path
        public static string csvPath = @"D:\workspace\aaa.csv"; // Shortcut file path
        public static string myToolPath = @"D:\workspace\sss";
        public static string dustBoxPath = @"DustBox";
        public static string safeBoxPath = @"SafeBox";
        public static string myFont = "瀬戸フォント";
        public static string fileOpenExe = "code";
        public static int btnSizeWidth = 80;
        public static int btnSizeHeight = 25;
        public static int labelSizeFont = 12;
        public static int btnSizeFont = 8;
        public static int tabSizeFont = 13;
        public static int textSizeFont = 13;
        public static int listSizeFont = 10;
        public static Color btnBackColor = Color.FromArgb(45, 45, 45);
        public static Color btnTextColor = Color.White;
        public static Color tabBackColor = Color.FromArgb(45, 45, 45);
        public static Color tabTextColor = Color.White;
        public static Color labelBackColor = Color.Transparent;
        public static Color labelTextColor = Color.White;
        public static Color comboBoxBackColor = Color.FromArgb(45, 45, 45);
        public static Color comboBoxTextColor = Color.White;
        public static Color contxtBackColor = Color.FromArgb(45, 45, 45);
        public static Color contxtTextColor = Color.White;
        public static Color menuBackColor = Color.FromArgb(45, 45, 45);
        public static Color menuTextColor = Color.White;
        public static Color dropDownBackColor = Color.FromArgb(45, 45, 45);
        public static Color dropDownTextColor = Color.White;
        public static Color listBoxBackColor = Color.FromArgb(30, 30, 30);
        public static Color listBoxTextColor = Color.White;
        public static Color txtBackColor = Color.FromArgb(30, 30, 30);
        public static Color txtTextColor = Color.White;
        public static Color formBackColor = Color.FromArgb(20, 20, 20);
        public static Color formTextColor = Color.White;
        public static void LoadSettings(string configFilePath)
        {
            if (!File.Exists(configFilePath))
            {
                MessageBox.Show(string.Format("設定ファイル '{0}' が見つかりません。デフォルト設定を使用します。",configFilePath));
                return;
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(configFilePath);

            XmlNodeList settings = xmlDoc.SelectNodes("/configuration/appSettings/add");
            NameValueCollection appSettings = new NameValueCollection();

            foreach (XmlNode setting in settings)
            {
                string key = setting.Attributes["key"].Value;
                string value = setting.Attributes["value"].Value;
                appSettings.Add(key, value);
            }

            homeDirectory = appSettings["homeDirectory"] ?? homeDirectory;
            csvPath = appSettings["csvPath"] ?? csvPath;
            myToolPath = appSettings["myToolPath"] ?? myToolPath;
            dustBoxPath = appSettings["dustBoxPath"] ?? dustBoxPath;
            safeBoxPath = appSettings["safeBoxPath"] ?? safeBoxPath;
            myFont = appSettings["myFont"] ?? myFont;
            fileOpenExe = appSettings["fileOpenExe"] ?? fileOpenExe;

            int width;
            btnSizeWidth = int.TryParse(appSettings["btnSizeWidth"], out width) ? width : btnSizeWidth;

            int height;
            btnSizeHeight = int.TryParse(appSettings["btnSizeHeight"], out height) ? height : btnSizeHeight;

            int labelFont;
            labelSizeFont = int.TryParse(appSettings["labelSizeFont"], out labelFont) ? labelFont : labelSizeFont;

            int btnFont;
            btnSizeFont = int.TryParse(appSettings["btnSizeFont"], out btnFont) ? btnFont : btnSizeFont;

            int tabFont;
            tabSizeFont = int.TryParse(appSettings["tabSizeFont"], out tabFont) ? tabFont : tabSizeFont;

            int textFont;
            textSizeFont = int.TryParse(appSettings["textSizeFont"], out textFont) ? textFont : textSizeFont;

            int listFont;
            listSizeFont = int.TryParse(appSettings["listSizeFont"], out listFont) ? listFont : listSizeFont;

            btnBackColor = ColorTranslator.FromHtml(appSettings["btnBackColor"] ?? "#2D2D2D");
            btnTextColor = ColorTranslator.FromHtml(appSettings["btnTextColor"] ?? "#FFFFFF");
            tabBackColor = ColorTranslator.FromHtml(appSettings["tabBackColor"] ?? "#2D2D2D");
            tabTextColor = ColorTranslator.FromHtml(appSettings["tabTextColor"] ?? "#FFFFFF");
            labelBackColor = ColorTranslator.FromHtml(appSettings["labelBackColor"] ?? "#00000000");
            labelTextColor = ColorTranslator.FromHtml(appSettings["labelTextColor"] ?? "#FFFFFF");
            comboBoxBackColor = ColorTranslator.FromHtml(appSettings["comboBoxBackColor"] ?? "#2D2D2D");
            comboBoxTextColor = ColorTranslator.FromHtml(appSettings["comboBoxTextColor"] ?? "#FFFFFF");
            contxtBackColor = ColorTranslator.FromHtml(appSettings["contxtBackColor"] ?? "#2D2D2D");
            contxtTextColor = ColorTranslator.FromHtml(appSettings["contxtTextColor"] ?? "#FFFFFF");
            menuBackColor = ColorTranslator.FromHtml(appSettings["menuBackColor"] ?? "#2D2D2D");
            menuTextColor = ColorTranslator.FromHtml(appSettings["menuTextColor"] ?? "#FFFFFF");
            dropDownBackColor = ColorTranslator.FromHtml(appSettings["dropDownBackColor"] ?? "#2D2D2D");
            dropDownTextColor = ColorTranslator.FromHtml(appSettings["dropDownTextColor"] ?? "#FFFFFF");
            listBoxBackColor = ColorTranslator.FromHtml(appSettings["listBoxBackColor"] ?? "#1E1E1E");
            listBoxTextColor = ColorTranslator.FromHtml(appSettings["listBoxTextColor"] ?? "#FFFFFF");
            txtBackColor = ColorTranslator.FromHtml(appSettings["txtBackColor"] ?? "#1E1E1E");
            txtTextColor = ColorTranslator.FromHtml(appSettings["txtTextColor"] ?? "#FFFFFF");
            formBackColor = ColorTranslator.FromHtml(appSettings["formBackColor"] ?? "#141414");
            formTextColor = ColorTranslator.FromHtml(appSettings["formTextColor"] ?? "#FFFFFF");

            // パスと色のマッピングを読み込む
            foreach (string key in appSettings.AllKeys)
            {
                if (key.StartsWith("/") || key.Contains(":")) // パスのキーを識別
                {
                    string colorValue = appSettings[key];
                    Color color;
                    try
                    {
                        color = ColorTranslator.FromHtml(colorValue);
                    }
                    catch
                    {
                        color = Color.Gray; // 無効なカラーコードの場合のデフォルト色
                    }
                    pathColorMapping[key] = color;
                }
            }
        }
    }

    static class CommonLibrary
    {
        public static void MoveDirectory(string sourcePath, string destinationPath)
        {
            // 移動先のディレクトリを作成
            Directory.CreateDirectory(destinationPath);

            // ソースディレクトリ内のすべてのファイルをコピー
            foreach (string file in Directory.GetFiles(sourcePath))
            {
                string destFile = Path.Combine(destinationPath, Path.GetFileName(file));
                File.Copy(file, destFile);
            }

            // ソースディレクトリ内のすべてのサブディレクトリを再帰的にコピー
            foreach (string dir in Directory.GetDirectories(sourcePath))
            {
                string destDir = Path.Combine(destinationPath, Path.GetFileName(dir));
                MoveDirectory(dir, destDir);
            }

            // 元のディレクトリを削除
            Directory.Delete(sourcePath, true);
        }
    }

    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // アプリケーションの実行ディレクトリを取得
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            // 相対パスを使用して設定ファイルのパスを指定
            string configFilePath = Path.Combine(appDirectory, @"Setting\setting.config");
            GlobalSettings.LoadSettings(configFilePath); // 設定を読み込む

            switch (args.Length)
            {
                case 0:
                    // 引数がない場合
                    break;
                case 1:
                    // 引数が1つの場合
                    string directoryPath = args[0];
                    if (Directory.Exists(directoryPath))
                    {
                        GlobalSettings.homeDirectory = directoryPath;
                    }
                    else
                    {
                        // ディレクトリが存在しない場合
                        Console.WriteLine(string.Format("ディレクトリ'{0}'は存在しません。プログラムを終了します。", directoryPath));
                        Environment.Exit(1); // プログラムを終了
                    }
                    break;
                default:
                    // 引数が2つ以上の場合
                    Console.WriteLine("引数の数が不正です。");
                    Environment.Exit(1); // プログラムを終了
                    break;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainForm mainForm = new MainForm();
            mainForm.OpenNewTab(GlobalSettings.homeDirectory); // 初期ディレクトリを指定
            Application.Run(mainForm);
        }
    }
}
