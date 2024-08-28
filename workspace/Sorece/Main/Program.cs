using System;
using System.IO;
using System.Windows.Forms;
using System.Text;
using System.Xml; // XmlDocument と XmlNodeList のため
using System.Collections.Specialized; // NameValueCollection のため

namespace MD_Explorer
{
    // Static class to hold global-like variables
    static class GlobalSettings
    {
        public static string homeDirectory = @"C:"; // Home directory path
        public static string csvPath = @"D:\workspace\aaa.csv"; // Shortcut file path
        public static string myToolPath = @"D:\workspace\sss";
        public static string myFont = "瀬戸フォント";
        public static string fileOpenExe = "code";
        public static int btnSizeWidth = 80;
        public static int btnSizeHeight = 25;
        public static int labelSizeFont = 12;
        public static int btnSizeFont = 8;
        public static int tabSizeFont = 13;
        public static int textSizeFont = 13;

        public static void LoadSettings(string configFilePath)
        {
            if (!File.Exists(configFilePath))
            {
                Console.WriteLine(string.Format("設定ファイル '{0}' が見つかりません。デフォルト設定を使用します。",configFilePath));
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
