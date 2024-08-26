using System;
using System.IO;
using System.Windows.Forms;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace MD_Explorer
{
    // Static class to hold global-like variables
    static class GlobalSettings
    {
        public static string homeDirectory = "C:"; // Home directory path
        public static string csvPath = "D:\\workspace\\aaa.csv"; // Shortcut file path
        public static string myToolPath = "D:\\workspace\\sss";
        public static string myFont = "瀬戸フォント";
        public static string fileOpenExe = "code";
        public static int btnSizeWidth = 80;
        public static int btnSizeHeight = 25;
        public static int labelSizeFont = 12;
        public static int btnSizeFont = 8;
        public static int tabSizeFont = 13;
        public static int textSizeFont = 13;
    }

    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
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
                        Console.WriteLine(string.Format("ディレクトリ'{0}'は存在しません。プログラムを終了します。",directoryPath));
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
