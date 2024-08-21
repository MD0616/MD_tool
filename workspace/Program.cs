using System;
using System.Windows.Forms;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace MD_Explorer
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainForm mainForm = new MainForm();
            mainForm.OpenNewTab(mainForm.HomeDirectory); // 初期ディレクトリを指定
            Application.Run(mainForm);
        }
    }
}
