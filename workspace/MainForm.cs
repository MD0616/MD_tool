
using System; // 基本的な型やコア機能を提供する .NET の基本クラスを使用するための名前空間
using System.Diagnostics; // プロセス、イベントログ、パフォーマンスカウンターなどの機能を提供するクラスを使用するための名前空間
using System.Drawing; // GDI+ 基本グラフィック機能を使用するための名前空間
using System.Windows.Forms; // Windows フォームアプリケーションを作成するためのクラスを使用するための名前空間
using System.IO; // ファイルとデータ ストリームの読み書きを可能にする型を使用するための名前空間
using System.Text;
using Microsoft.Win32.SafeHandles;

// MD_Explorerという名前空間。同じ名前空間内の型は互いに名前で直接アクセスできます。
namespace MD_Explorer
{
    // MainFormという名前のクラスを定義。このクラスはFormクラスを継承しています。'partial'キーワードは、このクラスの定義が複数のファイルに分割されていることを示します。
    public partial class MainForm : Form
    {
        private string homeDirectory = "C:"; // ホームディレクトリのパスを保持するプライベートフィールド。
        private string csvPath = "D:\\workspace\\aaa.csv"; // ショートカットの記載パス
        private string myToolPath = "D:\\workspace\\sss";
        private string myFont = "瀬戸フォント";
        private string fileOpenExe = "code";
        private int btnSizeWidth = 80;
        private int btnSizeHeight = 25;
        private int btnSizeFont = 8;
        private int tabSizeFont = 13;
        private int textSizeFont = 13;
        private TextBox txtSearchBar; // 検索バーを表すテキストボックス
        private Button btnSearch; // 検索を開始するボタン
        private TabControl tabControl1; // タブを管理するコントロール
        private Button btnHome; // ホームディレクトリに移動するボタン
        private Button btnRefresh; // ディレクトリの表示を更新するボタン
        private Button btnExplorer; // ファイルエクスプローラを開くボタン
        private Button btnTerminal; // ターミナルを開くボタン
        private Button btnCopyPath; // 現在のパスをコピーするボタン
        private Button btnShortcut;
        private Button btnNewFile; // 新しいファイルを作成するボタン
        private Button btnNewFolder; // 新しいフォルダを作成するボタン
        private TextBox txtPowerShellOutput; // PowerShellの出力を表示するテキストボックス
        private TextBox txtPowerShellInput; // PowerShellの入力を受け取るテキストボックス
        private Process powerShellProcess; // PowerShellプロセスを管理するためのプロセスオブジェクト
        private Button btnExecute;
        private ComboBox scriptComboBox;
        private ContextMenuStrip contextMenu;
        private ContextMenuStrip dropDownMenu;

        // MainFormクラスのコンストラクタ。オブジェクトが生成されるときに呼び出されます。
        public MainForm() 
        {
            InitializeComponent(); // フォーム上のコントロールの初期化を行います。このメソッドはデザイナによって自動的に生成されます。
        }

        // HomeDirectoryプロパティ。homeDirectoryフィールドの値を取得または設定します。
        public string HomeDirectory 
        {
            get { return homeDirectory; } // getアクセサー。HomeDirectoryプロパティの値を取得するために使用します。
            set { homeDirectory = value; } // setアクセサー。HomeDirectoryプロパティの値を設定するために使用します。
        }
    }
}
