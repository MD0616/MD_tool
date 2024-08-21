MainForm.cs: メインフォームの定義
Program.cs: アプリケーションのエントリーポイント
UIComponents.cs: UIコンポーネントの初期化

コンパイル　 csc.exe /out:MD_Explorer.exe /reference:"C:\Program Files (x86)\Reference Assemblies\Microsoft\WindowsPowerShell\3.0\System.Management.Automation.dll" MainForm.cs Program.cs Init.cs CompDef.cs ActDef.cs DrawDesign.cs