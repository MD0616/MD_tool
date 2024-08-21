# 最新の.NET Frameworkバージョンを判断し、環境変数PATHに設定するシェルスクリプト

try {
    # Microsoft.PowerShell.Managementモジュールをインポート
    # このモジュールにはGet-ChildItemコマンドが含まれています
    Import-Module Microsoft.PowerShell.Management

    # 最新の.NET Frameworkバージョンを取得
    # Get-ChildItemコマンドを使用して、Frameworkディレクトリ内のバージョンを取得し、降順にソートして最初の1つを選択
    $latestVersion = (Get-ChildItem "$env:WINDIR\Microsoft.NET\Framework" | Sort-Object Name -Descending | Select-Object -First 1).Name

    # 環境変数PATHを更新
    # 最新の.NET Frameworkバージョンのパスを環境変数PATHに追加
    [System.Environment]::SetEnvironmentVariable("PATH", "$env:WINDIR\Microsoft.NET\Framework\$latestVersion;$env:PATH", [System.EnvironmentVariableTarget]::Machine)

    # 成功メッセージを出力
    Write-Output "PATH has been updated to the latest .NET Framework version."
} catch {
    # エラーメッセージを出力
    # 例外が発生した場合にエラーメッセージを表示
    Write-Error "An error occurred: $_"
}
