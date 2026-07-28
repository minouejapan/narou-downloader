# narou-downloader

na6dlは、小説家になろうおよび姉妹サイトで公開されている小説を青空文庫形式のテキストファイルでダウンロードするためのツールです。

## GUI版（Windows 10/11）

Windows向けのGUI版を `NarouDownloaderGui` フォルダーに収録しています。作品URL、保存先、ファイル名、開始話を画面上で指定でき、実行ログの確認や途中停止も行えます。

### 一般ユーザー向け：portable版

1. GitHubの[Releases](../../releases)から `NarouDownloaderGui-win-x64.zip` をダウンロードします。
2. ZIPを任意のフォルダーへ展開します。
3. `NarouDownloaderGui.exe` をダブルクリックします。

インストールや.NETランタイムの追加導入は不要です。公式リリース `ver5.9.1.0` の `na6dl.exe` と、実行に必要な.NET／Windows App SDKファイルを同梱しています。

Windows SmartScreenが表示された場合は、発行元を確認したうえで「詳細情報」→「実行」を選択してください。現在のportable版にはコード署名がありません。

### 開発者向け

.NET 10 SDKとWinUI 3開発環境が必要です。

```powershell
dotnet build .\NarouDownloaderGui\NarouDownloaderGui.csproj -c Debug -p:Platform=x64
dotnet run --project .\NarouDownloaderGui\NarouDownloaderGui.csproj -c Debug -p:Platform=x64
```

自己完結型portable版のZIPは次のコマンドで作成できます。

```powershell
.\scripts\publish-gui.ps1
```

成果物は `artifacts\NarouDownloaderGui-win-x64.zip` に出力されます。`gui-v*` 形式のタグをpushすると、GitHub Actionsが同じZIPをビルドしてReleaseへ添付します。

## CLI版

na6dlは小説家になろうおよび姉妹サイトで公開されている小説を青空文庫形式のテキストファイルでダウンロードするためのツールです。<br>
URLがhttps://ncode.syosetu.com/およびhttps://novel18.syosetu.com/で始まる作品をダウンロードすることが出来ます。<br>

### 動作環境
Windows10/11上のコマンドプロンプト上で動作します。

### 実行ファイルの作り方
* Delphi (XE2以降)の場合：na6dl.dprojを開いてビルドしてください。尚、ビルドするためにはTregExprライブラリが必要です。
* Lazarus（3.6以降）の場合：na6dl.lpiを開いてビルドして下さい。尚、ビルドするためにはTregExprライブラリとSimpleHTMLParserが必要です。Lazarusの場合はWindows上でクロスコンパイルするかLinux上のLazarusでビルドすることでLinux用の実行ファイルも出力可能です。
  * TregExprライブラリ：https://github.com/andgineer/TRegExpr
  * SimpleHTMLParser：https://github.com/minouejapan/SimpleHTMLParser

### 使い方
コマンドプロンプト上で、<br>
na6dl ダウンロードしたいなろう系小説トップページのURL (保存したいテキストファイル名)<br>
と入力して実行キーを押します。正常に実行されればna6dl.exeがあるフォルダにダウンロードした小説が青空文庫形式のテキストファイルで保存されます。<br>

尚、保存したファイル名の指定は省略できます。省略した場合はダウンロードした小説のタイトル名からファイル名を作成して保存します。<br>

### 禁止事項
1. na6dlを用いてWeb小 説サイトからダウンロードしたテキストファイルの第三者への販売や不特定多数への配信。 
2. ダウンロードしたオリジナル作品を著作者の了解なく加工（文章の流用や作品の翻訳等）しての再公開。 
3. その他、著作者の権利を踏みにじるような行為。 


### ライセンス
MIT
