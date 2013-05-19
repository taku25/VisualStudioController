#VisualStudioController  
version 2013/5/17  

##説明
起動中のVisualStudioを外部から操作または情報を取得するコンソールアプリケーション  

##用途 & 目的
スクリプト言語や拡張機能に対応したその他エディターからVisualStudioを操作することを主な用途として作成しています  
なので使用するツールに依存するような機能は入れない予定です  
作者は普段vim上からVisualStudioを操作する目的で使用しています  
https://github.com/taku25/vim-visualstudio  


##動作確認
* Visual Studio 2010 / 2012  
 * EnvDTEを使用しているので2005以降なら動作するかもしれません  
 * Microsoft Visual Studio Express には非対応です
* Windows 7 32bit / 64 bit

##機能
###version 2013/5/17現在
* 対象のVisualStudioで編集中のソリューションをビルドする
* 対象のVisualStudioで編集中のソリューションをクリーンする
* 対象のVisualStudioで編集中のソリューションを実行する
* 対象のVisualStuidoで編集中のソリューションをデバッグ実行する
* 対象のVisualStudioで現在編集中ファイルのファイル名(フルパス)をコンソールに表示する
* 対象のVisualStudioで編集中のソリューションに含まれるファイル名(フルパス)をすべてコンソールに表示する
* 対象のVisualStudioの出力ウインドウに表示されている内容をコンソールに出力する
* 対象のVisualStudioの検索結果ウインドウ1に表示されいている内容をコンソールに出力する
* 対象のVisualStudioの検索結果ウインドウ2に表示されいている内容をコンソールに出力する
* 対象のVisualStudioのエラー一覧ウインドウに表示されいている内容をコンソールに出力する
 * Visual Studio 2005以上で有効 
* 対象のVisualstudioに含まれているファイル名(フルパス)と行数を指定してブレイクポイントを設定する



##使い方
    $ VisualStudioContoller.exe builld -t mysolution

##引数 & 説明
    Usage: VisualStudioController <commnad> <options>                    
    <commnad>
    build               : ビルド
    rebuild             : リビルド
    clean               : クリーン
    run                 : 実行
    debugrun            : デバッグ実行
    getfile             : 編集中ファイル取得
    getallfile          : 全ファイル名をファイル取得
    getoutput           : 出力Windowの中身を取得
    getfindresult1      : 検索結果1を取得
    getfindresult2      : 検索結果2を取得
    geterrorlist        : エラー一覧の取得
    addbreakpoint       : ブレークポイントの追加
    <options>           :
    -t                  : [-t SourceFilePath(fullpath) or -t SolutionName(name)] ターゲットソリューション名 か ターゲットソリューションに含まれているソースファイル名
    -w                  : 終わるまで待つ(build and rebuild時に有効)
    -d                  : 詳細情報も出力(debug用)
    -f                  : BreakPointを設定したい対象のファイル名(FullPath) -tにSourceFilePathを設定していた場合はそれを使います
    -column             : 列を指定します
    -line               : 行を指定します
* 引数なし実行で以上のヘルプを見ることができます
