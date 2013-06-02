#VisualStudioController  
version 2013/06/01  

##説明
起動中のVisualStudioを外部から操作または情報を取得するコンソールアプリケーション  

##用途 & 目的
スクリプト言語や拡張機能に対応したその他エディターからVisualStudioを操作することを主な用途として作成しています  
なので使用するツールに依存するような機能は入れない予定です  
ちなみに私はvim上からVisualStudioを操作する目的で使用しています  
https://github.com/taku25/vim-visualstudio  


##動作確認
* Visual Studio 2010 / 2012  
 * EnvDTEを使用しているので2005以降なら動作するかもしれません  
 * Microsoft Visual Studio Express には非対応です
* Windows 7 32bit / 64 bit

##機能
###version 2013/06/01現在
* 対象のVisualStudioで編集中のソリューションをビルドする
* 対象のVisualStudioで編集中のソリューションをリビルドする
* 対象のVisualStudioで編集中のソリューションをクリーンする
* 対象のVisualStudioで編集中の指定したプロジェクトのみをビルドする
* 対象のVisualStudioで編集中の指定したプロジェクトのみをリビルドする
* 対象のVisualStudioで編集中の指定したプロジェクトのみをクリーンする
* 対象のVisualStudioで編集中の指定したファイルのみをコンパイルする
* 対象のVisualstudioで編集中の指定したプロジェクトにファイルを追加する
* 対象のVisualStudioがビルド中ならビルドをキャンセルする
* 対象のVisualStudioで編集中のソリューションを実行する
* 対象のVisualStuidoで編集中のソリューションをデバッグ実行する
* 対象のVisualStuidoで実行中のデバッグを停止する
* 対象のVisualStudioで現在編集中ファイルのファイル名(フルパス)をコンソールに出力する
* 対象のVisualStudioで編集中のソリューションに含まれるファイル名(フルパス)をすべてコンソールに出力する
* 対象のVisualStudioの出力ウインドウに表示されている内容をコンソールに出力する
* 対象のVisualStudioの検索結果ウインドウ1に表示されいている内容をコンソールに出力する
* 対象のVisualStudioの検索結果ウインドウ2に表示されいている内容をコンソールに出力する
* 対象のVisualStudioのシンボル検索結果ウインドウに表示されいている内容をコンソールに出力する
* 対象のVisualStudioのエラー一覧ウインドウに表示されいている内容をコンソールに出力する
 * Visual Studio 2005以上で有効 
* 対象のVisualstudioに含まれているファイル名(フルパス)と行数を指定してブレイクポイントを設定する
* 対象のVisualstudioで指定したファイルを開く
* 対象のVisualstudioで開いているソリューションを閉じる
* 対象のVisualstudioでカレントのビルド構成をコンソールに出力する
* 対象のVisualstudioで検索を行う
 * Visual Studio 2005以上で有効 

##使い方
    $ VisualStudioContoller.exe builld -t mysolution

##引数 & 説明
    Usage: VisualStudioController <commnad> <options> 
    version 2013/5/27
    <commnad>
    build                       : ビルド
    rebuild                     : リビルド
    clean                       : クリーン
    buildproject                : プロジェクトのビルド  強制的にwaitします
    rebuildproject              : プロジェクトのリビルド  強制的にwaitします
    cleanproject                : プロジェクトのクリーン  強制的にwaitします
    compilefile                 : ファイルのコンパイル
    cancelbuild                 : ビルドのキャンセル
    run                         : 実行
    debugrun                    : デバッグ実行
    stopdebugrun                : 実行のストップ
    find                        : 検索
    addfile                     : 対象のプロジェクトに既存ファイルを追加
    addbreakpoint               : ブレークポイントの追加
    openfile                    : ファイルを開く
    closesolution               : ソリューションを閉じる           
    getfile                     : 編集中ファイル取得" );
    getallfile                  : 全ファイル名をファイル取得" );
    getoutput                   : 出力Windowの中身を取得
    getfindresult1              : 検索結果1を取得
    getfindresult2              : 検索結果2を取得
    getfindsymbolresult         : シンボルの検索結果を取得
    geterrorlist                : エラー一覧の取得
    getbuildconfig              : 現在のビルド構成を取得
    
    <options>
    -[h]elp                     : ヘルプの表示
    -[t]arget                   : [-t SourceFilePath(fullpath) or -t SolutionName(name)] ターゲットソリューション名 か ターゲットソリューションに含まれているソースファイル名
                                : SolutionName(name)で指定する場合 名前の先頭一部でも有効です
    -[w]ait                     : 終わるまで待つ(build and rebuild時に有効)
    -[f]ile                     : 対象のソースファイル名(FullPath)を指定します. また-fに何も設定されていなかった場合かつ-tにSourceFilePathを設定していた場合はそれを使います
    -[p]roj                     : 対象のプロジェクト名を指定します 名前の先頭一部でも有効です
    -[l]ine                     : 行を指定します
    -[c]olumn                   : 列を指定します
    -findwhat[fw]               : 検索文字列の指定
    -findtarget[ft]             : 検索対象設定
                                : [project カレントプロジェクト(default)]
                                : [solution ソリューション]
    -enablefindmatchcase[efm]   : 大文字小文字判定有り無し
                                : [default 判定なし]
    -debugwrite[dw]             : 詳細情報も出力(debug用)
* 引数なし か helpオプションで上記のヘルプを見ることができます
