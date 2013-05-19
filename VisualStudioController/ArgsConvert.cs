using System;
using System.Collections.Generic;
using System.Text;

namespace VisualStudioController {
	public class ArgsConvert {

		public enum CommandType
		{
            Build = 0,
            Clean,
            ReBuild,
            Run,
            DebugRun,
            GetFile,
            GetAllFile,
            GetOutput,
            GetFindResult1,
            GetFindResult2,
            GetFindSimbolResult,
            GetErrorList,
            
            AddBreakPoint,
            kMax,
		};

        public ArgsConvert ()
        {
            for(int i = 0; i < (int)CommandType.kMax; i++){
                commnadType_[i] = false;
            }
        }

        private bool [] commnadType_ = new bool[(int)CommandType.kMax];
        private bool isWait_ = false;
        private System.String targetName_ = "";
        private System.String fileFullPath_ = "";
        private int line_ = 1;
        private int column_ = 1;
        public bool Commnad (CommandType commandType)
        {
            return commnadType_[(int)commandType];
        }

        public System.String TargetName { get { return targetName_; } }
        public bool IsWait { get { return isWait_; } }
        public System.String FileFullPath { get { return fileFullPath_; } }

        public int Line { get { return line_; } }
        public int Column { get { return column_; } }


        public bool Analysis (string[] args)
        {
            if(Analysis_ (args) == false){
                ArgsInfo ();
                return false;
            }
            return true;
        }
        
        private bool Analysis_ (string[] args)
        {
            try{
                if(args.Length <= 0){
                    return false;
                }
                if (args[0] == "build"){
                    commnadType_[(int)CommandType.Build] = true;
                }else if (args[0] == "run"){
                    commnadType_[(int)CommandType.Run] = true;
                }else if (args[0] == "debugrun"){
                    commnadType_[(int)CommandType.DebugRun] = true;
                }else if (args[0] == "clean"){
                    commnadType_[(int)CommandType.Clean] = true;
                }else if (args[0] == "rebuild"){
                    commnadType_[(int)CommandType.ReBuild] = true;
                }else if (args[0] == "getfile"){
                    commnadType_[(int)CommandType.GetFile] = true;
                }else if (args[0] == "getoutput"){
                    commnadType_[(int)CommandType.GetOutput] = true;
                }else if (args[0] == "getfindresult1"){
                    commnadType_[(int)CommandType.GetFindResult1] = true;
                }else if (args[0] == "getfindresult2"){
                    commnadType_[(int)CommandType.GetFindResult2] = true;
                }else if (args[0] == "getfindsymbolresult"){
                    commnadType_[(int)CommandType.GetFindSimbolResult] = true;
                }else if (args[0] == "getallfile"){
                    commnadType_[(int)CommandType.GetAllFile] = true;
                }else if (args[0] == "addbreakpoint"){
                    commnadType_[(int)CommandType.AddBreakPoint] = true;
                }else if (args[0] == "geterrorlist"){
                    commnadType_[(int)CommandType.GetErrorList] = true;
                }
            
                for(int i = 1; i < args.Length; i++){
                    if(args[i].ToLower() == "-t"){
                        targetName_ = args[i + 1];
                        i+=1;
                    }else if(args[i].ToLower() == "-w"){
                        isWait_ = true;
                    }else if (args[i].ToLower() == "-d"){
                        ConsoleWriter.DebugEnable = true;
                    }else if (args[i].ToLower () == "-f"){
                        fileFullPath_ = args[i + 1];
                        i+=1;
                    }else if (args[i].ToLower () == "-line"){
                        line_ = System.Convert.ToInt32(args[i + 1]);
                        i+=1;
                    }else if (args[i].ToLower () == "-column"){
                        column_ = System.Convert.ToInt32(args[i + 1]);
                        i+=1;
                    }
                }

                foreach(String str in args){
                    ConsoleWriter.WriteDebugLine(str);
                }

                if(commnadType_[(int)CommandType.AddBreakPoint] == true && System.String.IsNullOrEmpty(fileFullPath_) == true){
                    fileFullPath_ = targetName_;
                    if(System.String.IsNullOrEmpty(fileFullPath_) == true){
                        ConsoleWriter.WriteDebugLine(@"ブレイクポイントを設定するファイルが指定されていません");
                        return false;
                    }
                }

            }catch(System.Exception e){
                ConsoleWriter.WriteLine(e.ToString());
                return false;
            }
            return true;
        }

        public void ArgsInfo ()
        {
            ConsoleWriter.WriteLine ("Usage: VisualStudioController <commnad> <options> ");
            ConsoleWriter.WriteLine ("version 0.2");
            ConsoleWriter.WriteLine ("<commnad>:");
            ConsoleWriter.WriteLine ("build               : ビルド");
            ConsoleWriter.WriteLine ("rebuild             : リビルド");
            ConsoleWriter.WriteLine ("clean               : クリーン");
            ConsoleWriter.WriteLine ("run                 : 実行");
            ConsoleWriter.WriteLine ("debugrun            : デバッグ実行");
            ConsoleWriter.WriteLine ("getfile             : 編集中ファイル取得" );
            ConsoleWriter.WriteLine ("getallfile          : 全ファイル名をファイル取得" );
            ConsoleWriter.WriteLine ("getoutput           : 出力Windowの中身を取得");
            ConsoleWriter.WriteLine ("getfindresult1      : 検索結果1を取得");
            ConsoleWriter.WriteLine ("getfindresult2      : 検索結果2を取得");
            ConsoleWriter.WriteLine ("geterrorlist        : エラー一覧の取得");
            ConsoleWriter.WriteLine ("addbreakpoint       : ブレークポイントの追加");
            //ConsoleWriter.WriteLine ("getfindsimbolresult : シンボルの検索結果を取得");
            ConsoleWriter.WriteLine ("<options>           : ");
            ConsoleWriter.WriteLine ("-t                  : [-t SourceFilePath(fullpath) or -t SolutionName(name)] ターゲットソリューション名 か ターゲットソリューションに含まれているソースファイル名");
            ConsoleWriter.WriteLine ("-w                  : 終わるまで待つ(build and rebuild時に有効)");
            ConsoleWriter.WriteLine ("-d                  : 詳細情報も出力(debug用)");
            ConsoleWriter.WriteLine ("-f                  : BreakPointを設定したい対象のファイル名(FullPath) -tにSourceFilePathを設定していた場合はそれを使います");
            ConsoleWriter.WriteLine ("-line               : 行を指定します");
            ConsoleWriter.WriteLine ("-column             : 列を指定します");
            //ConsoleWriter.WriteLine ("-enc                : cp932 utf8(default) コンソール出力を行うエンコード指定");
        }
    }
}
