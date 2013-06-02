using System;
using System.Collections.Generic;
using System.Text;

namespace VisualStudioController {
    public class ArgsConvert {

        public enum CommandType
        {
            Build = 0,
            ReBuild,
            Clean,
            BuildProject,
            ReBuildProject,
            CleanProject,
            CompileFile,
            CancelBuild,
            Run,
            DebugRun,
            StopDebugRun,
            Find,
            GetFile,
            GetAllFiles,
            GetOutput,
            GetFindResult1,
            GetFindResult2,
            GetFindSymbolResult,
            GetErrorList,
            GetBuildConfig,
            GetProjectName,
            GetCurrnetProjectName,
            GetStartUpProjectName,
            GetSolutionName,
            OpenFile,
            CloseSolution,
            AddBreakPoint,
            AddFile,
            kMax,
        };
        
        public enum FindTargetType
        {
            Project = 0,
            Solution,
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
        private System.String targetProjectName_ = "";
        private int line_ = 1;
        private int column_ = 1;
        private bool showHelp_ = false;
        private System.String findWhat_ = "";
        private FindTargetType findTarget_ = FindTargetType.Project;
        private bool findMatchCase_ = false;

        public bool Commnad (CommandType commandType)
        {
            return commnadType_[(int)commandType];
        }

        public System.String TargetName { get { return targetName_; } }
        public bool IsWait { get { return isWait_; } }
        public System.String FileFullPath { get { return fileFullPath_; } }
        public System.String TargetProjectName { get { return targetProjectName_; } }

        public System.String FindWhat { get { return findWhat_; } }
        public FindTargetType FindTarget { get { return findTarget_; } }
        public bool FindMatchCase { get { return findMatchCase_; } }

        public int Line { get { return line_; } }
        public int Column { get { return column_; } }


        public bool ShowHelp
        {
            get { return showHelp_; }
        }

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
                }else if (args[0] == "buildproject"){
                    commnadType_[(int)CommandType.BuildProject] = true;
                }else if (args[0] == "rebuildproject"){
                    commnadType_[(int)CommandType.ReBuildProject] = true;
                }else if (args[0] == "cleanproject"){
                    commnadType_[(int)CommandType.CleanProject] = true;
                }else if (args[0] == "compilefile"){
                    commnadType_[(int)CommandType.CompileFile] = true;
                }else if (args[0] == "getfile"){
                    commnadType_[(int)CommandType.GetFile] = true;
                }else if (args[0] == "getoutput"){
                    commnadType_[(int)CommandType.GetOutput] = true;
                }else if (args[0] == "getfindresult1"){
                    commnadType_[(int)CommandType.GetFindResult1] = true;
                }else if (args[0] == "getfindresult2"){
                    commnadType_[(int)CommandType.GetFindResult2] = true;
                }else if (args[0] == "getfindsymbolresult"){
                    commnadType_[(int)CommandType.GetFindSymbolResult] = true;
                }else if (args[0] == "getallfiles"){
                    commnadType_[(int)CommandType.GetAllFiles] = true;
                }else if (args[0] == "addbreakpoint"){
                    commnadType_[(int)CommandType.AddBreakPoint] = true;
                }else if (args[0] == "geterrorlist"){
                    commnadType_[(int)CommandType.GetErrorList] = true;
                }else if (args[0] == "openfile"){
                    commnadType_[(int)CommandType.OpenFile] = true;
                }else if (args[0] == "cancelbuild"){
                    commnadType_[(int)CommandType.CancelBuild] = true;
                }else if (args[0] == "getbuildconfig"){
                    commnadType_[(int)CommandType.GetBuildConfig] = true;
                }else if (args[0] == "closesolution"){
                    commnadType_[(int)CommandType.CloseSolution] = true;
                }else if (args[0] == "stopdebugrun"){
                    commnadType_[(int)CommandType.StopDebugRun] = true;
                }else if (args[0] == "find"){
                    commnadType_[(int)CommandType.Find] = true;
                }else if (args[0] == "addfile"){
                    commnadType_[(int)CommandType.AddFile] = true;
                }else if (args[0] == "getprojectname"){
                    commnadType_[(int)CommandType.GetProjectName] = true;
                }else if (args[0] == "getcurrentprojectname"){
                    commnadType_[(int)CommandType.GetCurrnetProjectName] = true;
                }else if (args[0] == "getsolutionname"){
                    commnadType_[(int)CommandType.GetSolutionName] = true;
                }else if (args[0] == "getstartupprojectname"){
                    commnadType_[(int)CommandType.GetStartUpProjectName] = true;
                }

                
                
                for(int i = 1; i < args.Length; i++){
                    if(args[i].ToLower() == "-target" || args[i].ToLower() == "-t"){
                        targetName_ = args[i + 1];
                        i+=1;
                    }else if(args[i].ToLower() == "-wait" || args[i].ToLower() == "-w"){
                        isWait_ = true;
                    }else if (args[i].ToLower() == "-debugwrite" || args[i].ToLower() == "-dw"){
                        ConsoleWriter.DebugEnable = true;
                    }else if (args[i].ToLower () == "-file" || args[i].ToLower() == "-f"){
                        fileFullPath_ = args[i + 1];
                        i+=1;
                    }else if (args[i].ToLower () == "-line" || args[i].ToLower () == "-l"){
                        line_ = System.Convert.ToInt32(args[i + 1]);
                        i+=1;
                    }else if (args[i].ToLower () == "-column" || args[i].ToLower () == "-c"){
                        column_ = System.Convert.ToInt32(args[i + 1]);
                        i+=1;
                    }else if (args[i].ToLower () == "-help" || args[i].ToLower() == "-h"){
                        showHelp_ = true;
                    }else if (args[i].ToLower() == "-proj" || args[i].ToLower() == "-p"){
                        this.targetProjectName_ = args[i + 1];
                        i+=1; 
                    }else if (args[i].ToLower() == "-findwhat" || args[i].ToLower() == "-fw"){
                        findWhat_ = args[i + 1];
                        i+=1; 
                    }else if (args[i].ToLower() == "-findtarget" || args[i].ToLower() == "-ft"){
                        if(args[i + 1] == "solution"){
                            findTarget_ = FindTargetType.Solution;
                        }
                        i+=1; 
                    }else if (args[i].ToLower() == "-enablefindmatchcase" || args[i].ToLower() == "-efm"){
                        findMatchCase_ = true; 
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
            ConsoleWriter.WriteLine ("version 2013/06/01");
            ConsoleWriter.WriteLine ("<commnad>");
            ConsoleWriter.WriteLine ("build                       : ビルド");
            ConsoleWriter.WriteLine ("rebuild                     : リビルド");
            ConsoleWriter.WriteLine ("clean                       : クリーン");
            ConsoleWriter.WriteLine ("buildproject                : プロジェクトのビルド  強制的にwaitします");
            ConsoleWriter.WriteLine ("rebuildproject              : プロジェクトのリビルド  強制的にwaitします");
            ConsoleWriter.WriteLine ("cleanproject                : プロジェクトのクリーン  強制的にwaitします");
            ConsoleWriter.WriteLine ("compilefile                 : ファイルのコンパイル");
            ConsoleWriter.WriteLine ("cancelbuild                 : ビルドのキャンセル");
            ConsoleWriter.WriteLine ("run                         : 実行");
            ConsoleWriter.WriteLine ("debugrun                    : デバッグ実行");
            ConsoleWriter.WriteLine ("stopdebugrun                : 実行のストップ");
            ConsoleWriter.WriteLine ("find                        : 検索");
            ConsoleWriter.WriteLine ("addfile                     : 対象のプロジェクトに既存ファイルを追加");
            ConsoleWriter.WriteLine ("addbreakpoint               : ブレークポイントの追加");
            ConsoleWriter.WriteLine ("openfile                    : ファイルを開く");
            ConsoleWriter.WriteLine ("closesolution               : ソリューションを閉じる");           
            ConsoleWriter.WriteLine ("getfile                     : 編集中ファイル取得");
            ConsoleWriter.WriteLine ("getallfiles                 : 全ファイル名をファイル取得");
            ConsoleWriter.WriteLine ("getoutput                   : 出力Windowの中身を取得");
            ConsoleWriter.WriteLine ("getfindresult1              : 検索結果1を取得");
            ConsoleWriter.WriteLine ("getfindresult2              : 検索結果2を取得");
            ConsoleWriter.WriteLine ("getfindsymbolresult         : シンボルの検索結果を取得");
            ConsoleWriter.WriteLine ("geterrorlist                : エラー一覧の取得");
            ConsoleWriter.WriteLine ("getbuildconfig              : 現在のビルド構成を取得");
            ConsoleWriter.WriteLine ("getprojectname              : 対象のファイルが含まれているプロジェクト名を取得");
            ConsoleWriter.WriteLine ("getcurrentprojectname       : カレントプロジェクト名を取得");
            ConsoleWriter.WriteLine ("getstartupprojectname       : スタートアッププロジェクト名を取得");
            ConsoleWriter.WriteLine ("getsolutionname             : 対象のファイルorプロジェクトが含まれているプロジェクト名を取得");
            ConsoleWriter.WriteLine ("<options>");
            ConsoleWriter.WriteLine ("-[h]elp                     : ヘルプの表示");
            ConsoleWriter.WriteLine ("-[t]arget                   : [SourceFilePath(fullpath) or ProjectName(name) or SolutionName(name)] ソリューション名、プロジェクト名かソリューションに含まれているソースファイル名");
            ConsoleWriter.WriteLine ("                            : SolutionName(name) or ProjectName(name)で指定する場合 名前の先頭一部でも有効です");
            ConsoleWriter.WriteLine ("-[w]ait                     : 終わるまで待つ(build and rebuild時に有効)");
            ConsoleWriter.WriteLine ("-[f]ile                     : 対象のソースファイル名(FullPath)を指定します. また-fに何も設定されていなかった場合かつ-tにSourceFilePathを設定していた場合はそれを使います");
            ConsoleWriter.WriteLine ("-[p]roj                     : 対象のプロジェクト名を指定します 名前の先頭一部でも有効です");
            ConsoleWriter.WriteLine ("-[l]ine                     : 行を指定します");
            ConsoleWriter.WriteLine ("-[c]olumn                   : 列を指定します");
            ConsoleWriter.WriteLine ("-findwhat[fw]               : 検索文字列の指定");
            ConsoleWriter.WriteLine ("-findtarget[ft]             : 検索対象設定");
            ConsoleWriter.WriteLine ("                            : [project カレントプロジェクト(default)]");
            ConsoleWriter.WriteLine ("                            : [solution ソリューション]");
            ConsoleWriter.WriteLine ("-enablefindmatchcase[efm]   : 大文字小文字判定有り無し");
            ConsoleWriter.WriteLine ("                            : [default 判定なし]");
            ConsoleWriter.WriteLine ("-debugwrite[dw]             : 詳細情報も出力(debug用)");
        }
    }
}
