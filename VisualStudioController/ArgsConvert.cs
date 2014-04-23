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
            GetCurrentBuildConfig,
            GetBuildConfigList,
            GetBuildPlatformList,
            GetProjectName,
            GetCurrnetProjectName,
            GetStartUpProjectName,
            GetSolutionName,
            GetSolutionFileName,
            GetSolutionFullPath,
            GetBuildStatus,
            OpenFile,
            CloseSolution,
            AddBreakPoint,
            AddFile,
        };
        
        public enum FindTargetType
        {
            Project = 0,
            Solution,
            kMax,
        };

        public ArgsConvert ()
        {
            foreach(CommandType command in Enum.GetValues(typeof(CommandType))){
                commnadType_[(int)command] = false;
            }

            commandHelpArray_[(int)CommandType.Build]                   = "ビルド";
            commandHelpArray_[(int)CommandType.ReBuild]                 = "リビルド";
            commandHelpArray_[(int)CommandType.Clean]                   = "クリーン";
            commandHelpArray_[(int)CommandType.BuildProject]            = "プロジェクトのビルド  強制的にwaitします";
            commandHelpArray_[(int)CommandType.ReBuildProject]          = "プロジェクトのリビルド  強制的にwaitします";
            commandHelpArray_[(int)CommandType.CleanProject]            = "プロジェクトのクリーン  強制的にwaitします";
            commandHelpArray_[(int)CommandType.CompileFile]             = "ファイルのコンパイル";
            commandHelpArray_[(int)CommandType.CancelBuild]             = "ビルドのキャンセル";
            commandHelpArray_[(int)CommandType.Run]                     = "実行";
            commandHelpArray_[(int)CommandType.DebugRun]                = "デバッグ実行";
            commandHelpArray_[(int)CommandType.StopDebugRun]            = "実行の停止";
            commandHelpArray_[(int)CommandType.Find]                    = "検索";
            commandHelpArray_[(int)CommandType.GetFile]                 = "編集中ファイル取得";
            commandHelpArray_[(int)CommandType.GetAllFiles]             = "全ファイル名をファイル取得";
            commandHelpArray_[(int)CommandType.GetOutput]               = "出力ウインドの内容を取得";
            commandHelpArray_[(int)CommandType.GetFindResult1]          = "検索結果ウインド1の結果を取得";           
            commandHelpArray_[(int)CommandType.GetFindResult2]          = "検索結果ウインド2の結果を取得";
            commandHelpArray_[(int)CommandType.GetFindSymbolResult]     =  "未実装";
            commandHelpArray_[(int)CommandType.GetErrorList]            = "エラーウインド一覧の取得";
            commandHelpArray_[(int)CommandType.GetCurrentBuildConfig]   =  "検索結果1を取得";
            commandHelpArray_[(int)CommandType.GetBuildConfigList]      =  "検索結果2を取得";
            commandHelpArray_[(int)CommandType.GetBuildPlatformList]    =  "シンボルの検索結果を取得";
            commandHelpArray_[(int)CommandType.GetProjectName]          =  "プロジェクト名を取得";
            commandHelpArray_[(int)CommandType.GetCurrnetProjectName]   =  "カレントのプロジェクト名を取得";
            commandHelpArray_[(int)CommandType.GetStartUpProjectName]   =  "スターっとアッププロジェクト名を取得";
            commandHelpArray_[(int)CommandType.GetSolutionName]         =  "ソリューションの名前を取得";
            commandHelpArray_[(int)CommandType.GetSolutionFileName]     =  "ソリューションのファイル名を取得";
            commandHelpArray_[(int)CommandType.GetSolutionFullPath]     =  "ソリューションのフルパス取得";
            commandHelpArray_[(int)CommandType.GetBuildStatus]          =  "ビルド状況の取得";
            commandHelpArray_[(int)CommandType.OpenFile]                =  "ファイルを開く";
            commandHelpArray_[(int)CommandType.CloseSolution]           =  "ソリューションを閉じる";
            commandHelpArray_[(int)CommandType.AddBreakPoint]           = "ブレークポイントの追加";
            commandHelpArray_[(int)CommandType.AddFile]                 = "ファイルの追加";
        }

        private bool [] commnadType_ = new bool[Enum.GetValues(typeof(CommandType)).Length];
        private System.String [] commandHelpArray_ = new string[Enum.GetValues(typeof(CommandType)).Length];

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
            if(Analysis_ (args) == false || ShowHelp == true){
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

                foreach(CommandType command in Enum.GetValues(typeof(CommandType))){
                    if (args[0] == command.ToString().ToLower()){
                        commnadType_[(int)command] = true;
                        break;
                    }
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


                fileFullPath_ = ConvertPathSeparator(fileFullPath_);
                targetName_ = ConvertPathSeparator(targetName_);

            }catch(System.Exception e){
                ConsoleWriter.WriteLine(e.ToString());
                return false;
            }
            return true;
        }

        private System.String ConvertPathSeparator(System.String valu)
        {
            System.String temp = valu;
            if(System.String.IsNullOrEmpty(temp) == false){
                if (System.IO.Path.DirectorySeparatorChar != '/'){
                    temp = valu.Replace('/', System.IO.Path.DirectorySeparatorChar);
                }else if (System.IO.Path.DirectorySeparatorChar != '\\'){
                    temp = valu.Replace('\\', System.IO.Path.DirectorySeparatorChar);
                }   
            }
            
            return temp;
        }

        public void ArgsInfo ()
        {
            
            ConsoleWriter.WriteLine ("Usage: VisualStudioController <commnad> <options> ");
            ConsoleWriter.WriteLine ("version 2014/04/22");
            ConsoleWriter.WriteLine ("<commnad>");

            foreach(CommandType command in Enum.GetValues(typeof(CommandType))){
                ConsoleWriter.Write(command.ToString().ToLower());
                int commandLength = command.ToString().Length;
                //とりあえずスペースは30個適当
                for(int i = 30 - commandLength; i > 0; i--){ 
                    ConsoleWriter.Write(" ");
                }
                ConsoleWriter.WriteLine(": " + commandHelpArray_[(int)command]);
            }

            ConsoleWriter.WriteLine ("<options>");
            ConsoleWriter.WriteLine ("-[h]elp                       : ヘルプの表示");
            ConsoleWriter.WriteLine ("-[t]arget                     : [SourceFilePath(fullpath) or ProjectName(name) or SolutionName(name)] ソリューション名、プロジェクト名かソリューションに含まれているソースファイル名");
            ConsoleWriter.WriteLine ("                              : SolutionName(name) or ProjectName(name)で指定する場合 名前の先頭一部でも有効です");
            ConsoleWriter.WriteLine ("-[w]ait                       : 終わるまで待つ(build and rebuild時に有効)");
            ConsoleWriter.WriteLine ("-[f]ile                       : 対象のソースファイル名(FullPath)を指定します. また-fに何も設定されていなかった場合かつ-tにSourceFilePathを設定していた場合はそれを使います");
            ConsoleWriter.WriteLine ("-[p]roj                       : 対象のプロジェクト名を指定します 名前の先頭一部でも有効です");
            ConsoleWriter.WriteLine ("-[l]ine                       : 行を指定します");
            ConsoleWriter.WriteLine ("-[c]olumn                     : 列を指定します");
            ConsoleWriter.WriteLine ("-findwhat[fw]                 : 検索文字列の指定");
            ConsoleWriter.WriteLine ("-findtarget[ft]               : 検索対象設定");
            ConsoleWriter.WriteLine ("                              : [project カレントプロジェクト(default)]");
            ConsoleWriter.WriteLine ("                              : [solution ソリューション]");
            ConsoleWriter.WriteLine ("-enablefindmatchcase[efm]     : 大文字小文字判定有り無し");
            ConsoleWriter.WriteLine ("                              : [default 判定なし]");
            ConsoleWriter.WriteLine ("-debugwrite[dw]               : 詳細情報も出力(debug用)");
        }
    }
}
