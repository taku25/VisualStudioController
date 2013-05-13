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
            GetFile,
            GetOutput,
            GetFindResult1,
            GetFindResult2,
            GetFindSimbolResult,
            Run,
            //Addbreck
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

        public bool Commnad (CommandType commandType)
        {
            return commnadType_[(int)commandType];
        }

        public System.String TargetName { get { return targetName_; } }
        public bool IsWait { get { return isWait_; } }

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
                }else if (args[0] == "Run"){
                    commnadType_[(int)CommandType.ReBuild] = true;
                }else if (args[0] == "clean"){
                    commnadType_[(int)CommandType.ReBuild] = true;
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
                }
            
                for(int i = 1; i < args.Length; i++){
                    if(args[i].ToLower() == "-t"){
                        targetName_ = args[i + 1];
                        i+=1;
                    }else if(args[i].ToLower() == "-w"){
                        isWait_ = true;
                    }else if (args[i].ToLower() == "-d"){
                        ConsoleWriter.DebugEnable = true;
                    }
                }

                foreach(String str in args){
                    ConsoleWriter.WriteDebugLine(str);
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
            ConsoleWriter.WriteLine ("getfile             : 編集中ファイル取得" );
            ConsoleWriter.WriteLine ("getoutput           : 出力Windowの中身を取得");
            ConsoleWriter.WriteLine ("getfindresult1       : 検索結果1を取得");
            ConsoleWriter.WriteLine ("getfindresult2       : 検索結果2を取得");
            //ConsoleWriter.WriteLine ("getfindsimbolresult : シンボルの検索結果を取得");
            //Console.WriteLine ("addbreck                : ブレークポイントの追加");
            ConsoleWriter.WriteLine ("<options>           : ");
            ConsoleWriter.WriteLine ("-t                  : [-t SorceFilePath(fullpath) or -t SolutionName(name)] ターゲットソリューション名 か ターゲットソリューションに含まれているソースファイル名");
            ConsoleWriter.WriteLine ("-w                  : 終わるまで待つ(build and rebuild時に有効)");
            ConsoleWriter.WriteLine ("-d                  : 詳細情報も出力(debug用)");
            //Console.WriteLine ("-enc      : コンソール出力を行うエンコード指定");
        }
    }
}
