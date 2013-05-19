using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;

namespace VisualStudioController {

    public class VSController : System.IDisposable{

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        public VSController ()
        {
            targetDTE_ = null;
        }
        
        public void Dispose ()
        {
        }
        
        
#region めんば
        static private System.String VisualStudioProcessName = @"!VisualStudio.DTE";
//        private System.String targetVersoinErrorListGUID_;
//        private const System.String errorListWindowGUID = "{D78612C7-9962-4B83-95D9-268046DAD23A}";
        DTE targetDTE_;
        EnvDTE80.DTE2 targetDTE2_; //2005以上

#endregion

        public bool Initialize(System.String targetName)
        {
            System.Collections.Generic.List<DTE> dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            
            targetDTE_ = GetDTEFromSolutionName(targetName, dtelist);
            if(targetDTE_ == null && System.String.IsNullOrEmpty(targetName) == false){
                targetDTE_ = GetDTEFromEditFileName(targetName, dtelist);
            }
            

            if(targetDTE_ == null){
                ConsoleWriter.WriteDebugLine(@"VisualStudioのプロセスが見つかりません");
            }
            //error一覧とかがほしいので一応これももっておく
            targetDTE2_ = targetDTE_ as EnvDTE80.DTE2;
            return targetDTE_ != null;
        }

        //この引数めっちゃきもちわるいなぁ～ さてどうすっかなぁ～..
        public void Run (ArgsConvert argsConvert)
        {
            if(argsConvert.Commnad(ArgsConvert.CommandType.Build) == true){
                BuildSolution(false, argsConvert.IsWait);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.ReBuild) == true){
                BuildSolution(true, argsConvert.IsWait);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.Clean) == true){
                CleanSolution(argsConvert.IsWait);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.Run) == true){
                RunSolution();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.DebugRun) == true){
                DebugRunSolution();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFile) == true){
                WriteCurrentEditFileName();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetOutput) == true){
                WriteOutputWindowText();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindResult1) == true){
                WriteFindResultWindowText(0);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindResult2) == true){
                WriteFindResultWindowText(1);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindSimbolResult) == true){
                WriteFindSymbolResultWindowText();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetAllFile) == true){
                GetAllFile();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.AddBreakPoint) == true){
                AddBreakPoint(argsConvert.FileFullPath, argsConvert.Line, argsConvert.Column);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetErrorList) == true){
                WriteErrorWindowText();
            }
        }


        private DTE GetDTEFromEditFileName(String name, System.Collections.Generic.List<DTE> dtelist = null)
        {
            name = name.ToLower();
            if(dtelist == null){
                dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            }
            foreach(DTE dte in dtelist){
                foreach (Project project in dte.Solution.Projects){
                    foreach (ProjectItem item in project.ProjectItems){
                        if ((item.Kind == Constants.vsProjectItemKindPhysicalFile)){
                            if(item.FileNames[0].ToLower () == name){
                                return dte;
                            }
                        }
                    }
                }
            }
            return null;
        }

        private DTE GetDTEFromSolutionName(String name, System.Collections.Generic.List<DTE> dtelist = null)
        {
            if(dtelist == null){
                dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            }
            //とりあえずいっこめをかえす
            if(System.String.IsNullOrEmpty(name) == true){
                foreach(DTE dte in dtelist){
                    return dte;
                }
            }
            name = name.ToLower();
            foreach(DTE dte in dtelist){
                //先頭から一部分のひかくでもokにする
                //最初に完全一致かどうか
                System.String solutionFullName = System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower();
                if(solutionFullName == name){
                    return dte;
                }
                //次に先頭から同じかどうか
                if(solutionFullName.IndexOf (name) == 0){
                    return dte;
                }
            }
            return null;
        }

        private System.Collections.Generic.List<DTE> GetDTEFromProcessListFromName(String name)
        {
            System.Collections.Generic.List<DTE> list = new System.Collections.Generic.List<DTE>();

            IRunningObjectTable rot = null;
            IEnumMoniker enumMoniker = null;
            int retVal = GetRunningObjectTable(0, out rot);

            if (retVal == 0)
            {
                rot.EnumRunning(out enumMoniker);

                IntPtr fetched = IntPtr.Zero;
                IMoniker[] moniker = new IMoniker[1];
                while (enumMoniker.Next(1, moniker, fetched) == 0)
                {
                    object runningObject = null;
                    IBindCtx bindCtx = null;
                    CreateBindCtx(0, out bindCtx);
                    string displayName;
                    moniker[0].GetDisplayName(bindCtx, null, out displayName);
                    if(displayName.Contains(name) == true){
                        Marshal.ThrowExceptionForHR(rot.GetObject(moniker[0], out runningObject));
                        list.Add((DTE)runningObject);
                    }

                    if(bindCtx != null){
                        Marshal.ReleaseComObject(bindCtx);
                    }
                }

                if (enumMoniker != null){
                    Marshal.ReleaseComObject(enumMoniker);
                }
            }

            if(rot != null){
                Marshal.ReleaseComObject(rot);
            }

            return list;
        }

        public void BuildSolution(bool rebuild, bool wait)
        {
            if(rebuild == true){
                CleanSolution(true);
            }
            targetDTE_.Solution.SolutionBuild.Build(wait);
        }
        public void CleanSolution(bool wait)
        {
            targetDTE_.Solution.SolutionBuild.Clean(wait);
        }
        
        public void RunSolution()
        {
            targetDTE_.Solution.SolutionBuild.Run();
        }

        public void DebugRunSolution()
        {
            targetDTE_.Solution.SolutionBuild.Debug();
        }

        public System.String WriteCurrentEditFileName()
        {
            EnvDTE.Document document = targetDTE_.ActiveDocument;
            ConsoleWriter.WriteLine(document.FullName);
         
            return document.FullName;
        }


        public void AddBreakPoint(System.String fileFullPath, int line, int column)
        {
            targetDTE_.Solution.DTE.Debugger.Breakpoints.Add("", fileFullPath, line, column);
        }

        public void GetAllFile()
        {
            foreach (Project project in targetDTE_.Solution.Projects){
                foreach (ProjectItem item in project.ProjectItems){
                    if ((item.Kind == Constants.vsProjectItemKindPhysicalFile)){
                        ConsoleWriter.WriteLine(item.FileNames[0]);
                    }
                }
            }
        }


        public void WriteOutputWindowText ()
        {
            OutputWindow outputWindow = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object as OutputWindow;
            EnvDTE.TextDocument textDocument = outputWindow.ActivePane.TextDocument;
            textDocument.Selection.SelectAll();
            System.String outputString = textDocument.Selection.Text;
            ConsoleWriter.WriteLine(outputString);
        }
        
        public void WriteFindResultWindowText (int type)
        {
            Window window = null;
            if(type == 0){
               window = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindResults1);
            }else if(type == 1){
               window = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindResults2);
            }
            if(window == null){
                ConsoleWriter.WriteDebugLine("検索結果Windowがみつかりませんでした");
                return;
            }
            TextSelection textSelection = window.Selection as TextSelection;
            textSelection.SelectAll();
            System.String outputString = textSelection.Text;
            ConsoleWriter.WriteLine(outputString);
        }

        public void WriteFindSymbolResultWindowText ()
        {
            
            Window window = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindSymbolResults);
            
            if(window == null){
                ConsoleWriter.WriteDebugLine("検索結果Windowがみつかりませんでした");
                return;
            }
            

            //とりかたを模索中
            //これにもはいっていなかった...
            //もしかしてerrorとおなじでぜんぜん違う場所にかくのうされているのか？
            foreach(EnvDTE.ContextAttribute contextAttribute in window.ContextAttributes){
                if(contextAttribute.Values is object[]){
                    object[] objectarray = contextAttribute.Values as object[];
                    if ("FindSymbolResultsWindow" == (objectarray[0] as System.String)){
                        foreach(EnvDTE.ContextAttribute childContextAttribute in contextAttribute.Collection){
                        }
                    }
                }
            }
            
            
        }

        public void WriteErrorWindowText ()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }

            if(targetDTE2_.ToolWindows.ErrorList == null){
                ConsoleWriter.WriteDebugLine("ErrorListがみつかりません");
                return;
            }

            EnvDTE80.ErrorItems errorItems =  targetDTE2_.ToolWindows.ErrorList.ErrorItems;
            if(errorItems == null){
                ConsoleWriter.WriteDebugLine("ItemListがみつかりません");
                return;
            }
      
            for(int i = 1; i <= errorItems.Count; i++){
                EnvDTE80.ErrorItem item = errorItems.Item(i);
                ConsoleWriter.WriteLine("message@" + item.Description + " file@" + item.FileName + " line@" + item.Line + " column@" + item.Column.ToString());
            }
        }
    }
}

