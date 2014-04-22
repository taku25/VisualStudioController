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
        
        public class ProjectBuildInfo{
            public ProjectBuildInfo(System.String projectName, bool isBuild)
            {
                projectName_ = projectName;
                isBuild_ = isBuild;
            }

            private System.String projectName_;
            private bool isBuild_;
            public System.String ProjectName
            {
                get { return projectName_; }
            }

            public bool IsBuild
            {
                get { return isBuild_; }
            }
        }

        
#region めんば
        static private System.String VisualStudioProcessName = @"!VisualStudio.DTE";
        DTE targetDTE_ = null;
        EnvDTE80.DTE2 targetDTE2_ = null; //2005以上
        EnvDTE.Project targetProject_ = null;
        System.Collections.Generic.List<ProjectBuildInfo> projectBuildInfoList_ = new System.Collections.Generic.List<ProjectBuildInfo>();
#endregion

        public bool Initialize(System.String targetName, System.String projectName = null)
        {
            System.Collections.Generic.List<DTE> dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            

            if(System.String.IsNullOrEmpty(System.IO.Path.GetExtension(targetName)) == true){
                targetDTE_ = GetDTEFromSolutionName(targetName, dtelist);
                if(targetDTE_ == null){
                    targetDTE_ = GetDTEFromProjectName(targetName, dtelist);
                }
            }else{
                targetDTE_ = GetDTEFromItemFileFullPathName(targetName, dtelist);
            }

            if(targetDTE_ == null){
                ConsoleWriter.WriteDebugLine(@"VisualStudioのプロセスが見つかりません");
            }
            //error一覧とかがほしいので一応これももっておく
            targetDTE2_ = targetDTE_ as EnvDTE80.DTE2;


            if(System.String.IsNullOrEmpty(projectName) == false){
                targetProject_ = GetProjectFromName(projectName);
                projectBuildInfoList_ = CreateProjectBuildInfoList();
                ConsoleWriter.WriteDebugLine(@"対象のプロジェクトが見つかりません");
            }


            return targetDTE_ != null;
        }

        //この引数めっちゃきもちわるいなぁ〜 さてどうすっかなぁ〜..
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
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindSymbolResult) == true){
                WriteFindSymbolResultWindowText();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetAllFiles) == true){
                WriteAllFiles();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.AddBreakPoint) == true){
                AddBreakPoint(argsConvert.FileFullPath, argsConvert.Line, argsConvert.Column);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetErrorList) == true){
                WriteErrorWindowText();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.BuildProject) == true){
                BuildProject(false);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.ReBuildProject) == true){
                BuildProject(true);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.CleanProject) == true){
                CleanProject();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.OpenFile) == true){
                OpenFile(argsConvert.FileFullPath);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.CompileFile) == true){
                CompileFile(argsConvert.FileFullPath, argsConvert.IsWait);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.CancelBuild) == true){
                CancelBuild();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetBuildConfig) == true){
                WriteBuildConfig();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.StopDebugRun) == true){
                StopDebugRun(argsConvert.IsWait);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.CloseSolution) == true){
                CloseSolution();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.Find) == true){
                Find(argsConvert.FindWhat, argsConvert.FindTarget, argsConvert.FindMatchCase, argsConvert.IsWait);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.AddFile) == true){
                AddFile(argsConvert.FileFullPath);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetProjectName) == true){
                WriteProjectName(argsConvert.FileFullPath);
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetCurrnetProjectName) == true){
                WriteCurrentProjectName();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetStartUpProjectName) == true){
                WriteStartUpProjectName();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetSolutionName) == true){
                WriteSolutionName();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetSolutionFileName) == true){
                WriteSolutionFileName();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetSolutionFullPath) == true){
                WriteSolutionFullPath();
            }else if (argsConvert.Commnad(ArgsConvert.CommandType.GetBuildStatus) == true){
                WriteBuildStatus();
            }
        }


        private bool _getDTEFromItemFileFullPathName (String name, ProjectItem item)
        {
            if(item.Kind == Constants.vsProjectItemKindPhysicalFile){
                if(item.FileNames[0].ToLower() == name){
                    return true;
                }
            }else{
                foreach (ProjectItem childItem in item.ProjectItems){
                    if(_getDTEFromItemFileFullPathName(name, childItem) == true){
                        return true;
                    }
                }
            }
            return false;
        }

        public EnvDTE.Project GetProjectFromItemFullPathName(System.String itemFullPathName, EnvDTE.DTE dte)
        { 
            for(int i = 0; i < dte.Solution.Projects.Count; i++){
                Project project = dte.Solution.Projects.Item(i + 1);
                foreach (ProjectItem item in project.ProjectItems){
                    if(_getDTEFromItemFileFullPathName (itemFullPathName, item) == true){
                        return project;
                    }
                }
            }
            return null;
        }

        private DTE GetDTEFromItemFileFullPathName(String name, System.Collections.Generic.List<DTE> dtelist = null)
        {
            name = name.ToLower();
            if(dtelist == null){
                dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            }
            //不正なDTEが取れることがあるのでとりあえずの対処
            foreach(DTE dte in dtelist){
                try{
                    if(GetProjectFromItemFullPathName(name, dte) != null){
                        return dte;
                    }
                    //solutionの名前でもチェックをする
                    if(dte.Solution.FullName.ToLower () == name.ToLower()){
                        return dte;
                    }
                }catch{
                }
            }
            return null;
        }

        private DTE GetDTEFromProjectName(String name, System.Collections.Generic.List<DTE> dtelist = null)
        {
            name = name.ToLower();
            if(dtelist == null){
                dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            }
            
            name = name.ToLower();
            DTE tempDte = null;
            try{
            //最初に完全一致かどうか
            foreach(DTE dte in dtelist){
                System.String solutionFullName = System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower();
                //foreachがうまくいかないので
                for(int i = 0; i < dte.Solution.Projects.Count; i++){
                    Project project = dte.Solution.Projects.Item(i + 1);
                    if(project.Name.ToLower() == name){
                        tempDte = dte;
                        break;
                    }
                }
            }

            if(tempDte  == null){
                //先頭から一部分のひかくでもokにする
                foreach(DTE dte in dtelist){
                    for(int i = 0; i < dte.Solution.Projects.Count; i++){
                        Project project = dte.Solution.Projects.Item(i + 1);
                        if(project.Name.ToLower().IndexOf(name) == 0){
                            tempDte = dte;
                            break;
                        }
                    }
                }
            }
            }catch{
            }

            return tempDte;
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
            DTE tempDte = null;
            try{
            //最初に完全一致かどうか
            foreach(DTE dte in dtelist){
                System.String solutionFullName = System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower();
                if(solutionFullName == name){
                    tempDte = dte;
                    break;
                }
            }

            if(tempDte  == null){
                //先頭から一部分のひかくでもokにする
                foreach(DTE dte in dtelist){
                    System.String solutionFullName = System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower();
                    if(solutionFullName.IndexOf (name) == 0){
                        tempDte = dte;
                        break;
                    }
                }
            }
            }catch{
            }
            return tempDte;
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

        public EnvDTE.Project GetProjectFromName(System.String projectName)
        {
            EnvDTE.Project tempProject = null;
            projectName = projectName.ToLower();
            foreach (Project project in targetDTE_.Solution.Projects){
                if(projectName == project.Name.ToLower()){
                    tempProject = project;
                    break;
                }
            }

            if(tempProject == null){
                foreach (Project project in targetDTE_.Solution.Projects){
                    if(project.Name.ToLower().IndexOf(projectName) == 0){
                        tempProject = project;
                        break;
                    }
                }
            }
            return tempProject;
        }



        //
        private System.Collections.Generic.List<ProjectBuildInfo> CreateProjectBuildInfoList()
        {
            System.Collections.Generic.List<ProjectBuildInfo> projectBuildInfoList = new System.Collections.Generic.List<ProjectBuildInfo>();
            foreach (EnvDTE.SolutionContext context in targetDTE_.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts){
                projectBuildInfoList.Add(new ProjectBuildInfo(context.ProjectName, context.ShouldBuild));
            }
            return projectBuildInfoList;
        }

        private void RestoreProjectBuildInfo(System.Collections.Generic.List<ProjectBuildInfo> projectBuildInfoList)
        {
            foreach (ProjectBuildInfo projectInfo in projectBuildInfoList){
                foreach (EnvDTE.SolutionContext context in targetDTE_.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts){
                    if(projectInfo.ProjectName == context.ProjectName){
                        context.ShouldBuild = projectInfo.IsBuild;
                        break;
                    }
                }
            }
        }

        public void BuildProject(bool rebuild)
        {
            foreach (EnvDTE.SolutionContext context in targetDTE_.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts){
                if (context.ProjectName == targetProject_.UniqueName){
                    context.ShouldBuild = true;
                }else{
                    context.ShouldBuild = false;
                }
            }
   
            if(rebuild == true){
                CleanProject();
            }
            //プロジェクトのbuildはbuildinfoを戻さないといけないので かならずbuild待ちを行うなにかいい方法がないものか...
            targetDTE_.Solution.SolutionBuild.Build(true);

            RestoreProjectBuildInfo(projectBuildInfoList_);    
        }

        public void CleanProject()
        {
            foreach (EnvDTE.SolutionContext context in targetDTE_.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts){
                if (context.ProjectName == targetProject_.UniqueName){
                    context.ShouldBuild = true;
                }else{
                    context.ShouldBuild = false;
                }
            }

            targetDTE_.Solution.SolutionBuild.Clean(true);

            RestoreProjectBuildInfo(projectBuildInfoList_);
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

        public void OpenFile(System.String fileFullPath)
        {
            targetDTE_.ItemOperations.OpenFile(fileFullPath);
        }

        public void CompileFile(System.String fileFullPath, bool wait)
        {
            OpenFile(fileFullPath);
            //あまりコマンド使いたくないのだけどわからないのであきらめ
            targetDTE_.ExecuteCommand("Build.Compile");
            

            while (wait){
                try{
                    if(targetDTE_.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateDone){
                        System.Threading.Thread.Sleep(100);
                    }else{
                        break;
                    }
                }catch{
                }
            }
        }

        public void CancelBuild()
        {
            while(true){
                try{
                    if(targetDTE_.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateDone){
                        targetDTE_.ExecuteCommand("Build.Cancel");
                    }else{
                        break;
                    }
                }catch{
                }
            }
        }

        private void _writeAllFileName (ProjectItem item)
        {
            if ((item.Kind == Constants.vsProjectItemKindPhysicalFile)){
                ConsoleWriter.WriteLine(item.FileNames[0]);
            }else{
                foreach (ProjectItem childItem in item.ProjectItems){
                    _writeAllFileName(childItem);
                }
            }
        }

        public void WriteAllFiles()
        {
            foreach (Project project in targetDTE_.Solution.Projects){
                foreach (ProjectItem item in project.ProjectItems){
                    _writeAllFileName(item);
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
            
            //クリップボードをいじるので退避
            System.Windows.Forms.IDataObject clipboarData = System.Windows.Forms.Clipboard.GetDataObject();

            try{
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[0-9]+");
                int resultCount = Convert.ToInt32(regex.Match(window.Caption).Value);
                for (int i = 0; i < resultCount; i++){
                    if(i != 0){
                        window.Activate();
                        targetDTE_.Application.ExecuteCommand("Edit.GoToNextLocation", String.Empty);
                    }
                    window.Activate();
                    targetDTE_.Application.ExecuteCommand("Edit.Copy");

                    ConsoleWriter.WriteLine(System.Windows.Forms.Clipboard.GetText());
                }
            }catch{
         
            }

            try{
                System.Windows.Forms.Clipboard.SetDataObject(clipboarData);
            }catch{
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

                System.String lineAndCol = "(" + item.Line.ToString() + ")";

                ConsoleWriter.WriteLine(item.FileName + " " + lineAndCol + ": "  + item.Description);
            }
        }


        public void WriteBuildConfig ()
        {
            
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }

            if(targetDTE2_.ToolWindows.ErrorList == null){
                ConsoleWriter.WriteDebugLine("ErrorListがみつかりません");
                return;
            }

            EnvDTE80.SolutionConfiguration2 config = targetDTE2_.Solution.SolutionBuild.ActiveConfiguration as EnvDTE80.SolutionConfiguration2;
            ConsoleWriter.WriteLine(config.Name + "/" + config.PlatformName);
        }

        public void StopDebugRun(bool wait)
        {
            while(true){
                try{
                    if(targetDTE_.Debugger.CurrentMode != dbgDebugMode.dbgDesignMode){
                        targetDTE_.Debugger.Stop(wait);
                    }else{
                        break;
                    }
                }catch{
                }
            }
        }

        public void CloseSolution()
        {
            targetDTE_.Solution.Close();
        }

        public void Find(System.String findWhat, ArgsConvert.FindTargetType findTarget, bool findMatchCase, bool wait)
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }


            targetDTE2_.Find.Action = vsFindAction.vsFindActionFindAll;
            targetDTE2_.Find.FindWhat = findWhat;
            targetDTE2_.Find.MatchCase = findMatchCase;
            targetDTE2_.Find.Target = (findTarget == ArgsConvert.FindTargetType.Project) ? vsFindTarget.vsFindTargetCurrentProject : vsFindTarget.vsFindTargetSolution;

            if(wait == true){
                while((targetDTE2_.Find.Execute() == vsFindResult.vsFindResultPending)){
                }
            }else{
                targetDTE2_.Find.Execute();
            }
            /*
            //↓C#も左から条件判定だっけ？ あやしいのでやめる
            while((targetDTE2_.Find.Execute() != vsFindResult.vsFindResultPending) && wait == true){
            }
            */
        }

        void AddFile(System.String addFileName)
        {
            if(targetProject_ == null){
                ConsoleWriter.WriteDebugLine("対象のプロジェクトが見つかりませんでした");
                return;
            }

            EnvDTE.ProjectItem projectItem = targetProject_.ProjectItems.AddFromFile(addFileName);
            if(projectItem == null){
                ConsoleWriter.WriteDebugLine("追加失敗");
                return;
            }
            this.OpenFile(projectItem.FileNames[0]);
        }
 
        void WriteProjectName(System.String fileFullPathName)
        {
            EnvDTE.Project project = GetProjectFromItemFullPathName(fileFullPathName, targetDTE_);
            if(project == null){
                ConsoleWriter.WriteDebugLine("対象のプロジェクトが見つかりませんでした");
                return;
            }
            ConsoleWriter.WriteLine(project.Name);
        }
        
        void WriteCurrentProjectName()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }
            foreach(EnvDTE.Project project in targetDTE2_.ActiveSolutionProjects){
                ConsoleWriter.WriteLine(project.Name);
                break;
            }
        }

        void WriteStartUpProjectName()
        {
             
            System.Array stringArray = targetDTE_.Solution.SolutionBuild.StartupProjects as System.Array;
            ConsoleWriter.WriteLine(System.IO.Path.GetFileNameWithoutExtension(stringArray.GetValue(0).ToString()));    
        }

        void WriteSolutionName()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetFileNameWithoutExtension(targetDTE_.Solution.FullName));
        }

        void WriteSolutionFileName()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetFileName(targetDTE_.Solution.FullName));
        }
        
        void WriteSolutionFullPath()
        {
            ConsoleWriter.WriteLine(targetDTE_.Solution.FullName);
        }

        void WriteBuildStatus()
        {
            if(targetDTE_.Solution.SolutionBuild.BuildState == vsBuildState.vsBuildStateDone){
                ConsoleWriter.WriteLine("Done");
            }else if(targetDTE_.Solution.SolutionBuild.BuildState == vsBuildState.vsBuildStateInProgress){
                ConsoleWriter.WriteLine("InProgress");
            }else if (targetDTE_.Solution.SolutionBuild.BuildState == vsBuildState.vsBuildStateNotStarted){
                ConsoleWriter.WriteLine("NotStarted");
            }else{
                ConsoleWriter.WriteLine("unknown");
            }
        }


    }
}

