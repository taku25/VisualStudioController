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
            commandArray_[(int)ArgsConvert.CommandType.Build]   = BuildSolution;
            commandArray_[(int)ArgsConvert.CommandType.ReBuild] = ReBuildSolution;
            commandArray_[(int)ArgsConvert.CommandType.Clean] = CleanSolution;
            commandArray_[(int)ArgsConvert.CommandType.Run] = RunSolution;
            commandArray_[(int)ArgsConvert.CommandType.DebugRun] = DebugRunSolution;
            commandArray_[(int)ArgsConvert.CommandType.GetFile] = WriteCurrentEditFileName;
            commandArray_[(int)ArgsConvert.CommandType.GetOutput] = WriteOutputWindowText;
            commandArray_[(int)ArgsConvert.CommandType.GetFindResult1] = WriteFindResultWindowText1;
            commandArray_[(int)ArgsConvert.CommandType.GetFindResult2] = WriteFindResultWindowText2;
            commandArray_[(int)ArgsConvert.CommandType.GetFindSymbolResult] = WriteFindSymbolResultWindowText;
            commandArray_[(int)ArgsConvert.CommandType.GetAllFiles] = WriteAllFiles;
            commandArray_[(int)ArgsConvert.CommandType.AddBreakPoint] = AddBreakPoint;
            commandArray_[(int)ArgsConvert.CommandType.GetErrorList] = WriteErrorWindowText;
            commandArray_[(int)ArgsConvert.CommandType.BuildProject] = BuildProject;
            commandArray_[(int)ArgsConvert.CommandType.ReBuildProject] = ReBuildProject;
            commandArray_[(int)ArgsConvert.CommandType.CleanProject] = CleanProject;
            commandArray_[(int)ArgsConvert.CommandType.OpenFile] = OpenFile;
            commandArray_[(int)ArgsConvert.CommandType.CompileFile] = CompileFile;
            commandArray_[(int)ArgsConvert.CommandType.CancelBuild] = CancelBuild;
            commandArray_[(int)ArgsConvert.CommandType.StopDebugRun] = StopDebugRun;
            commandArray_[(int)ArgsConvert.CommandType.CloseSolution] = CloseSolution;
            commandArray_[(int)ArgsConvert.CommandType.Find] = Find;
            commandArray_[(int)ArgsConvert.CommandType.AddFile] = AddFile;

            //プロジェクト
            commandArray_[(int)ArgsConvert.CommandType.GetProjectName] = WriteProjectName;
            commandArray_[(int)ArgsConvert.CommandType.GetCurrnetProjectName] = WriteCurrentProjectName;
            commandArray_[(int)ArgsConvert.CommandType.GetStartUpProjectName] = WriteStartUpProjectName;
            commandArray_[(int)ArgsConvert.CommandType.SetStartUpProject] = SetStartUpProject;
            commandArray_[(int)ArgsConvert.CommandType.GetProjectList] = WriteProjectList;

            //ソリューション
            commandArray_[(int)ArgsConvert.CommandType.GetSolutionName] = WriteSolutionName;
            commandArray_[(int)ArgsConvert.CommandType.GetSolutionFileName] = WriteSolutionFileName;
            commandArray_[(int)ArgsConvert.CommandType.GetSolutionFullPath] = WriteSolutionFullPath;


            //ビルドコンフィグ回り
            commandArray_[(int)ArgsConvert.CommandType.GetCurrentBuildConfig] = WriteCurrentBuildConfig;
            commandArray_[(int)ArgsConvert.CommandType.SetCurrentBuildConfig] = SetCurrnetBuildConfig;
            commandArray_[(int)ArgsConvert.CommandType.GetBuildConfigList] = WriteBuildConfigList;
            commandArray_[(int)ArgsConvert.CommandType.GetPlatformList] = WritePlatformList;
            
            commandArray_[(int)ArgsConvert.CommandType.GetSolutionDirectory] = WriteSolutionDirectory;
            
            
            commandArray_[(int)ArgsConvert.CommandType.GetBuildStatus] = WriteBuildStatus;
            
                
            commandArray_[(int)ArgsConvert.CommandType.UnKnown] = UnknownAction;
            
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
        private DTE targetDTE_ = null;
        private EnvDTE80.DTE2 targetDTE2_ = null; //2005以上
        private EnvDTE.Project targetProject_ = null;
        
        private System.Collections.Generic.List<ProjectBuildInfo> projectBuildInfoList_ = new System.Collections.Generic.List<ProjectBuildInfo>();

        private bool isWait_ = false;
        private System.String fileFullPath_ = "";
        private int line_ = 1;
        private int column_ = 1;
        private ArgsConvert.FindTargetType findTarget_;
        private System.String findWhat_ = "";
        private bool findMatchCase_ = false;
        private Action[] commandArray_ = new Action[Enum.GetValues(typeof(ArgsConvert.CommandType)).Length];
        private Action currentCommand_ = null;

        private System.String platformName_ = "";
        private System.String buildConfigName_ = "";
        private EnvDTE80.SolutionConfiguration2 currentBuildConfiguration_ = null;
        private System.Collections.Generic.List<EnvDTE80.SolutionConfiguration2> buildConfigurationList_ = new System.Collections.Generic.List<EnvDTE80.SolutionConfiguration2>();
        private EnvDTE.vsFindResultsLocation findResultsLocation_ = vsFindResultsLocation.vsFindResults1;


#endregion

        #region あくせっさ
        public bool IsWait
        {
            get { return isWait_; }
            set { isWait_ = value; }
        }

        public System.String FileFullPath
        {
            get { return fileFullPath_; }
            set { fileFullPath_  = value; }
        }

        public int Line
        {
            get { return line_; }
            set { line_ = value; }
        }
        
        public int Column
        {
            get { return column_; }
            set { column_ = value; }
        }

        public ArgsConvert.FindTargetType FindTarget
        {
            get { return findTarget_; }
            set { findTarget_ = value; }
        }
        public System.String FindWhat
        {
            get { return findWhat_; }
            set { findWhat_  = value; }
        }
     
        public bool FindMatchCase
        {
            get { return findMatchCase_; }
            set { findMatchCase_  = value; }
        }

        public System.String PlatformName
        {
            get { return platformName_; }
            set { platformName_ = value; }
        }

        public System.String BuildConfigName
        {
            get { return buildConfigName_; }
            set { buildConfigName_ = value; }
        }

        public EnvDTE.vsFindResultsLocation  FindResultsLocation 
        {
            get { return  findResultsLocation_; }
            set { findResultsLocation_ = value; }
        }


        #endregion

        public bool Initialize(ArgsConvert argsConvert)
        {
            System.Collections.Generic.List<DTE> dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);


            targetDTE_ = GetDTEFromItemFileFullPathName(argsConvert.TargetName, dtelist);
            if (targetDTE_ == null){
                targetDTE_ = GetDTEFromSolutionName(argsConvert.TargetName, dtelist);
                if(targetDTE_ == null){
                    targetDTE_ = GetDTEFromProjectName(argsConvert.TargetName, dtelist);
                }
            }

            if(targetDTE_ == null){
                ConsoleWriter.WriteDebugLine(@"VisualStudioのプロセスが見つかりません");
                return false;
            }

            //error一覧とかがほしいので一応これももっておく
            try{
                targetDTE2_ = targetDTE_ as EnvDTE80.DTE2;
            }catch{
            }

            if(System.String.IsNullOrEmpty(argsConvert.TargetProjectName) == false){
                targetProject_ = GetProjectFromName(argsConvert.TargetProjectName);
                ConsoleWriter.WriteDebugLine(@"対象のプロジェクトが見つかりません");
            }else{
                targetProject_ = GetProjectFromItemFullPathName(argsConvert.TargetName, targetDTE_);
                if(targetProject_ == null){
                    //startupプロジェクトにする
                    targetProject_ = GetProjectFromName(GetStartUpProjectName());
                }
            }
            
            if( targetProject_ != null){
                projectBuildInfoList_ = CreateProjectBuildInfoList();
            }


            IsWait = argsConvert.IsWait;
            FileFullPath = argsConvert.FileFullPath;
            Line = argsConvert.Line;
            Column = argsConvert.Column;
            FindTarget = argsConvert.FindTarget;
            FindWhat = argsConvert.FindWhat;
            FindMatchCase = argsConvert.FindMatchCase;

            PlatformName = argsConvert.PlatformName;
            BuildConfigName = argsConvert.BuildConfigName;

            if(argsConvert.FindResultLocations == ArgsConvert.FindResultLocation.one){
                FindResultsLocation = vsFindResultsLocation.vsFindResults1;
            }else{
                FindResultsLocation = vsFindResultsLocation.vsFindResults2;
            }

            currentBuildConfiguration_ = targetDTE_.Solution.SolutionBuild.ActiveConfiguration as EnvDTE80.SolutionConfiguration2;
            foreach(EnvDTE80.SolutionConfiguration2 config in targetDTE_.Solution.SolutionBuild.SolutionConfigurations){                
                buildConfigurationList_.Add(config);
            }

            //実行するコマンド
            currentCommand_ = commandArray_[(int)argsConvert.GetRunCommandType()];


            return true;
        }

  
        public void Run ()
        {
            currentCommand_();
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

                        if(System.IO.Path.GetFileNameWithoutExtension (project.FullName).ToLower() == name){
                            tempDte = dte;
                        }else if(System.IO.Path.GetFileName (project.FullName).ToLower() == name){
                            tempDte = dte;
                        }else if(project.FullName.ToLower() == name){
                            tempDte = dte;
                        }else if (System.IO.Path.GetFileName (project.FullName).ToLower().IndexOf(name) == 0){
                            tempDte = dte;
                        }
                        if(tempDte != null){
                            break;
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
                foreach(DTE dte in dtelist){
                    //
                    if(System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower() == name){
                        tempDte = dte;
                    }else if(System.IO.Path.GetFileName (dte.Solution.FullName).ToLower() == name){
                        tempDte = dte;
                    }else if(dte.Solution.FullName.ToLower() == name){
                        tempDte = dte;
                    }else if (System.IO.Path.GetFileName (dte.Solution.FullName).ToLower().IndexOf(name) == 0){
                        tempDte = dte;
                    }
                    if(tempDte != null){
                        break;
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
                    }else{
                        try{
                            Marshal.ThrowExceptionForHR(rot.GetObject(moniker[0], out runningObject));
                            if(runningObject is DTE){
                                list.Add((DTE)runningObject);
                            }
                        }catch{
                        }
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

        public void BuildSolution()
        {
            BuildSolution(false);
        }

        public void ReBuildSolution()
        {
            BuildSolution(true);
        }

        public void BuildSolution(bool rebuild)
        {
            if(rebuild == true){
                CleanSolution();
            }
            targetDTE_.Solution.SolutionBuild.Build(IsWait);
            
        }

        public void CleanSolution()
        {
            targetDTE_.Solution.SolutionBuild.Clean(true);
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

        private void SetBuildMarkProjectBuildInfo(System.String markProjectName, System.Collections.Generic.List<ProjectBuildInfo> projectBuildInfoList)
        {
            foreach (EnvDTE.SolutionContext context in targetDTE_.Solution.SolutionBuild.ActiveConfiguration.SolutionContexts){
                if (context.ProjectName == markProjectName){
                    context.ShouldBuild = true;
                }else{
                    context.ShouldBuild = false;
                }
            }
       
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


        public void BuildProject()
        {
            BuildProject(false);
        }

        public void ReBuildProject()
        {
            BuildProject(true);
        }

        public void BuildProject(bool rebuild)
        {
            if(rebuild == true){
                CleanProject(true);
            }
            
            targetDTE_.Solution.SolutionBuild.BuildProject(targetDTE_.Solution.SolutionBuild.ActiveConfiguration.Name, targetProject_.UniqueName, IsWait);

        }

        public void CleanProject()
        {
            CleanProject(true);
        }

        public void CleanProject(bool changeProjectMark)
        {
            if(changeProjectMark == true){
                SetBuildMarkProjectBuildInfo(targetProject_.UniqueName, projectBuildInfoList_);
            }

            targetDTE_.Solution.SolutionBuild.Clean(true);

            if(changeProjectMark == true){
                RestoreProjectBuildInfo(projectBuildInfoList_);
            }
        }


        public void WriteCurrentEditFileName()
        {
            EnvDTE.Document document = targetDTE_.ActiveDocument;
            ConsoleWriter.WriteLine(document.FullName);
        }


        public void AddBreakPoint()
        {
            targetDTE_.Solution.DTE.Debugger.Breakpoints.Add("", FileFullPath, Line, Column);
        }

        public void OpenFile()
        {
            targetDTE_.ItemOperations.OpenFile(FileFullPath);
        }


        public void CompileFile()
        {
            targetDTE_.ItemOperations.OpenFile(FileFullPath);
            //あまりコマンド使いたくないのだけどわからないのであきらめ
            targetDTE_.ExecuteCommand("Build.Compile");
            

            while (IsWait){
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
            //build中でないのであればむし
            if(targetDTE_.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateInProgress){
                return;
            }
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
        
        public void WriteFindResultWindowText1 ()
        {
            WriteFindResultWindowText(0);
        }

        public void WriteFindResultWindowText2 ()
        {
            WriteFindResultWindowText(1);
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

                System.String lineAndCol = "(" + item.Line.ToString() + "," + item.Column.ToString() + ")";

                ConsoleWriter.WriteLine(item.FileName + " " + lineAndCol + ": "  + item.Description);
            }
        }

        public void SetCurrnetBuildConfig ()
        {
            EnvDTE80.SolutionConfiguration2 config = GetSolutinConfiguration ();
            if(config == null){
                return;
            }

            config.Activate();           
        }

        public EnvDTE80.SolutionConfiguration2 GetSolutinConfiguration()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return null;
            }

            System.String buildConfigName = this.BuildConfigName;
            if(System.String.IsNullOrEmpty(buildConfigName) == true){
                buildConfigName = this.currentBuildConfiguration_.Name;
            }

            System.String platoformName = this.PlatformName;
            if(System.String.IsNullOrEmpty(platoformName) == true){
                platoformName = this.currentBuildConfiguration_.PlatformName;
            }


            foreach(EnvDTE80.SolutionConfiguration2 config in buildConfigurationList_){
                if(config.Name == buildConfigName && config.PlatformName == platoformName){
                    return config;
                }
            }

            return null;
        }

        public void WriteCurrentBuildConfig ()
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

        public void WriteBuildConfigList ()
        {
            
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }

            if(targetDTE2_.ToolWindows.ErrorList == null){
                ConsoleWriter.WriteDebugLine("ErrorListがみつかりません");
                return;
            }
        
            System.Collections.Generic.List<System.String> configmNameList = new System.Collections.Generic.List<string> ();

            foreach(EnvDTE80.SolutionConfiguration2 config in targetDTE2_.Solution.SolutionBuild.SolutionConfigurations){
                if (config != null){
                    if (configmNameList.Contains(config.Name) == false){
                        configmNameList.Add(config.Name);
                    }
                }
            }   

            foreach(System.String config in configmNameList){
                ConsoleWriter.WriteLine(config);
            }
        }

        public void WritePlatformList ()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }

            if(targetDTE2_.ToolWindows.ErrorList == null){
                ConsoleWriter.WriteDebugLine("ErrorListがみつかりません");
                return;
            }

        
            System.Collections.Generic.List<System.String> platformNameList = new System.Collections.Generic.List<string> ();

            foreach(EnvDTE80.SolutionConfiguration2 config in targetDTE2_.Solution.SolutionBuild.SolutionConfigurations){
                if (config != null){
                    if (platformNameList.Contains(config.PlatformName) == false){
                        platformNameList.Add(config.PlatformName);
                    }
                }
            }   

            foreach(System.String platform in platformNameList){
                ConsoleWriter.WriteLine(platform);
            }
        }


        public void StopDebugRun()
        {
            while(true){
                try{
                    if(targetDTE_.Debugger.CurrentMode != dbgDebugMode.dbgDesignMode){
                        targetDTE_.Debugger.Stop(IsWait);
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

        public void Find()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("visual studio2005以上でないと使用できません");
                return;
            }


            targetDTE2_.Find.Action = vsFindAction.vsFindActionFindAll;
            targetDTE2_.Find.FindWhat = FindWhat;
            targetDTE2_.Find.MatchCase = FindMatchCase;
            targetDTE2_.Find.Target = (FindTarget == ArgsConvert.FindTargetType.Project) ? vsFindTarget.vsFindTargetCurrentProject : vsFindTarget.vsFindTargetSolution;
            targetDTE2_.Find.ResultsLocation = FindResultsLocation;

            if(IsWait){
                while((targetDTE2_.Find.Execute() == vsFindResult.vsFindResultPending)){
                    System.Threading.Thread.Sleep(100);
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

        void AddFile()
        {
            if(targetProject_ == null){
                ConsoleWriter.WriteDebugLine("対象のプロジェクトが見つかりませんでした");
                return;
            }

            EnvDTE.ProjectItem projectItem = targetProject_.ProjectItems.AddFromFile(FileFullPath);
            if(projectItem == null){
                ConsoleWriter.WriteDebugLine("追加失敗");
                return;
            }

            targetDTE_.ItemOperations.OpenFile(FileFullPath);
        }
 
        void WriteProjectName()
        {
            ConsoleWriter.WriteLine(targetProject_.Name);
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

        System.String GetStartUpProjectName ()
        {    
            System.Array stringArray = targetDTE_.Solution.SolutionBuild.StartupProjects as System.Array;
            return System.IO.Path.GetFileNameWithoutExtension(stringArray.GetValue(0).ToString());    
        }

        void SetStartUpProject()
        {
            if(targetProject_ == null){
                return;
            }
            targetDTE_.Solution.SolutionBuild.StartupProjects = targetProject_.FullName;

        }

        void WriteProjectList()
        {
            try{
                for(int i = 0; i < targetDTE_.Solution.Projects.Count; i++){
                    Project project = targetDTE_.Solution.Projects.Item(i + 1);
                    ConsoleWriter.WriteLine(project.Name); 
                }
            }catch{

            }
        }

        void WriteStartUpProjectName()
        {
            ConsoleWriter.WriteLine(GetStartUpProjectName());    
        }

        void WriteSolutionDirectory ()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetDirectoryName(targetDTE_.Solution.FullName));
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

        void UnknownAction()
        {
            ConsoleWriter.WriteLine("UnknownAction");
        }

    }
}

