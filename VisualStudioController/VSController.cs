using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;

namespace VisualStudioController {

    public partial class VSController : System.IDisposable{

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
            commandArray_[(int)ArgsConvert.CommandType.GetCurrentFileInfo] = WriteCurrentFileInfo;
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
            commandArray_[(int)ArgsConvert.CommandType.FindSymbol] = FindSymbol;
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

            commandArray_[(int)ArgsConvert.CommandType.GoToDefinition] = GoToDefinition;
            commandArray_[(int)ArgsConvert.CommandType.GoToDeclaration] = GoToDeclaration;
            commandArray_[(int)ArgsConvert.CommandType.GetLanguageType] = WriteLanguageType;
            
                
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
        private EnvDTE.ProjectItem targetProjectItem_ = null;
        
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

            if(System.String.IsNullOrEmpty(FileFullPath) == false){
                targetProjectItem_ = GetProjectItemFromItemFullPathName(FileFullPath, targetDTE_);   
            }

           
            //実行するコマンド
            currentCommand_ = commandArray_[(int)argsConvert.GetRunCommandType()];


            return true;
        }

  
        public void Run ()
        {
            currentCommand_();
        }

        private void UnknownAction()
        {
            ConsoleWriter.WriteLine("UnknownAction");
        }

        private ProjectItem _getProjectItemFromItemFileFullPathName (String name, ProjectItem item)
        {
            if(item.Kind == Constants.vsProjectItemKindPhysicalFile){
                if(item.FileNames[0].ToLower() == name.ToLower()){
                    return item;
                }
            }else if(item.ProjectItems != null){
                foreach (ProjectItem childItem in item.ProjectItems){
                    ProjectItem childitem = _getProjectItemFromItemFileFullPathName(name, childItem);
                    if(childitem != null){
                        return childitem;
                    }
                }
            }
            return null;
        }

        private EnvDTE.Project GetProjectFromItemFullPathName(System.String itemFullPathName, EnvDTE.DTE dte)
        { 
            for(int i = 0; i < dte.Solution.Projects.Count; i++){
                Project project = dte.Solution.Projects.Item(i + 1);
                foreach (ProjectItem item in project.ProjectItems){
                    if(_getProjectItemFromItemFileFullPathName (itemFullPathName, item) != null){
                        return project;
                    }
                }
            }
            return null;
        }

        private EnvDTE.ProjectItem GetProjectItemFromItemFullPathName(System.String itemFullPathName, EnvDTE.DTE dte)
        {
            for(int i = 0; i < dte.Solution.Projects.Count; i++){
                Project project = dte.Solution.Projects.Item(i + 1);
                foreach (ProjectItem item in project.ProjectItems){
                    ProjectItem projItem = _getProjectItemFromItemFileFullPathName (itemFullPathName, item);
                    if(projItem != null){
                        return projItem;
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
                            //これで無理やりexpressにも対応
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


        
        private void RunSolution()
        {
            targetDTE_.Solution.SolutionBuild.Run();
        }

        private void DebugRunSolution()
        {
            targetDTE_.Solution.SolutionBuild.Debug();
        }

        private EnvDTE.Project GetProjectFromName(System.String projectName)
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

        private void WriteCurrentFileInfo()
        {
            EnvDTE.Document document = targetDTE_.ActiveDocument;

            int line = 0;
            int column = 0;
            if(document.Selection != null){
                if(document.Selection is TextSelection){
                    TextSelection textSelection = document.Selection as TextSelection;
                    line =  textSelection.ActivePoint.Line;
                    column = textSelection.ActivePoint.DisplayColumn;
                }
            }

            ConsoleWriter.WriteLine("Name=" + document.FullName);
            ConsoleWriter.WriteLine("LanguageType=" + GetLanguageType(targetDTE_.ActiveDocument.ProjectItem));
            ConsoleWriter.WriteLine("Line=" + line.ToString());
            ConsoleWriter.WriteLine("Column=" + column.ToString());

        }


        private void AddBreakPoint()
        {
            targetDTE_.Solution.DTE.Debugger.Breakpoints.Add("", FileFullPath, Line, Column);
        }

        private void OpenFile()
        {
            targetDTE_.ItemOperations.OpenFile(FileFullPath);
        }


        private void _writeAllFileName (ProjectItem item)
        {
            if ((item.Kind == Constants.vsProjectItemKindPhysicalFile)){
                ConsoleWriter.WriteLine(item.FileNames[0]);
            }else if( item.ProjectItems != null){
                foreach (ProjectItem childItem in item.ProjectItems){
                    _writeAllFileName(childItem);
                }
            }
        }

        private void WriteAllFiles()
        {
            foreach (Project project in targetDTE_.Solution.Projects){
                ConsoleWriter.WriteLine(project.FullName);
                foreach (ProjectItem item in project.ProjectItems){
                    _writeAllFileName(item);
                }
            }
        }


        private void WriteOutputWindowText ()
        {
            OutputWindow outputWindow = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object as OutputWindow;
            EnvDTE.TextDocument textDocument = outputWindow.ActivePane.TextDocument;
          
            textDocument.Selection.SelectAll();
            System.String outputString = textDocument.Selection.Text;
            ConsoleWriter.WriteLine(outputString);
        }
        


        

        private void WriteErrorWindowText ()
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

        private void SetCurrnetBuildConfig ()
        {
            EnvDTE80.SolutionConfiguration2 config = GetSolutinConfiguration ();
            if(config == null){
                return;
            }

            config.Activate();           
        }

        private EnvDTE80.SolutionConfiguration2 GetSolutinConfiguration()
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

        private void WriteCurrentBuildConfig ()
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

        private void WriteBuildConfigList ()
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

        private void WritePlatformList ()
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


        private void StopDebugRun()
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

        private void CloseSolution()
        {
            targetDTE_.Solution.Close();
        }


        private void AddFile()
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
 
        private void WriteProjectName()
        {
            ConsoleWriter.WriteLine(targetProject_.Name);
        }
        
        private void WriteCurrentProjectName()
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

        private System.String GetStartUpProjectName ()
        {    
            System.Array stringArray = targetDTE_.Solution.SolutionBuild.StartupProjects as System.Array;
            return System.IO.Path.GetFileNameWithoutExtension(stringArray.GetValue(0).ToString());    
        }

        private void SetStartUpProject()
        {
            if(targetProject_ == null){
                return;
            }
            targetDTE_.Solution.SolutionBuild.StartupProjects = targetProject_.FullName;

        }

        private void WriteProjectList()
        {
            try{
                for(int i = 0; i < targetDTE_.Solution.Projects.Count; i++){
                    Project project = targetDTE_.Solution.Projects.Item(i + 1);
                    ConsoleWriter.WriteLine(project.Name); 
                }
            }catch{

            }
        }

        private void WriteStartUpProjectName()
        {
            ConsoleWriter.WriteLine(GetStartUpProjectName());    
        }

        private void WriteSolutionDirectory ()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetDirectoryName(targetDTE_.Solution.FullName));
        }

        private void WriteSolutionName()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetFileNameWithoutExtension(targetDTE_.Solution.FullName));
        }

        private void WriteSolutionFileName()
        {
            ConsoleWriter.WriteLine(System.IO.Path.GetFileName(targetDTE_.Solution.FullName));
        }
        
        private void WriteSolutionFullPath()
        {
            ConsoleWriter.WriteLine(targetDTE_.Solution.FullName);
        }

        private void WriteBuildStatus()
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

        private void GoToDefinition()
        {
            if(targetProjectItem_ == null){
                ConsoleWriter.WriteDebugLine(FileFullPath.ToString () + " can not be found. please check option -f");
                return;
            }

            //念のため開く
            if(targetProjectItem_.IsOpen == false){
                targetProjectItem_.Open();
            }

            targetProjectItem_.Document.Activate();
            if((targetProjectItem_.Document.Selection is TextSelection) == false){
                return;
            }

    
            TextDocument textDocumet = targetProjectItem_.Document as TextDocument;
            TextSelection textSelection = targetProjectItem_.Document.Selection as TextSelection;
            textSelection.MoveToDisplayColumn(this.Line, this.Column);
           
           
            System.String findWhat = FindWhat;
            if(targetProjectItem_.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageVC || 
               targetProjectItem_.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageMC){

                if(System.String.IsNullOrEmpty(findWhat) == true){
                    ConsoleWriter.WriteDebugLine("findwhat can not be found. please check option -fw");
                    return;
                }
            }else{
                findWhat = "";
            }
     
            targetDTE_.ExecuteCommand("Edit.GoToDefinition", findWhat);

            //ExecuteCommnadを終了まちする方法がわからん...orz
            System.Threading.Thread.Sleep(500);
            
        }

        private void GoToDeclaration()
        {
            if(targetProjectItem_ == null){
                ConsoleWriter.WriteDebugLine(FileFullPath.ToString () + " can not be found. please check option -f");
                return;
            }

            //念のため開く
            if(targetProjectItem_.IsOpen == false){
                targetProjectItem_.Open();
            }

            targetProjectItem_.Document.Activate();
            if((targetProjectItem_.Document.Selection is TextSelection) == false){
                return;
            }

    
            TextDocument textDocumet = targetProjectItem_.Document as TextDocument;
            TextSelection textSelection = targetProjectItem_.Document.Selection as TextSelection;
            textSelection.MoveToDisplayColumn(this.Line, this.Column);
           
           
            System.String findWhat = FindWhat;
            if(targetProjectItem_.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageVC || 
               targetProjectItem_.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageMC){

                if(System.String.IsNullOrEmpty(findWhat) == true){
                    ConsoleWriter.WriteDebugLine("findwhat can not be found. please check option -fw");
                    return;
                }
            }else{
                ConsoleWriter.WriteDebugLine("GoToDeclaration is only used in C/C++");
                return;
            }
     
            targetDTE_.ExecuteCommand("Edit.GoToDeclaration", findWhat);

            //ExecuteCommnadを終了まちする方法がわからん...orz
            System.Threading.Thread.Sleep(500);
        }

        private void WriteLanguageType()
        {
            if(targetProjectItem_ == null){
                ConsoleWriter.WriteDebugLine(FileFullPath.ToString () + " can not be found. please check option -f");
                return;
            }
            ConsoleWriter.WriteLine(GetLanguageType(targetProjectItem_));
        }

        private System.String GetLanguageType(ProjectItem projectItem)
        {

            System.String language = "unknown";
            try{
                if(projectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageVC || 
                   projectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageMC){
                
                  language = "CPlusPlus";
                }else if(projectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageCSharp){
                    language = "CSharp";
                }else if(projectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageVB){
                    language = "VisualBasic";
                }else if(projectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageIDL){
                    language = "IDL";
                }
            }catch{
            }

            return language;
        }
    }
}

