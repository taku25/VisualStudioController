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
        
        
#region ÇﬂÇÒÇŒ
        static private System.String VisualStudioProcessName = @"!VisualStudio.DTE";
        DTE targetDTE_;
#endregion

        public bool Initialize(System.String targetName)
        {
            targetDTE_ = GetDTEFromSolutionName(targetName);
            if(targetDTE_ == null && System.String.IsNullOrEmpty(targetName) == true ){
                targetDTE_ = GetDTEFromEditFileName(targetName);
            }

            if(targetDTE_ == null){
                ConsoleWriter.WriteDebugLine(@"VisualStudioÇÃÉvÉçÉZÉXÇ™å©Ç¬Ç©ÇËÇ‹ÇπÇÒ");
            }
            return targetDTE_ != null;
        }

        //Ç±ÇÃà¯êîÇﬂÇ¡ÇøÇ·Ç´Ç‡ÇøÇÌÇÈÇ¢Ç»ÇüÅ` Ç≥ÇƒÇ«Ç§Ç∑Ç¡Ç©Ç»ÇüÅ`..
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
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFile) == true){
                GetCurrentEditFileName();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetOutput) == true){
                GetOutputWindowText();
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindResult1) == true){
                GetFindResultWindowText(0);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindResult2) == true){
                GetFindResultWindowText(1);
            }else if(argsConvert.Commnad(ArgsConvert.CommandType.GetFindSimbolResult) == true){
                GetFindResultWindowText(2);
            }
        }


        private DTE GetDTEFromEditFileName(String name)
        {
            name = name.ToLower();
            System.Collections.Generic.List<DTE> dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
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

        private DTE GetDTEFromSolutionName(String name)
        {
            System.Collections.Generic.List<DTE> dtelist = GetDTEFromProcessListFromName(VisualStudioProcessName);
            //Ç∆ÇËÇ†Ç¶Ç∏Ç¢Ç¡Ç±ÇﬂÇÇ©Ç¶Ç∑
            if(System.String.IsNullOrEmpty(name) == true){
                foreach(DTE dte in dtelist){
                    return dte;
                }
            }
            name = name.ToLower();
            foreach(DTE dte in dtelist){
                //êÊì™Ç©ÇÁàÍïîï™ÇÃÇ–Ç©Ç≠Ç≈Ç‡okÇ…Ç∑ÇÈ
                //ç≈èâÇ…äÆëSàÍívÇ©Ç«Ç§Ç©
                System.String solutionFullName = System.IO.Path.GetFileNameWithoutExtension (dte.Solution.FullName).ToLower();
                if(solutionFullName == name){
                    return dte;
                }
                //éüÇ…êÊì™Ç©ÇÁìØÇ∂Ç©Ç«Ç§Ç©
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

        public System.String GetCurrentEditFileName()
        {
            EnvDTE.Document document = targetDTE_.ActiveDocument;
            ConsoleWriter.WriteLine(document.FullName);
            return document.FullName;
        }

        public void GetOutputWindowText ()
        {
            OutputWindow outputWindow = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindOutput).Object as OutputWindow;
            EnvDTE.TextDocument textDocument = outputWindow.ActivePane.TextDocument;
            textDocument.Selection.SelectAll();
            System.String outputString = textDocument.Selection.Text;
            ConsoleWriter.WriteLine(outputString);
        }
        
        public void GetFindResultWindowText (int type)
        {
            Window window = null;
            if(type == 0){
               window = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindResults1);
            }else if(type == 1){
               window = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindResults2);
            }
            if(window == null){
                ConsoleWriter.WriteDebugLine("åüçıåãâ WindowÇ™Ç›Ç¬Ç©ÇËÇ‹ÇπÇÒÇ≈ÇµÇΩ");
                return;
            }
            TextSelection textSelection = window.Selection as TextSelection;
            textSelection.SelectAll();
            System.String outputString = textSelection.Text;
            ConsoleWriter.WriteLine(outputString);
        }
    }
}

