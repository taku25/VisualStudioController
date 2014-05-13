using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;

namespace VisualStudioController {
    public partial class VSController {

        private void BuildSolution()
        {
            BuildSolution(false);
        }

        private void ReBuildSolution()
        {
            BuildSolution(true);
        }

        private void BuildSolution(bool rebuild)
        {
            if(rebuild == true){
                CleanSolution();
            }
            targetDTE_.Solution.SolutionBuild.Build(IsWait);    
        }
        
        private void BuildProject()
        {
            BuildProject(false);
        }

        private void ReBuildProject()
        {
            BuildProject(true);
        }

        private void BuildProject(bool rebuild)
        {
            if(rebuild == true){
                CleanProject(true);
            }
            
            targetDTE_.Solution.SolutionBuild.BuildProject(targetDTE_.Solution.SolutionBuild.ActiveConfiguration.Name, targetProject_.UniqueName, IsWait);

        }

        private void CompileFile()
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

        private void CancelBuild()
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


    }
}
