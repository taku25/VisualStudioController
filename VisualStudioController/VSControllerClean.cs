using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;


namespace VisualStudioController {
    public partial class VSController {

           
        private void CleanSolution()
        {
            targetDTE_.Solution.SolutionBuild.Clean(true);
        }

        private void CleanProject()
        {
            CleanProject(true);
        }

        private void CleanProject(bool changeProjectMark)
        {
            if(changeProjectMark == true){
                SetBuildMarkProjectBuildInfo(targetProject_.UniqueName, projectBuildInfoList_);
            }

            targetDTE_.Solution.SolutionBuild.Clean(true);

            if(changeProjectMark == true){
                RestoreProjectBuildInfo(projectBuildInfoList_);
            }
        }
    }
}
