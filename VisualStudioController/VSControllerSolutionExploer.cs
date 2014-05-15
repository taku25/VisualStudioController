using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;

namespace VisualStudioController {
    public partial class VSController {
        private class SolutionInfo
        {
            private System.Collections.Generic.List<System.String> infoList_ = new System.Collections.Generic.List<string>();

            public void AddInfo(System.String typeId, System.String value)
            {
                infoList_.Add(typeId + "=" + value);
            }

            public System.String Info
            {
                get {
                    return System.String.Join("*", infoList_);
                }
            }

            public System.String GetValue(System.String typeId)
            {
                foreach(System.String value in infoList_){
                    if(value.IndexOf(typeId + "=") == 0){
                        return value.Replace(typeId + "=", "");
                    }
                }
                return "";
            }
        };

        
        private System.String CreateSolutionExploerNodeFullId(System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList)
        {
            System.Collections.Generic.List<System.String> idList = new List<string>();

            foreach(EnvDTE.UIHierarchyItem item in hierarchyItemList){
                idList.Add(item.Name);
            }

            return System.String.Join("|", idList);
        }

        private EnvDTE.UIHierarchyItem getUIHierarchyItemFromId_(System.String targetId, EnvDTE.UIHierarchyItem item, System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList)
        {

            System.String[] activeIdArray = targetId.Split('|');

            if(activeIdArray[0] != item.Name){
                return null;
            }
            
            hierarchyItemList.Add(item);
            if(targetId == item.Name){
                return item;
            }
          
            System.String nextTargetId = targetId.Replace(item.Name + "|", "");


            foreach(EnvDTE.UIHierarchyItem childItem in item.UIHierarchyItems){
                EnvDTE.UIHierarchyItem targetItem = getUIHierarchyItemFromId_(nextTargetId, childItem, hierarchyItemList);
                if(targetItem != null){
                    return targetItem;
                }
            }
            return null;
        }

        private EnvDTE.UIHierarchyItem GetUIHierarchyItemFromFullId(System.String fullId, System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList)
        {
            hierarchyItemList.Clear();
            if(uiHierarchy_.UIHierarchyItems.Count == 0){
                return null;
            }


            EnvDTE.UIHierarchyItem targetItem = null;
            foreach(EnvDTE.UIHierarchyItem item in uiHierarchy_.UIHierarchyItems){
                //fullIdが空なら一個目(root)を返す用に細工
                if(System.String.IsNullOrEmpty(fullId) == true){
                    fullId = item.Name;
                }

                targetItem = getUIHierarchyItemFromId_(fullId, item, hierarchyItemList);
                if(targetItem != null){
                    break;
                }
            }

            return targetItem;
        }


        private void WriteSolutionExploerNodeInfo()
        {
            System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList = new List<UIHierarchyItem> ();
            EnvDTE.UIHierarchyItem item = GetUIHierarchyItemFromFullId(TargetNode, hierarchyItemList);
            WriteSolutinExploerNodeInfo(item, CreateSolutionExploerNodeFullId(hierarchyItemList));
        }
        

        private void WriteSolutinExploerNodeInfo(EnvDTE.UIHierarchyItem uiItem, System.String fullId)
        {
            SolutionInfo solutionInfo = new SolutionInfo();
            solutionInfo.AddInfo("ID", uiItem.Name);
            solutionInfo.AddInfo("FULLID", fullId);
            solutionInfo.AddInfo("HASCHILD", uiItem.UIHierarchyItems.Count == 0 ? "FALSE" : "TRUE");
            
            if(uiItem.Object is EnvDTE.Solution){
                EnvDTE.Solution solution = uiItem.Object as Solution;
                solutionInfo.AddInfo("FULLFILEPATH", solution.FullName);
                solutionInfo.AddInfo("TYPE", "FILE");
            }else if(uiItem.Object is EnvDTE.Project){
                EnvDTE.Project project = uiItem.Object as Project;
                solutionInfo.AddInfo("FULLFILEPATH", project.FullName);
                solutionInfo.AddInfo("TYPE", "FILE");
            }else if(uiItem.Object is EnvDTE.ProjectItem){
                EnvDTE.ProjectItem projectItem = uiItem.Object as ProjectItem;
                if (projectItem.Kind == Constants.vsProjectItemKindPhysicalFile){
                    solutionInfo.AddInfo("FULLFILEPATH", projectItem.FileNames[0]);
                    solutionInfo.AddInfo("TYPE", "FILE");
                }else if(projectItem.Kind == Constants.vsProjectItemKindPhysicalFolder){
                    solutionInfo.AddInfo("FULLFILEPATH", projectItem.FileNames[0]);
                    solutionInfo.AddInfo("TYPE", "DIRECTORY");
                }else{
                    solutionInfo.AddInfo("TYPE", "VIRTUALDIRECTORY");
                }
            }else if(uiItem.Object is UIHierarchyItem){
                solutionInfo.AddInfo("TYPE", "OTHER");                
            }else{
                ConsoleWriter.WriteDebugLine("unknown type is " + uiItem.ToString());
            }

            ConsoleWriter.WriteLine(solutionInfo.Info);
        }

        private void WriteSolutionExploerChildrenNodeInfo()
        {
            System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList = new List<UIHierarchyItem> ();
           
            EnvDTE.UIHierarchyItem parentItem = GetUIHierarchyItemFromFullId(TargetNode, hierarchyItemList);
            if(parentItem == null){
                return;
            }

            System.String parentFullId = CreateSolutionExploerNodeFullId(hierarchyItemList);
            foreach(EnvDTE.UIHierarchyItem item in parentItem.UIHierarchyItems){

                WriteSolutinExploerNodeInfo(item, parentFullId + "|" + item.Name);
            }
        }

        private void DoActionSolutionExploerNode()
        {
            System.Collections.Generic.List<EnvDTE.UIHierarchyItem> hierarchyItemList = new List<UIHierarchyItem> ();
            EnvDTE.UIHierarchyItem parentItem = GetUIHierarchyItemFromFullId(TargetNode, hierarchyItemList);
            parentItem.Select(vsUISelectionType.vsUISelectionTypeSelect);
            
            uiHierarchy_.DoDefaultAction();
        }

        /*private void WriteSolutionExplorerInfo_(EnvDTE.UIHierarchyItem uiItem, System.String parentId = "")
        {
            SolutionInfo solutionInfo = new SolutionInfo();
            solutionInfo.AddInfo("NAME", uiItem.Name);

            if(System.String.IsNullOrEmpty(parentId) == false){
                solutionInfo.AddInfo("PARENTID", parentId);
            }

            if(uiItem.Object is EnvDTE.Solution){
                EnvDTE.Solution solution = uiItem.Object as Solution;
                solutionInfo.AddInfo("FULLPATH", solution.FullName);
                solutionInfo.AddInfo("ID", solution.ExtenderCATID);
                solutionInfo.AddInfo("TYPE", "SOLUTION");
            }else if(uiItem.Object is EnvDTE.Project){
                EnvDTE.Project project = uiItem.Object as Project;
                solutionInfo.AddInfo("FULLPATH", project.FullName);
                solutionInfo.AddInfo("ID", project.ExtenderCATID);
                solutionInfo.AddInfo("TYPE", "PROJECT");
            }else if(uiItem.Object is EnvDTE.ProjectItem){
                EnvDTE.ProjectItem projectItem = uiItem.Object as ProjectItem;

                if (projectItem.Kind == Constants.vsProjectItemKindPhysicalFile){
                    solutionInfo.AddInfo("FULLPATH", projectItem.FileNames[0]);
                    solutionInfo.AddInfo("ID", projectItem.ExtenderCATID);
                    solutionInfo.AddInfo("TYPE", "FILE");
                }else if(projectItem.Kind == Constants.vsProjectItemKindPhysicalFolder){
                    solutionInfo.AddInfo("FULLPATH", projectItem.FileNames[0]);
                    solutionInfo.AddInfo("ID", projectItem.ExtenderCATID);
                    solutionInfo.AddInfo("TYPE", "FOLDER");
                }else if(projectItem.Kind == Constants.vsProjectItemKindVirtualFolder){
                    solutionInfo.AddInfo("ID", projectItem.ExtenderCATID);
                    solutionInfo.AddInfo("TYPE", "VIRTUALFOLDER");
                }else{

                }
            }else{
                ConsoleWriter.WriteDebugLine("unknown object : " + uiItem.Object.ToString());
            }

            ConsoleWriter.WriteLine(solutionInfo.GetInfo());
            
            foreach(EnvDTE.UIHierarchyItem item in uiItem.UIHierarchyItems){
                writeSolutionExplorerInfo_(item, solutionInfo.GetValue("ID"));
            }*/
        
    }
}
