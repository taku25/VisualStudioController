using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;

namespace VisualStudioController {

    public partial class VSController : System.IDisposable{


        private void WriteFindResultWindowText1 ()
        {
            WriteFindResultWindowText(0);
        }

        private void WriteFindResultWindowText2 ()
        {
            WriteFindResultWindowText(1);
        }

        private void WriteFindResultWindowText (int type)
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


        private void WriteFindSymbolResultWindowText ()
        {
            Window findSymbolWindow = targetDTE_.Windows.Item(EnvDTE.Constants.vsWindowKindFindSymbolResults);
            if(findSymbolWindow == null){
                ConsoleWriter.WriteDebugLine("検索結果Windowがみつかりませんでした");
                return;
            }
            
            //クリップボードをいじるので退避
            System.Windows.Forms.IDataObject clipboarData = System.Windows.Forms.Clipboard.GetDataObject();

            try{
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[0-9]+");
                if(regex.IsMatch(findSymbolWindow.Caption) == false){
                    ConsoleWriter.WriteDebugLine("symbol not found");
                    return;
                }

                int resultCount = Convert.ToInt32(regex.Match(findSymbolWindow.Caption).Value);


                for (int i = 0; i < resultCount; i++){

                    findSymbolWindow.Activate();
                    targetDTE_.ExecuteCommand("Edit.GoToNextLocation");

                    //GoToNextLocationを送ると非アクティブになるのでそれをチェックする(シンボルウインドウがアクティブ = 下のTextSelectionがうまく取れない)
                    for(;;){
                        if(targetDTE_.ActiveWindow != findSymbolWindow){
                            break;
                        }
                    }
                    targetDTE_.ExecuteCommand("Edit.Copy");

                    try{
                        TextSelection textSelection = targetDTE_.ActiveDocument.Selection as TextSelection;

                        CodeElement codeElementFunction = textSelection.ActivePoint.get_CodeElement(vsCMElement.vsCMElementFunction) as CodeElement;
                        CodeElement codeElementProperty = null;
                        CodeElement tempElement = null;

                        System.String referencePointCode = System.Windows.Forms.Clipboard.GetText();
                        referencePointCode = referencePointCode.Replace("\r\n", "");
                        referencePointCode = referencePointCode.Replace("\n", "");
                        referencePointCode = referencePointCode.Replace("\r", "");
                        
                        if (codeElementFunction == null){
                            codeElementProperty = textSelection.ActivePoint.get_CodeElement(vsCMElement.vsCMElementProperty) as CodeElement;
                        }else if (codeElementFunction.FullName.Length == 0){
                            codeElementProperty = textSelection.ActivePoint.get_CodeElement(vsCMElement.vsCMElementProperty) as CodeElement;
                        }

                        tempElement = codeElementFunction != null ? codeElementFunction : codeElementProperty;

                        System.String lineInfo = "";
                        if(targetDTE_.ActiveWindow.ProjectItem.ContainingProject.CodeModel.Language == EnvDTE.CodeModelLanguageConstants.vsCMLanguageCSharp){
                            lineInfo = textSelection.ActivePoint.Line.ToString() + "," + textSelection.ActivePoint.DisplayColumn.ToString();

                        }else{
                            lineInfo = tempElement.StartPoint.Line.ToString ();
                        }

                        System.String tempValue = targetDTE_.ActiveDocument.FullName + "(" + lineInfo + "):"+ " reference: " + referencePointCode;

                        ConsoleWriter.WriteLine(tempValue);
                    }catch{
                    }
                }
            }catch (System.Exception e){
                ConsoleWriter.WriteLine(e.ToString());
            }

            try{
                System.Windows.Forms.Clipboard.SetDataObject(clipboarData);
            }catch{

            }

        }
        

        private void Find()
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
        }

        private void FindSymbol()
        {
            if(targetDTE2_ == null){
                ConsoleWriter.WriteDebugLine("");
                return;
            }
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

            if(System.String.IsNullOrEmpty(FindWhat) == true){
                return;
            }
    
            TextDocument textDocumet = targetProjectItem_.Document as TextDocument;
            TextSelection textSelection = targetProjectItem_.Document.Selection as TextSelection;
            textSelection.MoveToDisplayColumn(this.Line, this.Column);



            bool finddone = false;
            if(IsWait){
                targetDTE2_.Events.FindEvents.FindDone += (vsFindResult Result, bool Cancelled) =>
                {
                    ConsoleWriter.WriteDebugLine("おわた");
            
                    finddone = true;
                };
            }
   
            targetDTE_.ExecuteCommand("Edit.FindSymbol", FindWhat);
            
            if(IsWait){
                for(;;){
                    if(finddone == true){
                        break;
                    }
                }
            }
        }

        void FindEvents_FindDone(vsFindResult Result, bool Cancelled) {
            throw new NotImplementedException();
        }

    }
}

