using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EnvDTE;


// commnadてすと
// build -t CTest
// build -t CTest -d
// build -t CTest -w -d
// rebuild -t CTest -w -d
// getfile -t CTest -d
// getoutput -t CTest -d
// getfindresult1 -t CTest -d
// ↓まだ取得できない
// getfindsimbolresult -t CTest -d
namespace VisualStudioController {
    public class Program
    {
        [System.STAThreadAttribute()]
        static void Main(string[] args)
        {
            ArgsConvert argsConvert = new ArgsConvert ();
            if (argsConvert.Analysis (args) == false){
                return;
            }

             VSController vsController = new VSController();
             if(vsController.Initialize(argsConvert.TargetName) == false){
                 return;
             }
             vsController.Run(argsConvert);
        }
    }

}
