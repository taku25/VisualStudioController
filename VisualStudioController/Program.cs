using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


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
             if(vsController.Initialize(argsConvert.TargetName, argsConvert.TargetProjectName) == false){
                 return;
             }
             vsController.Run(argsConvert);
        }
    }

}
