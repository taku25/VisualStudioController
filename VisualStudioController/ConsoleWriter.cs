using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisualStudioController{

    static public class ConsoleWriter
    {
        static private bool isDebugEnable_ = false;

        static public bool DebugEnable
        {
            set { isDebugEnable_ = value; }
            get { return isDebugEnable_; }
        }


        static public void WriteDebugLine (System.String cString)
        {
            WriteLine_ (cString, true);
        }

        static public void WriteLine (System.String cString)
        {
            WriteLine_ (cString, false);
        }

        static private void WriteLine_ (System.String cString, bool debug)
        {
            if (debug == true && DebugEnable == false){
                return;
            }
            System.Console.WriteLine (cString);
        }

    }
}
