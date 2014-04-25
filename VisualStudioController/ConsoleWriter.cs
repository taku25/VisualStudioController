using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisualStudioController{

    static public class ConsoleWriter
    {
        static private bool isDebugEnable_ = false;

        public enum Encoding{
            Default = 0,
            UTF8,
        };

        static public bool DebugEnable
        {
            set { isDebugEnable_ = value; }
            get { return isDebugEnable_; }
        }

        static public void SetEncoding(Encoding encoding)
        {
            if(encoding == Encoding.Default){
                System.Console.OutputEncoding = System.Text.Encoding.Default;
            }else if(encoding == Encoding.UTF8){
                System.Console.OutputEncoding = System.Text.Encoding.UTF8;
            }
        }


        static public void WriteDebugLine (System.String cString)
        {
            WriteLine_ ("debug:" + cString, true);
        }

        static public void WriteLine (System.String cString)
        {
            WriteLine_ (cString, false);
        }

        static public void Write (System.String cString)
        {
            Write_ (cString, false);
        }

        static private void WriteLine_ (System.String cString, bool debug)
        {
            if (debug == true && DebugEnable == false){
                return;
            }
            System.Console.WriteLine (cString);
        }

        static private void Write_ (System.String cString, bool debug)
        {
            if (debug == true && DebugEnable == false){
                return;
            }
            System.Console.Write (cString);
        }

    }
}
