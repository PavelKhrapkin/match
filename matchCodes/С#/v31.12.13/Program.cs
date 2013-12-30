using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Dcs = match.Document.Document;
using Fls = match.MyFile;
using match.Lib;
using Log = match.Lib.Log;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {
            string docName = "Process";
            DateTime _t = DateTime.Now;
            Console.WriteLine("{0}> match.Declaration 3.0.0.0 -- отладка", _t);

            Log.set("Program");
            new Log("getDoc(" + docName + ")");

            Excel.Workbook Wb = Fls.FileOpenEvent.fileOpen("PP.xlsx");
            string newDocName = Dcs.recognizeDoc(Wb);
            Dcs doc = Dcs.loadDoc(newDocName, Wb);
            Log.exit();
            Console.ReadLine();
        }
    }
}