using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Dcs = match.Document.Document;
using Proc = match.Process.Process;
using Fls = match.MyFile;
using match.Lib;
using Log = match.Lib.Log;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {

            Log.START("match v3.0.1.12");

            Proc.Reset("LOAD_SF_DicAccSyn");  //позже вернемся в вопросу о месте для констанся - имени Процессов



            Excel.Workbook Wb = Fls.FileOpenEvent.fileOpen("PP.xlsx");
            string newDocName = Dcs.recognizeDoc(Wb);
            new Log("Входной файл распознан как Документ \"" + newDocName + "\"");
            Dcs doc = Dcs.loadDoc(newDocName, Wb);
            Console.ReadLine();
        }
    }
}