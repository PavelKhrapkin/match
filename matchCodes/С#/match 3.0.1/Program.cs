using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Dcs = match.Document.Document;
//using Proc = match.Process.Process;
using Fls = match.MyFile;
using match.Lib;
using Log = match.Lib.Log;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {

            DateTime _t = DateTime.Now;
            Console.WriteLine("{0}> match.Declaration 3.0.0.0 -- отладка", _t);

            //string docName = "Process111";
            //Dcs p = Dcs.getDoc(docName);
            //Excel.Range rng = p.Body.Range["A1:D3"];
            //List<string> s = MatchLib.ToStrList(rng);

            Log.set("Program");

            //List<string> tmp = new List<string>();
            //tmp = MatchLib.ToStrList("текст, и его фрагмент");
            //new Log("getDoc(" + docName + ")");

            Excel.Workbook Wb = Fls.FileOpenEvent.fileOpen("PP.xlsx");
            string newDocName = Dcs.recognizeDoc(Wb);
            Dcs doc = Dcs.loadDoc(newDocName, Wb);
            Log.exit();
            Console.ReadLine();
        }
    }
