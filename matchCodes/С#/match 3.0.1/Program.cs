using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Docs = match.Document.Document;
using Proc = match.Process.Process;
using Ftch = match.Fetch.Fetch;
using Fls = match.MyFile;
using match.Lib;
using Log = match.Lib.Log;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {
            Log.START("match v3.0.1.22");
            
            Docs doc1 = Docs.getDoc("Платежи");
 //           long t = test(doc1);
//            doc1.FetchInit("SFacc/2:3");
            doc1.FetchInit();

            //            Proc.Reset("LOAD_SF_DicAccSyn");  //позже вернемся в вопросу о месте для константы - имени Процессов

            Excel.Workbook Wb = Fls.FileOpenEvent.fileOpen("PP.xlsx");
            string newDocName = Docs.recognizeDoc(Wb);
            new Log("Входной файл распознан как Документ \"" + newDocName + "\"");
            Docs doc = Docs.loadDoc(newDocName, Wb);
            Console.ReadLine();
        }

        private static long test(Docs doc)
        {
            DateTime t0 = DateTime.Now;
            long checkSum = 0;
            int colCount = doc.Sheet.UsedRange.Columns.Count;
            object[,] bdy = (object[,])doc.Sheet.UsedRange.get_Value();

            foreach (var cll in bdy)
            {
                if (cll == null) continue;
                string str = cll.ToString().Trim();
                byte[] bt = Encoding.ASCII.GetBytes(str);
                foreach (var h in bt) checkSum += h;
            }

            DateTime t1 = DateTime.Now;
            new Log("-> " + (t1 - t0) + "\tChechSum=" + checkSum);

            return checkSum;
        }
    }
}