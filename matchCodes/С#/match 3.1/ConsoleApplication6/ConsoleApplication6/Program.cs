using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using Decl = match.Declaration.Declaration;
using Docs = match.Document.Document;
using Proc = match.Process.Process;
using FileOp = match.MyFile.FileOpenEvent;
using Mtr = match.Matrix.Matr;
using match.Lib;
using Log = match.Lib.Log;

namespace match
{
    class Program
    {
        static void Main(string[] args)
        {
            Log.START("match v3.1.2.1");

            NonDialogPass();

            //doc.isChanged = true; //ЗАТЫЧКА!!
            //doc.saveDoc();
            //Log.Save();
            //new Log("CheckSum =" + doc.CheckSum());

            //Console.ReadLine();
        }

        static void NonDialogPass()
        {
            // отладка -- потом убрать в UnitTest
            Docs docPay = Docs.getDoc("Платежи");
            docPay.FetchInit();

            Docs newPay = Docs.NewSheet("NewPayment");
            Docs neDogovor = Docs.NewSheet("NewContract");
        }
    }
}