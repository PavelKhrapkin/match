using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Box = System.Windows.Forms.MessageBox;
using Lib = match.MatchLib;

namespace ExcelAddIn2
{
    class Beginning
    {
        public void ActiveStart()
        {
            Box.Show("ActiveStart");
            //            Excel.Application app = new Excel.Application();
            //            Document contracts = new Document("Платежи");
        }
        
        public void test(Excel.Workbook Wb)
        {
            string name = Document.recognizeDoc(Wb);
            if (name == null) { return; }
            Document newDoc = Document.loadDoc(name, Wb);
//            Document contracts = new Document("Договоры");
            System.Windows.Forms.MessageBox.Show("Opening WB='"+Wb.Name+"' лист[1]='"+
                Wb.Sheets[1].Name+"'  строк="+ Lib.EOL(Wb.Sheets[1]));
//            Box.Show("Opening WB='"+Wb.Name+"' лист(Договоры)='"+
  //              Wb.Sheets["Договоры"].Name + "'  строк=" + Lib.EOL(Wb.Sheets["Договоры"]));
            ActiveStart();
        }
     }
}
