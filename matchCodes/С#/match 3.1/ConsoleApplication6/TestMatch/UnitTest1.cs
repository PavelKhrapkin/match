/*-----------------------------------------------------------------------
 * UnitTest -- Unit тесты проекты match 3.1
 * 
 *  1.02.2015  П.Храпкин, А.Пасс
 * ------------------------------------------- 
 * 18.01.15 test_get_set_Matr() - проверка работы Matrix Matr, get и set
 * 25.01.15 test_WrCSV_WrReport() - проверка WrCRV и WrReport
 *    2014  test_ToStrList()    - проверка MatchLib.TpStrList
 * 19.01.15 test_CheckSum()     - проверка вычисления контрольной суммы
 *  1.02.15 test_FileOp()       - проверка FileOP
 * 31.01.15 test_getDoc(name)   - проверка загрузки Документа name 
 * 22.01.15 test LstInit()      - проверка создания List Accounts по SFacc
 * 18.01.15 test_AddRow()       - проверка AddRow - добавления 1 строки
 * test_ToStrList()

 *  1.02.15 test_Fetch()        - проверка Fetch - извлечение данных по запросу вида SFacc/2:3/0
 !    2014  testxml()           - заткнут
 */
using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

using match;
using Lib = match.Lib;
using match.Lib;
using Mtr = match.Matrix.Matr;
using CS = match.Lib.CS;
using Log = match.Lib.Log;
using Docs = match.Document.Document;
using Lst = match.Document.Lst;
using FileOp = match.MyFile.FileOpenEvent;
using Decl = match.Declaration.Declaration;
//using Handl = match.Handler.Handler;

namespace TestMatch
{
    [TestClass]
    public class TestMatchLib
    {
        [TestMethod]
        // 18.01.2015  проверка работы Matrix Matr, get и set и индексера
        public void test_get_set_Matr()
        {
            object [,] init = { { 1, 2 }, { 3, 4 }, { 5, 6 } };
            // то есть структуры init инициалицзируется
            //      1 3 5
            //      2 4 6
            Assert.AreEqual(init.Length , 6);
            Assert.AreEqual(init.GetLength(0), 3);
            Assert.AreEqual(init.GetLength(1), 2);

            Mtr xx = new Mtr(init);
            bool eq = xx.Equals(init);
            Assert.AreEqual(eq, false);
            Assert.AreEqual(xx.iEOL(), 3);  // размеры xx и init совпадают,
            Assert.AreNotEqual(xx, init);   //.. но это разные объекты!
            Mtr yy = new Mtr(init);
            Assert.AreEqual(xx.Compare(yy), true);  //и Compare говорит - они равны.

            xx[2, 1] = 1230;                       // изменим xx
            Assert.AreEqual(xx.Compare(yy), true); //..и так же изменился yy
                                                   //..т.к. в памяти это init
        }
        [TestMethod]
        // 25.1.2015
        public void test_WrCSV_WrReport()
        {
            object[,] init = { { "H1","H2","H3" }, { "4","5","6" }, { "7","8","9" }, {"ж10","ш11","я12"} };
            Mtr xx = new Mtr(init);
            DataTable dt = new DataTable();
            dt = xx.DaTab();
            Assert.AreEqual(dt.Rows.Count, xx.iEOL());
            Assert.AreEqual(dt.Columns.Count, xx.iEOC());

            string col1name = dt.Columns[0].ColumnName;
            string col2name = dt.Columns[1].ColumnName;
            string col3name = dt.Columns[2].ColumnName;
        // ---- сортировка DataTable
            DataView dv = dt.DefaultView;
            dv.Sort = col1name + " desc";
            DataTable sdt = dv.ToTable();

            FileOp.WrCSV("test", sdt);
            FileOp.WrReport("test", sdt);
        }
        [TestMethod]
        public void test_ToStrList()
        {
            var strs = Lib.MatchLib.ToStrList("начало, продолжение, конец");
            Assert.AreEqual(strs.Count, 3);
            Assert.AreEqual(strs[0], "начало");
        }
        [TestMethod]
        // 19.1.2015
        public void test_CheckSum()
        {
            Docs doc = Docs.getDoc("Платежи");
            Assert.AreNotEqual(null, doc);
            Assert.AreEqual(doc.name, "Платежи");
            double iSum = Lib.CS.CheckSum(doc);
            Assert.AreNotEqual(iSum, 0);
            Assert.AreEqual(iSum > 0, true);
            doc = Docs.getDoc("SFopp");             //теперь загружаем другой документ
            Assert.AreEqual(doc.name, "SFopp");
            double iSumOpp = Lib.CS.CheckSum(doc);  //..и пересчитываем его контр.сумму
            Assert.AreNotEqual(iSum, iSumOpp);
            doc = Docs.getDoc("Платежи");           //.. и опять Платежи
            double reloadedSum = Lib.CS.CheckSum(doc);  //..сравниваем контр.суммы
            Assert.AreEqual(reloadedSum, iSum);
            doc.Body[12, 11] = "asdFghj";       // изменим ясейку в фале "Платежи"
            iSum = Lib.CS.CheckSum(doc);        //..и опять пересчитываем контр.сумму
            Assert.AreNotEqual(reloadedSum, iSum); //  теперь они не равны!  
        }

    }

    [TestClass]
    public class UnitTest
    {
        [TestMethod]
        // 1/02/2015
        public void test_FileOp()
        {
            Excel.Workbook Wb = FileOp.fileOpen(Decl.F_1C);
            bool isSheet = FileOp.sheetExists(Wb, "Платежи");
            Assert.AreEqual(isSheet, true);
            isSheet = FileOp.sheetExists(Wb, "Платеж");
            Assert.AreEqual(isSheet, false);
            isSheet = FileOp.sheetExists(Wb, "Договоры");
            Assert.AreEqual(isSheet, true);
            var doc = Docs.getDoc("Договоры");
            var doc1 = Docs.getDoc("Платежи");
            var doc2 = Docs.getDoc("SFacc");
            var doc3 = Docs.getDoc("SF");
            FileOp.Quit();
        }
        [TestMethod]
        // 31/1/2015
        public void test_getDoc ()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreNotEqual(null, doc);
            Assert.AreEqual(doc.name, "Платежи");
            doc = Docs.getDoc("ABCD");
            Assert.AreEqual(doc, null); // этот Документ не должен быть найден
            doc = Docs.getDoc("SFacc");
            Assert.AreEqual(doc.name, "SFacc");
            var doc1 = Docs.getDoc("SF");       // последовательно открываем SFacc и SF
            Assert.AreEqual(doc1.name, "SF");   //..тут были ошибки, связанные с отсутствием SF
            FileOp.Quit();
        }
        [TestMethod]
        //22.01.20
        public void test_LstInit()
        {
            Lst.Init(Lst.Entity.Accounts);
            Assert.AreEqual(Lst.Accounts.Count > 1000, true);
            Assert.AreEqual(Lst.Acc1Cs.Count, 0);
            Assert.AreEqual(Lst.Contracts.Count, 0);
            Assert.AreEqual(Lst.Opps.Count, 0);
            Assert.AreEqual(Lst.Pays.Count, 0);
            FileOp.Quit();
        }
        [TestMethod]
        public void test_load_recognize_Doc()
        {
            Excel.Workbook Wb = FileOp.fileOpen("PP.xlsx");
            Assert.AreNotEqual(null, Wb);
            string newDocName = Docs.recognizeDoc(Wb);
            Assert.AreNotEqual(null, newDocName);
            new Log("Входной файл распознан как Документ \"" + newDocName + "\"");
            Docs doc = Docs.loadDoc(newDocName, Wb);
            Assert.AreNotEqual(null, doc);
        }
        [TestMethod]
        // 1/02/2015 -- не идет, отлаживаю!
        public void test_AddLine()
        {
            Docs pays = Docs.getDoc("Платежи");
            Docs newPays = Docs.getDoc("NewPayment");
            newPays.FetchInit();
            pays.dt = pays.Body.DaTab();
            DataRow pmnt = pays.dt.Rows[2];
            object ext = "ddd";
            pays.AddLine(pmnt, newPays, ext);
            FileOp.Quit();
        }
        [TestMethod]
        // 19/01/2015
        public void test_AddRow()
        {
            var doc = Docs.getDoc("SF_DicAccSyn");
            Assert.AreEqual( doc.name, "SF_DicAccSyn");
            int lc = doc.Body.iEOL();
            Assert.AreEqual( lc > 1, true);

            doc.Body.AddRow();
            Assert.AreEqual( doc.Body.iEOL() - lc, 1);
////            reloadedSum = Lib.CS.CheckSum(docAcc);
            ////Assert.AreEqual(reloadedSum, iSum);
            ////string[] hh = { "One", "Two", "Three" };
            ////docAcc.Body.AddRow(hh);
            ////reloadedSum = Lib.CS.CheckSum(docAcc);
            ////Assert.AreNotEqual(reloadedSum, iSum); 

        }
        [TestMethod]
        // 1/02/2015
        public void test_Fetch()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreEqual("Платежи", doc.name);
            // string x = doc.Fetch(null, null);   //Fetch(null,null) --> Log.FATAL("параметр null");
            string x = doc.Fetch("", "пример1");
            Assert.AreEqual(x, "пример1");
            x = doc.Fetch("    ", "пример2"); 
            Assert.AreEqual(x, "пример2");
            doc.FetchInit();
            string id = doc.Fetch("SFacc/2:3/0", "ООО «ОРБИТА СПб»");
            Assert.AreEqual("001D000000m35wn", id);
            FileOp.Quit();
        }
        public void test_dt_Mtr()
        {

            //DateTime t0 = DateTime.Now;
            //DataTable dt = doc1.Body.DaTab();
            ////            string name = dt.TableName;
            ////            doc1.Wb.Worksheets["SF"].Delete();
            ////            dt.TableName = "SF1";
            //doc1.Wb.Worksheets.Add(dt);
            //DateTime t1 = DateTime.Now;
            //new Log(" DaTab  t=" + (t1 - t0));
            //Mtr mtr = new Mtr(dt);
            //new Log(" to Mtr t=" + (DateTime.Now - t1));
            //if (mtr != doc1.Body) Log.FATAL("увы!");

        }
        /// <summary>
        /// проверка экспорта в файл Excel
        /// </summary>
        /// <journal>
        /// 23.1.14
        /// </journal>
        //public void textExcelExport()
        //{
        //    Excel.Workbook Wb = FileOp.fileOpen("PP.xlsx");
        //    Excel.Worksheet Sh = Wb.Worksheets[1];

        //    string[,] arr = new string[44, 10000];
        //    for (int iSum = 0; iSum < 3; iSum++)
        //        for (int j = 0; j < 9; j++)
        //            arr[iSum, j] = "txt" + iSum + j;

        //    Sh.get_Range("A1", "DD1000").Value = arr;

        //    int iEOL = doc1.Body.iEOL();
        //    int iEOC = doc1.Body.iEOC();

        //    Excel.Range cll2 = Sh.Cells[iEOL, iEOC];
        //    Excel.Range rng = Sh.Range["A1", cll2];
        //    Sh.get_Range(rng).Value = doc1.Body;

        //    Sh.get_Range("A1", Sh.Cells[iEOL, iEOC]).Value = doc1.Body;
        //}
        [TestMethod]
        // 2014 - заткнут!
        public void testxml()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreNotEqual(null, doc);
            //var ser = new System.Xml.Serialization.XmlSerializer(typeof(Docs));
            //ser.Serialize(new StreamWriter("test.xml"), docAcc);
            FileOp.Quit();
        }
    }
}