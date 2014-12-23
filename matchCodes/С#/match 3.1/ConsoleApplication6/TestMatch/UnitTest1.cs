using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using match;
using Lib = match.Lib;
using match.Lib;
using Log = match.Lib.Log;
using Docs = match.Document.Document;
using FileOp = match.MyFile.FileOpenEvent;

namespace TestMatch
{
    [TestClass]
    public class TestMatchLib
    {
        [TestMethod]
        public void test_ToStrList()
        {
            var strs = Lib.MatchLib.ToStrList("начало, продолжение, конец");
            Assert.AreEqual(strs.Count, 3);
            Assert.AreEqual(strs[0], "начало");
        }
    }

    [TestClass]
    public class UnitTest
    {
        [TestMethod]
        public void testgetDoc ()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreNotEqual(null, doc);
            Assert.AreEqual(doc.name, "Платежи");
            doc = Docs.getDoc("ABCD");
            Assert.AreEqual(doc, null); // этот Документ не должен быть найден
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

        public void test_ProcReset()
        {
            //            Proc.Reset("LOAD_SF_DicAccSyn");  //позже вернемся в вопросу о месте для константы - имени Процессов

        }
        [TestMethod]
        public void test_Fetch()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreEqual("Платежи", doc.name, false);
            doc.FetchInit();
            string id = doc.Fetch("SFacc/2:3/0", "ООО «ОРБИТА СПб»");
            Assert.AreEqual("001D000000m35wn", id, false);
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
        //    for (int i = 0; i < 3; i++)
        //        for (int j = 0; j < 9; j++)
        //            arr[i, j] = "txt" + i + j;

        //    Sh.get_Range("A1", "DD1000").Value = arr;

        //    int iEOL = doc1.Body.iEOL();
        //    int iEOC = doc1.Body.iEOC();

        //    Excel.Range cll2 = Sh.Cells[iEOL, iEOC];
        //    Excel.Range rng = Sh.Range["A1", cll2];
        //    Sh.get_Range(rng).Value = doc1.Body;

        //    Sh.get_Range("A1", Sh.Cells[iEOL, iEOC]).Value = doc1.Body;
        //}
        [TestMethod]
        public void testxml()
        {
            var doc = Docs.getDoc("Платежи");
            Assert.AreNotEqual(null, doc);
            //var ser = new System.Xml.Serialization.XmlSerializer(typeof(Docs));
            //ser.Serialize(new StreamWriter("test.xml"), doc);
        }
    }
}