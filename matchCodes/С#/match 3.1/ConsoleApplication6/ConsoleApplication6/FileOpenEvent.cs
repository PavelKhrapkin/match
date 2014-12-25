//#define Test

using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using match.Document;
using match.Lib;
using Mtr = match.Matrix.Matr;
using Decl = match.Declaration.Declaration;

namespace match.MyFile
{
    public class FileOpenEvent
    {
        public static string dirDBs = null;

        private static Excel.Application _app = null;
        /*
                /// <summary>
                /// Open Workbook Event Handler
                /// </summary>
                /// <param name="Wb"></param>
                /// <journal>   15.12.2013
                /// </journal>
                public void newFile(Excel.Workbook Wb)
                {
                    _app = Wb.Application;
                    string name = Document.recognizeDoc(Wb);
                    if (name == null) return;
                    Document doc = Document.loadDoc(name, Wb);

                    Box.Show("Файл '" + Wb.Name + "' распознан как " + name);   //почему-то вылетает!
                }
        */
        /// <summary>
        /// открываем файл Excel по имени name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Excel.Workbook</returns>
        /// <journal>11.12.2013
        /// 7.1.14  - единая точка выхода из метода с finally
        /// 22.12.14 - сообщение о задании Переменной среды
        /// </journal>
        public static Excel.Workbook fileOpen(string name)
        {
            Log.set("fileOpen");
            try {
                if (dirDBs == null)
                {
                    dirDBs = Environment.GetEnvironmentVariable(Decl.DIR_DBS);
                    //if (dirDBs == null) dirDBs = "C:\\DirDbs";  // по умолчанию                                      
                    if (dirDBs == null)
                         Log.FATAL("Не задана переменная среды " + Decl.DIR_DBS +
                             ",\n\t\t\t   показывающая PATH DBs. Для ее определения:" +
                             "\n\n\tКомпьютер-Свойства-Дополонительные параметры системы-Переменные среды");
                }
                if (_app == null) _app = new Excel.Application();
                _app.Visible = true;
                foreach (Excel.Workbook W in _app.Workbooks) if (W.Name == name) return W;
                return _app.Workbooks.Open(dirDBs + "\\" + name);
            } catch (Exception ex) {
                Log.FATAL("не открыт файл" + name + "\nDirDBs=\"" + dirDBs + "\""
                    + "\n сообщение по CATCH '" + ex);
                return null;
            }
            finally { Log.exit(); }
        }
        public static void DisplayAlert(bool val) { _app.DisplayAlerts = val; }
        public static void fileSave(Excel.Workbook Wb) { Wb.Save(); }

        public static void Reset(Excel.Worksheet Sh)
        {
            Sh.UsedRange.Rows.Delete();
        }
        ////public static void Add(Excel.Worksheet Sh, Excel.Range rng)
        ////{
        ////    int iEOL = Sh.UsedRange.Rows.Count;
        ////    rng.Copy(Sh.Cells[iEOL + 1, 1]);
        ////}

#if Test
        public static void TestSheetExists()
        {
            Excel.Workbook Wb = fileOpen(Decl.F_1C);
            Console.WriteLine(" Платежи = ", Convert.ToString(sheetExists(Wb, "Платежи")));
            Console.WriteLine(" Платежи25 = ", Convert.ToString(sheetExists(Wb, "Платежи25")));
        }
#endif
        public static bool sheetExists(Excel.Workbook Wb, string name)
        {
            try { Excel.Worksheet Sh =  Wb.Worksheets[name]; return true; }
            catch { return false; }
        }

        public static long cellColorIndex(Excel.Worksheet Sh, int row, int col, string msg = "")
        {
            Log.set("cellColor");
            try
            {
                Excel.Range cell = Sh.Cells[row, col];
                return cell.Interior.ColorIndex;
            }
            catch
            {
                if (msg == null) return 0;
                if (msg == "") { msg = "Sheet[" + Sh.Name + "].Cell[" + row + "," + col + "]"; }
                Log.FATAL(msg);
                return 0;
            }
            finally { Log.exit(); }
        }

        public static Mtr getRngValue(Excel.Worksheet Sh, int r0, int c0, int r1, int c1, string msg = "")
        {          
            Log.set("getRngValue");
            try
            {
                Excel.Range cell1 = Sh.Cells[r0, c0];
                Excel.Range cell2 = Sh.Cells[r1, c1];
                Excel.Range rng = Sh.Range[cell1, cell2];
                return new Mtr(rng.get_Value());
            }
            catch
            {
                if (msg == "")
                {
                    msg = "Range[ [" + r0 + ", " + c0 + "] , [" + r1 + ", " + c1 + "] ]";
                }
                Log.FATAL(msg);
                return null;
            }
            finally { Log.exit(); }
        }

        public static Mtr getSheetValue(Excel.Worksheet Sh, string msg = "")
        {
            Log.set("getSheetValue");
            try { return new Mtr(Sh.UsedRange.get_Value()); }
            catch
            {
                if (msg == "") msg = "Лист \"" + Sh.Name + "\"" ;
                Log.FATAL(msg);
                return null;
            }
            finally { Log.exit(); }
        }
    }
}