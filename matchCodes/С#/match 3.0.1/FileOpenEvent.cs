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
    class FileOpenEvent
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

                    Box.Show("���� '" + Wb.Name + "' ��������� ��� " + name);   //������-�� ��������!
                }
        */
        /// <summary>
        /// ��������� ���� Excel �� ����� name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Excel.Workbook</returns>
        /// <journal>11.12.2013
        /// 7.1.14  - ������ ����� ������ �� ������ � finally
        /// </journal>
        public static Excel.Workbook fileOpen(string name)
        {
            Log.set("fileOpen");
            try {
                if (dirDBs == null)
                {
                    dirDBs = Environment.GetEnvironmentVariable(Decl.DIR_DBS);
                    //if (dirDBs == null) dirDBs = "C:\\DirDbs";  // �� ���������                                      
                    if (dirDBs == null)
                    {
                        // �� ���������                                      
#if msgBox                        
                        // Configure open file dialog box
                        Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                        dlg.FileName = "DirDbs"; // Default file name
                        dlg.DefaultExt = ""; // Default file extension
//                        dlg.Filter = "Text documents (.txt)|*.txt"; // Filter files by extension 

                        // Show open file dialog box
                        Nullable<bool> result = dlg.ShowDialog();

                        // Process open file dialog box results 
                        if (result == true)
#else
                        if (dirDBs == null) dirDBs = "C:\\DBs";
#endif

                    }
                }
                if (_app == null) _app = new Excel.Application();
                _app.Visible = true;
                foreach (Excel.Workbook W in _app.Workbooks) if (W.Name == name) return W;
                return _app.Workbooks.Open(dirDBs + "\\" + name);
            } catch (Exception ex) {
                Log.FATAL("�� ������ ����" + name + "\nDirDBs=\"" + dirDBs + "\""
                    + "\n ��������� �� CATCH '" + ex);
                return null;
            }
            finally { Log.exit(); }
        }
        public static void DisplayAlert(bool val) { _app.DisplayAlerts = val; }
        public static void fileSave(Excel.Workbook Wb) { Wb.Save(); }

#if Test
        public static void TestSheetExists()
        {
            Excel.Workbook Wb = fileOpen(Decl.F_1C);
            Console.WriteLine(" ������� = ", Convert.ToString(sheetExists(Wb, "�������")));
            Console.WriteLine(" �������25 = ", Convert.ToString(sheetExists(Wb, "�������25")));
        }
#endif
        public static bool sheetExists(Excel.Workbook Wb, string name)
        {
            try { Excel.Worksheet Sh =  Wb.Worksheets[name]; return true; }
            catch { return false; }
        }

        public static Mtr getRngValue(Excel.Worksheet Sh, int r0, int c0, int r1, int c1, string msg = "")
        {
            Log.set("ToMtr");
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
    }
}