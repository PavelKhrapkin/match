/*-----------------------------------------------------------------------
 * FileOpenEvent - ������ � �������� �������� ������� match 3.1
 * 
 *  24.01.2015  �.�������, �.����
 *  
 * -------------------------------------------
 * WrCSV(name)          - ���������� CSV ���� ��� ��� ����������� ���� � SalesForce
 * WrReport(name,dt)    - ���������� ��������� ���� name � ������� �������
 */
using System;
using System.IO;
using System.Data;
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

        /// <summary>
        /// WrCSV(name) - ���������� CSV ���� ��� ��� ����������� ���� � SalesForce
        /// </summary>
        /// <param name="name">string name  - ��� ����� ��� ������</param>
        /// <journal>23/1/2015</journal>
        public static void WrCSV(string name, DataTable dt)
        {
            string pathCSV = @"C:/SFconstr/";    // �������, ���� ��������� CSV �����
            FileInfo file = new FileInfo(pathCSV + name + @".csv");
            StreamWriter fs = file.CreateText();

            foreach (DataRow row in dt.Rows)
            {
                string str = "";
                foreach (DataColumn x in dt.Columns)
                {
                    if (str != "") str += ',';
                    str += '"';
                    str += row[x].ToString();
                    str += '"';
                }
                fs.WriteLine(str); 
            }
            fs.Close();
        }
        /// <summary>
        /// WrReport(name,dt)   - ���������� ��������� ���� name � ������� �������
        /// </summary>
        /// <param name="name">string name - ��� ����� - ������ *.txt</param>
        /// <param name="dt">DataTable dt - ������� � ������� ��� ������</param>
        /// <journal>23.01.2015</journal>
        public static void WrReport(string name, DataTable dt)
        {
            setDirDBs();
            string fileName = dirDBs + @"\Reports\" + name + @".txt";
            using (StreamWriter fs = new StreamWriter( fileName, true, System.Text.Encoding.Default))
            {
                fs.WriteLine("--- " + DateTime.Now.ToLongTimeString() + " " + name + " ------------------");
                foreach (DataRow row in dt.Rows)
                {
                    string str = "";
                    foreach (DataColumn x in dt.Columns)
                    {
                        if (str != "") str += '\t';
                        str += row[x].ToString();
                    }
                    fs.WriteLine(str);
                }
                fs.Close();
            }
        }
        private static void setDirDBs()
        {
            if (dirDBs == null) dirDBs = Environment.GetEnvironmentVariable(Decl.DIR_DBS);
            if (dirDBs == null)
                Console.WriteLine("�� ������ ���������� ����� " + Decl.DIR_DBS +
                    ",\n\t\t\t   ������������ PATH DBs. ��� �� �����������:" +
                    "\n\n\t���������-��������-��������������� ��������� �������-���������� �����");
        }

        /// <summary>
        /// ��������� ���� Excel �� ����� name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>Excel.Workbook</returns>
        /// <journal>11.12.2013
        /// 7.1.14  - ������ ����� ������ �� ������ � finally
        /// 22.12.14 - ��������� � ������� ���������� �����
        /// 24.01.15 - setDirDBs �������� � ��������� ������������
        /// </journal>
        public static Excel.Workbook fileOpen(string name)
        {
            Log.set("fileOpen");
            try {
                setDirDBs();
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
            Console.WriteLine(" ������� = ", Convert.ToString(sheetExists(Wb, "�������")));
            Console.WriteLine(" �������25 = ", Convert.ToString(sheetExists(Wb, "�������25")));
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
                if (msg == "") msg = "���� \"" + Sh.Name + "\"" ;
                Log.FATAL(msg);
                return null;
            }
            finally { Log.exit(); }
        }
    }
}