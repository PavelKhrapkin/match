using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using match.Document;
using match.Lib;
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
                    {
                        // по умолчанию                                      
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
                Log.FATAL("не открыт файл" + name + "\nDirDBs=\"" + dirDBs + "\""
                    + "\n сообщение по CATCH '" + ex);
                return null;
            }
            finally { Log.exit(); }
        }
        public static void fileSave(Excel.Workbook Wb) { Wb.Save(); }
    }
}