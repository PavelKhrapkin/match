/*-----------------------------------------------------------------------
 * MatchLib -- библиотека общих подпрограмм проекта match 3.0
 * 
 *  23.01.14 П.Храпкин, А.Пасс
 *  
 * - 20.11.13 переписано с VBA на С#
 * - 1.12.13 добавлен метод ToIntList
 * - 1.1.14 добавлен ToStrList с перегрузками и умолчаниями
 * - 02.01.14 двоичный поиск EOL
 *            добавлен метод getDateTime(dynamic inp)
 * - 11.01.14 EOC - определение числа колонок листа -- ПХ
 * - 23.01.14 ToStrList - используем ADO.Net
 * -------------------------------------------
 * EOL(Sh)                  - возвращает число непустых строк в листе Sh
 * ToIntList(s, separator)  - возвращает List<int>, разбирая строку s с разделителями separator
 * ToStrList(s,[separator]) - возвращает List<string> из Range или из строки s с разделителем
 */

#define binarySearchEOL
//#define TimeTest

using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using FileOp = match.MyFile.FileOpenEvent;
using Decl = match.Declaration.Declaration;
using Mtr = match.Matrix.Matr;

namespace match.Lib
{
    /// <summary>
    /// класс MatchLib  -- библиотека общих подпрограмм
    /// </summary>
    public static class MatchLib
    {

#if binarySearchEOL
        /// <summary>
        ///  EOL(Worksheet Sh)   - возвращает число непустых строк листа Sh
        /// </summary>
        /// <param name="Sh"></param>
        /// <returns></returns>
        /// <journal>21.11.2013
        /// 13.12.13 - ограничение количества просматриваемых колонок до 20
        /// 25.12.13 - модификация для усечения файлов с очень длинным пустым хвостом
        /// 28.12.13 - bug fix - при обнаружении непустой ячейки break в цикле по строке
        /// 02.01.14 - применен алгоритм двоичного поиска пустой строки
        ///         с поиском MAX_EMPTY_ROWS пустых строк в конце
        /// </journal>
        const int MAX_COL = 20;         // максимальное число просматриваемых колонок листа
        const int MAX_EMPTY_ROWS = 10;  // максимальное допустимое число пустых строк не в конце листа, 
        public static int EOL(Excel.Worksheet Sh)
        {
            int rowCount = Sh.UsedRange.Rows.Count;
            int colCount = Sh.UsedRange.Columns.Count;
            if (rowCount <= 0) return 0;        // особый случай

            int beg = 1, end = rowCount + 1;    // beg всегда показывает на непустую строку, end - на пустую или последнюю 
            int indx = (beg + end) / 2;
            while ((end - beg) > 1)
            {
                if (indx < rowCount && isLineFilled(Sh, indx, colCount))
                {
                    beg = indx;
                    indx += (end - indx) / 2;
                }
                else
                {
                    end = indx;
                    indx -= (indx - beg) / 2;
                }
            }
            //
            // проверяем, что за найденной строкой есть MAX_EMPTY_ROWS пустых (или конец листа)
            // найдя непустую строку - запоминаем ее как найденную и сдвигаемся
            //
            //            bool notEmpty = true;
            for (end = beg + MAX_EMPTY_ROWS; indx < end; indx++)
            {
                if (indx > rowCount)
                {
                    return beg;
                }
                if (isLineFilled(Sh, indx, colCount))
                {
                    beg = indx;
                    end = beg + MAX_EMPTY_ROWS;
                }
            }
            return beg;
        }   // end EOL()
        /// <summary>
        ///  EOC(Worksheet Sh)   - возвращает число непустых колонок листа Sh
        /// </summary>
        /// <param name="Sh"></param>
        /// <returns></returns>
        /// <journal>11.1.2014 переписано из EOL</journal>
        public static int EOC(Excel.Worksheet Sh)
        {
            object[,] bdy = (object[,])Sh.UsedRange.get_Value();
            return bdy.GetLength(1);
        }
        /*
         * возвращает true, если в строке есть непустые ячейки
         */
        //        private static bool isLineFilled(Excel.Worksheet sh, int row)
        private static bool isLineFilled(Excel.Worksheet sh, int row, int cols)
        {
            cols = Math.Min(cols, MAX_COL);
            for (int col = 1; col <= cols; col++)
            {
                if (!isCellEmpty(sh, row, col)) return true;
            }
            return false;
        }
#else
        /// <summary>
        ///  EOL(Worksheet Sh)   - возвращает число непустых строк листа Sh
        /// </summary>
        /// <param name="Sh"></param>
        /// <returns></returns>
        /// <journal>21.11.2013
        /// 13.12.13 - ограничение количества просматриваемых колонок до 20
        /// 25.12.13 - модификация для усечения файлов с очень длинным пустым хвостом
        /// 28.12.13 - bug fix - при обнаружении непустой ячейки break в цикле по строке
        /// </journal>
        public static int EOL(Excel.Worksheet Sh)
        {
            const int MAX_COL = 20;         // максимальное число просматриваемых колонок листа
            const int MAX_EMPTY_ROWS = 5;   // максимальное число безрезультатных строк вверх, 
            // .. после которых увеличиваем шаг
            int step = 1;
            int maxCol = Math.Min(Sh.UsedRange.Columns.Count, MAX_COL);

            for (int i = Sh.UsedRange.Rows.Count, row_count = 1; i > 0; i--, row_count++)
            {
                if (row_count > MAX_EMPTY_ROWS)
                {
                    row_count = 1;
                    step *= 2;
                    i -= step;
                    if (i <= 0) i = 1;

                }
                for (int col = 1; col < maxCol; col++)
                {
                    if (Sh.Cells[i, col].Value2 != null)
                    {
                        if (step == 1) return i;
                        i += step;
                        step = 1;
                        row_count = 1;
                        break;
                    }
                }
            }
            return 0;
        }
#endif
        /// <summary>
        /// ToStrList(Excel.Range)  - возвращает лист строк, содержащийся в ячейках
        /// </summary>
        /// <param name="rng"></param>
        /// <returns>List<streeng></streeng></returns>
        /// <jornal> 1.1.2014 P.Khrapkin
        /// </jornal>
        public static List<string> ToStrList(Excel.Range rng)
        {
            List<string> strs = new List<string>();
            foreach (Excel.Range cell in rng) strs.Add(cell.Text);
            return strs;
        }
        /// <summary>
        /// ToStrList(DataRow, int[] indx)  - возвращает лист строк, содержащийся в ячейках ряда,
        ///                                   указанных в массиве indx - в списке номеров колонок
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="indx"></param>
        /// <returns></returns>
        public static List<string> ToStrList(DataRow rw, int[] indx)
        {
            List<string> strs = new List<string>();
            foreach (int i in indx) strs.Add(rw[i] as string);
            return strs;
        }
        /// <summary>
        /// overloaded ToStrList(string, [separator = ','])
        /// </summary>
        /// <param name="s"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        public static List<string> ToStrList(string s, char separator = ',')
        {
            if (s == null) return null;
            List<string> strs = new List<string>();
            string[] ar = s.Split(separator);
            foreach (var item in ar) strs.Add(item);
            return strs;
        }
        /// <summary>
        /// если в rng null, пусто или ошибка - возвращаем 0, а иначе целое число
        /// </summary>
        /// <param name="rng"></param>
        /// <returns></returns>
        public static int RngToInt(Excel.Range rng)
        {
            int v = 0;
            try
            {
                string str = rng.Text;
                int.TryParse(str, out v);
            }
            catch { v = 0; }
            return v;
        }
        /// <summary>
        /// ToIntList(s, separator) - разбирает строку s с разделительным символом separatop;
        ///                           возвращает List int найденных целых чисел
        /// </summary>
        /// <param name="s"></param>
        /// <param name="separator"></param>
        /// <returns></returns>
        /// <journal> 12.12.13 A.Pass
        /// </journal>
        public static List<int> ToIntList(string s, char separator)
        {
            string[] ar = s.Split(separator);
            List<int> ints = new List<int>();
            foreach (var item in ar)
            {
                int v;
                if (int.TryParse(item, out v)) ints.Add(v);
            }
            return ints;
        }
        public static int ToInt(string s, string msg = "не разобрана строка")
        {
            int v;
            if (int.TryParse(s, out v)) return v;
            Log.Warning(msg + " \"" + s + "\"");
            return -1;
        }
        /// <summary>
        /// isCellEmpty(sh,row,col)     - возвращает true, если ячейка листа sh[rw,col] пуста или строка с пробелами
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        /// <journal> 13.12.13 A.Pass
        /// </journal>
        public static bool isCellEmpty(Excel.Worksheet sh, int row, int col)
        {
            var value = sh.UsedRange.Cells[row, col].Value2;
            return (value == null || value.ToString().Trim() == "");
        }

#if TimeTest
        public static int TestTime()
        {
            DateTime _t = getDateTime(null);
            Console.WriteLine("{0} -- отладка getDateTime(null)", _t);

            _t = getDateTime(DateTime.Now);
            Console.WriteLine("{0} -- отладка getDateTime(Now)", _t);

            _t = getDateTime("1/1/2014 0:0:0 PM");
            Console.WriteLine("{0} -- отладка getDateTime(1/1/2014 0:0:0 PM)", _t);

            _t = getDateTime("1/12/2013 0:0:0 PM");
            Console.WriteLine("{0} -- отладка getDateTime(1/12/2013 0:0:0 PM)", _t);

            _t = getDateTime("31/12/2013 0:0:0 PM");
            Console.WriteLine("{0} -- отладка getDateTime(31/12/2013 0:0:0 PM)", _t);

            _t = Lib.MatchLib.getDateTime("12/31/2013 0:0:0 PM");
            Console.WriteLine("{0} -- отладка getDateTime(12/31/2013 0:0:0 PM)", _t);

            _t = DateTime.Now;
            Console.WriteLine("{0} -- отладка getDateTime(Now)", _t.ToShortDateString());
            return 0;
        }
#endif
        /// <summary>
        /// getDateTime(dynamic inp)     - возвращает DateTime для любого значения (из ячейки Excel)
        /// </summary>
        /// <param name="inp"></param>
        /// <returns>DateTime</returns>
        /// <journal> 02.01.14 A.Pass
        /// </journal>
        public static DateTime getDateTime(dynamic inp)
        {

            if (inp == null) return new DateTime(0);
            if (inp.GetType() == typeof(string))
            {
                DateTime ret;
                try
                {
                    Double x = Double.Parse(inp);
                    ret = DateTime.FromOADate(x);
                    return ret;
                }
                catch
                {
                    string[] pDate = inp.Split(' ');
                    string[] day = pDate[0].Split('/');
                    string[] time = pDate[1].Split(':');
                    int hour = Convert.ToInt32(time[0]);    //час
                    if (pDate[2] == "PM" && hour != 12) hour += 12; // 12 PM - полдень
                    else if (hour == 12) hour = 0;          // 12 АМ - полночь
                    try
                    {
                        return new DateTime(
                             Convert.ToInt32(day[2])        //год
                            , Convert.ToInt32(day[1])       //мес
                            , Convert.ToInt32(day[0])       //день
                            , hour
                            , Convert.ToInt32(time[1])      //мин
                            , Convert.ToInt32(time[2]));    //сек
                    }
                    catch
                    {
                        return new DateTime(
                             Convert.ToInt32(day[2])        //год
                            , Convert.ToInt32(day[0])       //мес
                            , Convert.ToInt32(day[1])       //день
                            , hour
                            , Convert.ToInt32(time[1])      //мин
                            , Convert.ToInt32(time[2]));    //сек
                    }
                }
            }
            else if (inp.GetType() == typeof(Double))
            {
                return DateTime.FromOADate(inp);
            }
            return inp;
        }
    }   // конец класса MatchLib

    /// <summary>
    /// Log & Dump System
    /// </summary>
    /// <journal> 30.12.2013 P.Khrapkin
    /// 20.1.14 - рабта с Log в виде List
    /// </journal>
    public class Log
    {
        private static string _context;
        static Stack<string> _nameStack = new Stack<string>();

        struct msgLog
        {
            DateTime tm;
            string msg;

            public msgLog(string msg)
            {
                tm = DateTime.Now;
                this.msg = msg;
            }
        }
        static List <msgLog> msgs= new List<msgLog>();

//        private static bool _initialized = false;   // чтобы избежать зацикливания в fileOpen
        static Excel.Worksheet _logSheet;
        // статический конструктор класса Log открывает лист
        static Log()
        {      // открываем Sheet Log
            _logSheet = FileOp.fileOpen(Decl.F_MATCH).Worksheets[Decl.LOG];
//            _initialized = true;
        }

        public Log(string msg)
        {
            _context = "";
            foreach (string name in _nameStack) _context = name + ">" + _context;
            Console.WriteLine(DateTime.Now.ToLongTimeString() + " " + _context + " " + msg);
            msgs.Add(new msgLog(msg));
        }

        public static void set(string sub) { _nameStack.Push(sub); }
        public static void exit() { _nameStack.Pop(); }
        public static void FATAL(string msg)
        {
            new Log("[FATAL] " + msg);
            System.Diagnostics.Debugger.Break();
        }
        public static void Warning(string msg) { new Log("[warning] " + msg); }
        public static void START(string msg)
        {
//            Console.WriteLine(DateTime.Now.ToShortDateString() + " ---------< " + msg + " >---------"); 
            new Log("---------< " + msg + " >---------");
        }
        public static void Save()
        {
            object s = msgs.ToArray();
//            Excel.Range rng = s;
        }
    }
}