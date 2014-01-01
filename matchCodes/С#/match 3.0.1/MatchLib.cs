/*-----------------------------------------------------------------------
 * MatchLib -- библиотека общих подпрограмм проекта match 3.0
 * 
 *  28.12.2013  П.Храпкин, А.Пасс
 *  
 * - 20.11.13 переписано с VBA на С#
 * - 1.12.13 добавлен метод ToIntList
 * -------------------------------------------
 * EOL(Sh)                  - возвращает число непустых строк в листе Sh
 * ToIntList(s, separator)  - возвращает List<int>, разбирая строку s с разделителями separator
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace match.Lib
{
    /// <summary>
    /// класс MatchLib  -- библиотека общих подпрограмм
    /// </summary>
    public static class MatchLib
    {
        //       public static int EOL2(Excel.Worksheet Sh)
        /*
         * возвращает число непустых строк листа Sh 
         * работает правильно при условии, что в информативной части листа нет
         * последовательностей пустых строк длиной >= MAX_EMPTY_ROWS
         * 
         * 21.11.2013
         * 12.12.2013 - (AP) переписано из EOL Eдля файлов с большим "хвостом" пустых строк
         *      НЕ РАБОТАЕТ!!! Надо
         *      1) попробовать проверять 100 строк после найденного EOL
         *      2) огрваничить число колонок = (Sh.UsedRange.Columns.Count > 20)? 20 : Sh.UsedRange.Columns.Count;
         */
        /*        {
        public static int EOL(Excel.Worksheet Sh)
        {
            int rowCount = Sh.UsedRange.Rows.Count;
            int colCount = Sh.UsedRange.Columns.Count;
            if (rowCount <= 0) return 0;        // особый случай

            int beg = 1, end = rowCount;        // beg всегда показывает на непустую строку, end - на пустую или последнюю 
            int indx = (beg + end) / 2;
            while ((end - beg) > 1)
            {
                if (filledLines(Sh, indx, colCount, rowCount))
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
            return beg;
        }   // end EOL()
        /// <summary>
        /// filledLines(sh, rowBeg, rowCount) возвращает true, если в строке есть непустые ячейки
        /// <param name="sh"></param>
        /// <param name="rowBeg"></param>
        /// <param name="rowCount"></param>
        /// </summary>
        private static bool filledLines(Excel.Worksheet sh, int rowBeg, int cols, int rowCount)
        {
            const int MAX_COL = 20;         // максимальное число просматриваемых колонок листа
            const int MAX_EMPTY_ROWS = 5;   // максимальное число безрезультатных строк вверх, 
            cols = Math.Min(cols, MAX_COL);
            int rows = Math.Min(rowBeg + MAX_EMPTY_ROWS - 1, rowCount);
            for (int row = rowBeg; row <= rows; row++)
            {
                for (int col = 1; col <= cols; col++)
                {
                    if (!isCellEmpty(sh, row, col)) return true;
                }
            }
            return false;
        }
        //    */
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
        /// <summary>
        /// ToStrList(Excel.Range)  - возвращает лист строк, содержащийся в ячеках
        /// </summary>
        /// <param name="rng"></param>
        /// <returns>List<streeng></streeng></returns>
        /// <jornal> 31.12.2013 P.Khrapkin
        /// </jornal>
        public static List<string> ToStrList(Excel.Range rng)
        {
            List<string> strs = new List<string>();
            foreach (Excel.Range cell in rng) strs.Add(cell.Text);
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
            List<string> strs = new List<string>();
            string[] ar = s.Split(separator);
            foreach (var item in ar) strs.Add(item);
            return strs;
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
    }

    /// <summary>
    /// Log & Dump System
    /// </summary>
    public class Log
    {
        private static string _context;
        static Stack<string> _nameStack = new Stack<string>();

        public Log(string msg)
        {
            _context = "";
            foreach (string name in _nameStack)  _context = name + ">" + _context;
            Console.WriteLine(DateTime.Now.TimeOfDay + " " + _context + " " + msg);
        }
        public static void set(string sub)  { _nameStack.Push(sub); }
        public static void exit()           { _nameStack.Pop(); }
        public static void FATAL(string msg)
        {
            new Log("[FATAL] " + msg);
            System.Diagnostics.Debugger.Break();
        }
        public static void Warning(string msg) { new Log("[warning] " + msg); }
    }
}