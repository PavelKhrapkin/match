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
         *
         * 21.11.2013
         * 12.12.2013 - (AP) переписано из EOL Eдля файлов с большим "хвостом" пустых строк
         *      НЕ РАБОТАЕТ!!! Надо
         *      1) попробовать проверять 100 строк после найденного EOL
         *      2) огрваничить число колонок = (Sh.UsedRange.Columns.Count > 20)? 20 : Sh.UsedRange.Columns.Count;
         */
        /*        {
                    if (Sh.UsedRange.Rows.Count <= 0) return 0; // особый случай

                    int beg = 1, end = Sh.UsedRange.Rows.Count + 1;
                    // beg всегда показывает на непустую строку, end - на пустую или EOL 
                    int indx = (beg + end) / 2;
                    while ((end - beg) > 1) {
                        bool found = false;
                        for (int col = 1; col <= Sh.UsedRange.Columns.Count; col++) {
                            var x = Sh.UsedRange.Cells[indx, col].Value2;
                            if (Sh.UsedRange.Cells[indx, col].Value2 != null) {
                                found = true;
                                break;
                            }
                        }
                        if (found) {
                            beg = indx;
                            indx += (end - indx) / 2;
                        } else {
                            end = indx;
                            indx -= (indx - beg) / 2;
                        }
                    }
                    return beg;
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