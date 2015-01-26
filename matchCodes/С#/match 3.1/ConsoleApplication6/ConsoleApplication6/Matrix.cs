/*-----------------------------------------------------------------------
 * Matrix -- класс массивов и методов для работы с матрицами в общем виде
 *  22.01.2015  П.Храпкин, А.Пасс
 *  
 * -------------------------------------------
 * Matr(object[,] obj    - КОНСТРУКТОР порождает внутренний массив из obj[,] или Data Table
 *     Matr[,] Indexer   - индексация матрицы Matr
 *     get(int i, int j) - возвращает элемент [i,j]
 *     set(int i, int j) - записывает элемент [i,j] 
 * 
 * String(i,j)  - возвращает string в ячейке Matr[i,j]
 * Int(i,j,msg) - возвращает int в ячейке Matr[i,j] или выводит сообщение об ощибке msg
 * Float(i,j,msg) - возвращает Float в ячейке Matr[i,j] или выводит сообщение об ощибке msg
 * 
 * LBount(), UBound(), LBound(1), UBound(1) - верхние и нижние значения инднксов для строк и столбцов
 * iEOL(), iEOC()   - количество строк и столбцов в матрице Matr
 * 
 * Compare(Mart A, Matr B)  - возвращает true, если все элементы А==B
 * DataTable DaTab()        - копирует данные из (matr)this в (Data Table)
 ! AddRow(object[] Line)    - добавляет строку Line после iEOL()
 */
using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Log = match.Lib.Log;

namespace match.Matrix
{
    public class Matr : Object
    {
        private const int MATR_PAGE = 100;
        private object[,] _matr = new object[MATR_PAGE, MATR_PAGE];
        
        public Matr(object[,] obj)
        {
            _matr = obj;
        }
/* --- пока не нужно
        public Matr(DataTable dt)
        {
            try
            {
                int rw = 0;
                foreach (DataRow row in dt.Rows)
                {
                    for (int col = 0; col <= dt.Columns.Count; col++)
                        _matr[rw, col] = (object)row[col];
                    ++rw;
                }
            }
            catch (Exception ex)
            {
                string mes = ex.Message;
            }
        }

        public List<object> getRow(int iRow)
        {
            List<object> _row = new List<object>(); 
            for (int i = 0; i < iEOC(); i++) _row.Add(_matr[iRow, i]);
            return _row;
        }
        public List<object> getCol(int col)
        {
            List<object> _col = new List<object>(); 
            for (int i = 0; i < iEOC(); i++) _row.Add(_matr[iRow, i]);
            return _row;
        }
*/
        public object this[int i, int j]
        {
            get { return get(i, j); }
            set
            {
                try { this._matr[i, j] = value; }
                catch { Log.FATAL("ошибка при обращении к Matr[" + i + "," + j + "]"); }
            }
        }
        public object get(int i, int j)
        {
            object v = null;
            try { v = _matr[i, j]; }
            catch { Log.FATAL("ошибка при обращении к Matr[" + i + "," + j + "]"); }
            return v;
        }
#if пока_не_нужно
        //public object[,] set(int i, int j)
        //{
        //    try { _matr[i, j] = this._matr[i,j]; }
        //    catch { Log.FATAL("!"); }
        //    return _matr;
        //}
        /////// <summary>
        /////// setRow(int i, object[] obj) - записывает ряд значений в Matr
        /////// </summary>
        /////// <param name="i">    int i   - номер рекорда матрицы</param>
        /////// <param name="object[] obj"> object[] obj - массив объектов, записываемых в матрицу</param>
        /////// <journal>28.12.2014
        /////// </journal>
        ////public void setRow(int i, object[] obj)
        ////{
        ////    try { for (int col = 1; col <= this.iEOC(); col++) { _matr[i, col] = obj[col]; } }
        ////    catch { Log.FATAL("! строка " + i); }
        ////}
        /////// <summary>
        /////// int AddRow(object[] obj) - добавляет ряд значений в Matr
        /////// </summary>
        /////// <param name="object[] obj">object[] obj - массив объектов, записываемых в матрицу</param>
        /////// <journal>28.12.2014
        /////// </journal>
        ////public void AddRow(object[] obj)
        ////{
        ////    try
        ////    {
        ////        if (_EOL++ > MATR_PAGE)
        ////        {
        ////            Log.FATAL("!!!! не написано");
        ////        }
        ////        setRow(_EOL, obj);
        ////    }
        ////    catch { Log.FATAL("! EOL=" + _EOL); }
        ////}
#endif
        public string String(int i, int j)
        {
            var v = get(i, j);
            return (v == null) ? "" : v.ToString();
        }
        public int Int(int i, int j, string msg = "wrong int")
        {
            object v = get(i, j);
            if (v == null) return 0;
            if (v.GetType() == typeof(int)) { return (int)v; }
            try
            {
                int value;
                string val = v.ToString();
                if (int.TryParse(val, out value)) return value;
                Log.FATAL(msg);
            }
            catch { Log.FATAL(msg); }
            return 0;
        }
        public float Float(int i, int j, string msg = "wrong Float")
        {
            object v = get(i, j);
            if (v == null) return 0;
            if (v.GetType() == typeof(float)) { return (float)v; }
            try
            {
                int value;
                string val = v.ToString();
                if (int.TryParse(val, out value)) return value;
                Log.FATAL(msg);
            }
            catch { Log.FATAL(msg); }
            return 0;
        }
        public int LBound() { return _matr.GetLowerBound(0); }
        public int LBound(int dm) { return _matr.GetLowerBound(dm); } 
        public int UBound() { return _matr.GetUpperBound(0); }
        public int UBound(int dm) { return _matr.GetUpperBound(dm); }
        public int iEOL() { return _matr.GetLength(0); }
        public int iEOC() { return _matr.GetLength(1); }

        /// <summary>
        /// Compare(Mart A, Matr B) - возвращает true, если все элементы А==B
        /// </summary>
        /// <param name="A"></param>
        /// <param name="B"></param>
        /// <returns></returns>
        /// <journal>18.1.2015</journal>
        public bool Compare(Matr A, Matr B)
        {
            if (A.iEOL() != B.iEOL() || A.iEOC() != B.iEOC() ) return false;
            int i0 = A.LBound(0), i1 = A.UBound(0), j0 = A.LBound(1), j1 = A.UBound(1);
            for (int i = i0; i <= i1; ++i)
                for (int j = j0; j <= j1; ++j)
                    if (A._matr[i, j] != B._matr[i, j]) return false;
            return true;
        }
        public bool Compare(Matr B) { return B.Compare(this, B); }

        /// <summary>
        /// DataTable DaTab()   копирует данные из (matr)this в (Data Table)
        /// </summary>
        /// <returns></returns>
        /// <journal>2014
        /// 15.01.15 buf fix: все индексы dt и параметры get ведем в диапазоне 0..iEO, а не 1..iEO 
        /// 20.01.15 попытался опять использовать Data Table с Datab() и TabDa()
        /// 22.01.15 bug fix: zero based Data Table index
        /// </journal>
        public DataTable DaTab()
        {
            DataTable _dt = new DataTable();
            int minRow = LBound(0), minCol = LBound(1), maxRow = UBound(0), maxCol = UBound(1);

            for (int col = minCol; col <= maxCol; col++) _dt.Columns.Add();
            for (int row = minRow; row <= maxRow; row++) _dt.Rows.Add();

            for (int row = minRow, rdt = 0; row <= maxRow; row++, rdt++)
                for (int col = minCol, cdt = 0; col <= maxCol; col++)
                    _dt.Rows[rdt][cdt++] = get(row, col);
            return _dt;
        }
//!!! не написано
        private static Dictionary<string, object> docToDic(Document.Document doc)
        {
            Object[,] _ar = doc.Body._matr;

            Dictionary<string, object> _dic = new Dictionary<string, object>();
            //_dic = doc.Body._matr.ToDictionary(item => item.Key,
            //                        item => item.Value);

//            for (int rw
            return _dic;
        }

        /// <summary>
        /// AddRow() - добавляет одну строку в конце матрицы
        /// </summary>
        /// <journal> 17.01.15
        /// 19.01.2015 - смотрим на LBound() и UBound
        /// </journal>
        public Matr AddRow()
        {
            int minRow = LBound(0);
            int minCol = LBound(1);
            int maxRow = UBound(0);
            int maxCol = UBound(1);

            var newMatr = new object[maxRow + 1, maxCol];

            for (int i = minRow; i <= maxRow; i++)
                for (int j = minCol; j <= maxCol; j++)
                    newMatr[i,j] = get(i, j);
            for (int j = minCol; j <= maxCol; j++) newMatr[++maxRow, j] = null;
            _matr = newMatr;
            return this;
        }
        public Matr AddRow(object[] Line)
        {
            AddRow();
            int i0 = Line.GetLowerBound(0), i1 = Line.GetUpperBound(0);
            int cols = Math.Min(iEOC(), Line.Length);
            int rw = iEOL() + 1;
            for (int i = 0; i < cols; i++) _matr[rw, i] = Line[i];
            return this;
        }

        public int rw { get; set; }
    } // конец класса Matr
}
