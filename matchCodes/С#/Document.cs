/*-----------------------------------------------------------------------
 * Document -- ����� ���������� ������� match 3.0
 * 
 *  19.12.2013  �.�������, �.����
 *  
 * - 19.12.13 ���������� � VBA TOCmatch �� �#
 * -------------------------------------------
 * Document(name)       - ����������� ���������� ������ �������� � ������ name
 * loadDoc(name, wb)    - ��������� �������� name ��� ��� ���������� �� ����� wb
 * getDoc(name)         - ���������� �������� � ������ name; ��� ������������� - ��������� ���
 * isDocOpen(name)      - ���������, ��� �������� name ������
 * recognizeDoc(wb)     - ���������� ������ ���� ����� wb �� ������� �������
 * Check(rng,stampList)       - �������� ������� stampList � Range rng
 * 
 * ���������� ����� Stamp ������������ ��� ���������� ������� �������
 * ������ ����� �������� ���������, �� ���� ����������� �����, � ��� ��������� - ��� ���������
 * Stamp(Range rng)     - ��������� rng, ������� �� ������� TOCmatch ����� � List ������� � ���������
 */
using System;
using Box = System.Windows.Forms.MessageBox;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Lib = match.MatchLib;

namespace ExcelAddIn2
{
    /// <summary>
    /// ����� Document �������� ������� ���������� ���� ����������, ��������� ���������� match
    /// </summary>
    public class Document
    {
        private static Dictionary<string, Document> Documents = new Dictionary<string, Document>();   //��������� ����������
 
        private string name;
        private bool isOpen = false;
        private string FileName;
        private string SheetN;
        private string MadeStep;
        private DateTime MadeTime;
        private ulong chkSum;
        private int EOLinTOC;
        private Stamp stamp;        //������ �������� ��������� �� ������� �������� ��� �����
        private DateTime creationDate;  // ���� �������� ���������
        private string Loader;
        private bool isPartialLoadAllowed;
        private string BodyPtrn;
        private string SummPtrn;
        public Excel.Range Body;
        public Excel.Range Summary;

        /// <summary>
        /// F_MATCH = "match.xlsm" - ��� ����� ������ ���������� match
        /// </summary>
        public const string F_MATCH = "match.xlsm";
        /// <summary>
        /// F_1C = "1C.xlsx"    - ���� ������� 1C: ��������, ���������, ������ ��������
        /// </summary>
        public const string F_1C = "1C.xlsx";
        /// <summary>
        /// F_SFDC = "SFDC.xlsx"    - ���� ������� Salesforce.com
        /// </summary>
        public const string F_SFDC = "SFDC.xlsx";
        /// <summary>
        /// F_ADSK = "ADSK.xlsm"    - ���� ������� Autodesk
        /// </summary>
        public const string F_ADSK = "ADSK.xlsm";
        /// <summary>
        /// F_STOCK = "Stock.xlsx"  - ���� ������� �� ������ � �������� �������
        /// </summary>
        public const string F_STOCK = "Stock.xlsx";
        /// <summary>
        /// ��������� ���� ��� ������������� �����������
        /// </summary>
        public const string F_TMP = "W_TMP.xlsm";

        private const string TOC = "TOCmatch";
        private const int TOC_DIRDBS_COL = 10;  //� ������ ������ � ������� TOC_DIRDBS_COL ������� ���� � dirDBs
        private const int TOC_LINE = 4;         //������ ����� TOL_LINE ������� ��� ��������� � ������ ����� ���������.
//        private const string dirDBs = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //��������!!!

        static Document()
        {
            Document doc = null;
            Excel.Workbook db_match = FileOpenEvent.fileOpen(F_MATCH);
            Excel.Worksheet wholeSheet = db_match.Worksheets[TOC];
            Excel.Range tocRng = wholeSheet.Range["5:" + Lib.EOL(wholeSheet)];

            for (int i = 1; i <= tocRng.Rows.Count; i++)
            {
                Excel.Range rw = tocRng.Rows[i];
 
                string docName = rw.Range["B1"].Value2;
                if (!String.IsNullOrEmpty(docName))
                {
                    doc = new Document();
                    doc.MadeTime = DateTime.FromOADate(rw.Range["A1"].Value2);
                    doc.name = docName;
                    string tx = rw.Range["C1"].Value2.ToString();
                    doc.EOLinTOC = String.IsNullOrEmpty(tx) ? 0 : Convert.ToInt32(tx);
                    //                    MyCol        = rw.Range["D1"].Value2;
                    //                    ResLines     = rw.Range["E1"].Value2;
                    doc.MadeStep = rw.Range["F1"].Value2;
                    //                    Period    = rw.Range["G1"].Value2;
                    doc.FileName = rw.Range["H1"].Value2;
                    doc.SheetN = rw.Range["I1"].Value2;
                    Documents.Add(docName, doc);

                    // ��������� Range, ���������� ��� ������ ���������
                    int j;
                    for (j = i + 1; j <= tocRng.Rows.Count
                            && (String.IsNullOrEmpty(tocRng.Range["B" + j].Value2)); j++) ;
                    bool isSF = doc.FileName == F_SFDC;
                    doc.stamp = new Stamp(tocRng.Range["J" + i + ":M" + --j], isSF);

                    doc.creationDate = DateTime.FromOADate(rw.Range["N1"].Value2);

                    doc.BodyPtrn = rw.Range["P1"].Value2;
                    doc.SummPtrn = rw.Range["Q1"].Value2;
                    doc.Loader = rw.Range["T1"].Value2;

                    // ����, ����������� ��������� ���������� ��������� ���� �������� ���������
                    switch (docName)
                    {
                        case "�������":
                        case "��������": doc.isPartialLoadAllowed = true;
                            break;
                        default: doc.isPartialLoadAllowed = false;
                            break;
                    }
                }
            }
            //if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2)
            //{
            //    Box.Show("���� '" + F_MATCH + "' �������� �� ���������� �����!");
            //    // ������������� match -- ����� ������ �����
            //}
        }
        /// <summary>
        /// loadDoc(name, wb)   - �������� ����������� ��������� name �� ����� wb
        /// </summary>
        /// <param name="name"></param>
        /// <param name="wb"></param>
        /// <returns>Document   - ��� ������������� ������ name �� ����� � match � ������� ��� � ������� � wb</returns>
        /// <journal> �� ��������
        /// 15.12.2013 - �������������� � getDoc(name)
        /// </journal>
        public static Document loadDoc(string name, Excel.Workbook wb)
        {
            Document doc = getDoc(name);
            Excel.Workbook wb_sf = FileOpenEvent.fileOpen(doc.FileName);
            Excel.Worksheet Sh = wb_sf.Worksheets[doc.SheetN];
//            Excel.Worksheet Sh = fileOpen(doc.FileName).Worksheets[doc.SheetN];
            if (doc.isPartialLoadAllowed)
            {
// ������ ������������� ��������� ��� ������ ���������� �������� ���������
// ����� ������ ���� ���������, �� ���� ����� ����� ����������� Merge
            }
            wb.Worksheets[1].Name = "TMP";
            wb.Worksheets[1].Move(Sh);
// ����� �� wb ��������� ������ � ������ ����
// � � ����� ��������� Loader
            return doc;
        }
        /// <summary>
        /// getDoc(name)            - ���������� ��������� name. ���� ��� �� ������� - �� �����
        /// </summary>
        /// <param name="name">��� ������������ ���������</param>
        /// <returns>Document</returns>
        /// <journal> �� ��������
        /// 15.12.2013 - ������ �� �����, ������������ Range Body � Summary
        /// </journal>
        public static Document getDoc(string name)
        {
            try
            {
                Document doc = Documents[name];
                if (!doc.isOpen)
                {
                    // �������� ��������� �� �����
                    // ���� ��������� ������ ��������� ��������� � Sh, ��� EOL,
                    // ���������� ��� Range Body � Summary                }
                    //!!!!!                if (!Check(rng, doc.stampList))
// ��� �����������!!                    Excel.Range rng = (doc.FileName == F_SFDC) ? doc.Summary : doc.Body;
                }
                return doc;
            }
            catch
            {
                // ���� ���������, ��� Document name �� ����������
                // � ������, ���� ����������, �� �� ������� ��������� - ������� ������� FATAL_ERR
                return null;
            }
        }
        /// <summary>
        /// isDocOpen(name)     - ���������, ��� �������� name ������ � ��������
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <juornal> 10.12.2013
        /// </juornal> 
        public bool isDocOpen(string name) { return (Documents.ContainsKey(name)); }
        /// <summary>
        /// recognizeDoc(wb)        - ������������� ��������� � �����[1] wb
        /// </summary>
        /// <param name="wb"></param>
        /// <returns>��� ������������� ��������� ��� null, ���� �������� �� ���������</returns>
        /// <journal> 14.12.2013
        /// 16.12.13 (��) ���������� ������������� � ������ if( is_wbSF(wb) )
        /// </journal>
        public static string recognizeDoc(Excel.Workbook wb) {
            Excel.Worksheet wholeSheet = wb.Worksheets[1];
            Excel.Range rng = wholeSheet.Range["1:" + Lib.EOL(wholeSheet).ToString()];

            Stamp stmpSF = Documents["SFDC"].stamp;
            bool is_wbSF = Stamp.Check(rng, stmpSF);
            // ���� ���������� �������� � TOCmatch
            foreach (var doc in Documents)
            {
                if (is_wbSF && (doc.Value.FileName != F_SFDC)) continue;
                if (doc.Value.name == "SFDC" || doc.Value.name == "Process") continue;
                if (Stamp.Check(rng, doc.Value.stamp)) return doc.Value.name;
            }       // ����� ����� �� ����������
            return null;        // ������ �� �����
        }

        /// <summary>
        /// ����� Stamp, ����������� ��� ������ ���������
        /// </summary>    
        private class Stamp
        {
            public List<OneStamp> stamps = new List<OneStamp>();
            /*
             * �����������. 
             *  rng - range, ���������� ������� � J �� � ��� ���� �����, ����������� ��������.
             */
            public Stamp(Excel.Range rng, bool isSF)
            {       // ����
                if ((char)rng.Range["B1"].Value2[0] != 'N')
                {
                    for (int i = 1; i <= rng.Rows.Count; i++) stamps.Add(new OneStamp(rng.Rows[i], isSF));
                }
            }
            /// <summary>
            /// Check(rng, stmp)        - ��������, ��� Range rng ������������� ������� ������� � stmp
            /// </summary>
            /// <param name="rng">Range rng - ����������� ��������</param>
            /// <param name="stmp">Stamp stmp   - ������� �������, ��������������� ������� ���������</param>
            /// <returns>true, ���� ��������� �������� �������������, ����� false</returns>
            /// <journal> 12.12.13
            /// 16.12.13 (��) ������� � ����� Stamp � ���������
            /// </journal>
            public static bool Check(Excel.Range rng, Stamp stmp)
            {
                foreach (OneStamp st in stmp.stamps)
                    if (!OneStamp.Check(rng, st)) return false;
                return true;
            }
        }

        /// <summary>
        /// �����, ����������� ����� ��������� (� ���������� �������, ��������� � ����� ����� TOCmatch)
        /// </summary>
        public class OneStamp
        {
            private string signature;  // ����������� ����� ������ - ���������
            private char typeStamp;   // '=' - ������ ������������ ���������; 'I' - "����� ��������.."
            private List<int[]> stampPosition = new List<int[]>();   // �������������� ������� �������� �������
            private bool _isSF;
 
            /// <summary>
            /// ����������� OneStanp(rng, isSF)
            /// </summary>
            /// <param name="rng">rng - range, ���������� ���� ������ ������ (�.�. ���������)</param>
            /// <param name="isSF">isSF</param>
            /// <example>
            /// �������: {[1, "1, 6"]} --> [1,1] ��� [1,6]
            ///  .. {["4,1", "2,3"]} --> [4,2]/[4,3]/[1,2]/[1,3]
            /// </example>
            /// <journal> 12.12.2013 (AP)
            /// 16.12.13 (��) �������� �������� isSF - ����������� � ��������� ������
            /// </journal>
            public OneStamp(Excel.Range rng, bool isSF)
            {
                signature = rng.Range["A1"].Value2; 
                typeStamp = rng.Range["B1"].Value2[0];
                _isSF = isSF;

                List<int> rw = intListFrCell("C1", rng);
                List<int> col = intListFrCell("D1", rng);
                // ��������� ������������ �������� rw � col
                rw.ForEach(r => col.ForEach(c => stampPosition.Add(new int[] { r, c })));
            }
            /// <summary>
            /// Check(rng, stmp)        - �������� ��������� ������ stmp � rng ��� ��� ���� ���������� �������
            /// </summary>
            /// <param name="rng"></param>
            /// <param name="stmp"></param>
            /// <returns></returns>
            public static bool Check(Excel.Range rng, OneStamp stmp)
            {
                int shiftToEol = (stmp._isSF) ? rng.Rows.Count - 6 : 0;
                string sig = stmp.signature.ToLower();
                foreach (var pos in stmp.stampPosition)
                {
                    var x = rng.Cells[pos[0] + shiftToEol, pos[1]].Value2;
                    if (x == null) continue;
                    string strToCheck = x.ToLower();

                    if (stmp.typeStamp == '=')
                    {
                        if (strToCheck == sig) return true;
                    }
                    else
                    {
                        if (strToCheck.Contains(sig)) return true;
                    }
                }
                return false;
            }

            private List<int> intListFrCell(string coord, Excel.Range rng)
            {
                return Lib.ToIntList(rng.Range[coord].Value2.ToString(), ',');
            }

        }   // ����� ������ OneStamp
    }    // ����� ������ Document
}