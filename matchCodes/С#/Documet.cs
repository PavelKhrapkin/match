/*-----------------------------------------------------------------------
 * Document -- ����� ���������� ������� match 3.0
 * 
 *  24.11.2013  �.�������, �.����
 *  
 * - 24.11.13 ���������� � VBA TOCmatch �� �#
 * -------------------------------------------
 * Document(Name)          - ����������� ���������� ������ �������� � ������ Name
 * 
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
    public class Document {

        struct TOCrow
        {
            protected string name;
            protected string FileName;
            protected string SheetN;
            protected string MadeStep;
            protected DateTime MadeTime;
            protected int iTOC;
            protected ulong chkSum;
            protected int EOLinTOC;
            //            protected Stamp;
        }
        public const string F_MATCH = "match.xlsm";
            private const string TOC = "TOCmatch";
                private const int TOC_DIRDBS_COL = 10;  //� ������ ������ � ������� TOC_DIRDBS_COL ������� ���� � dirDBs
                private const int TOC_LINE = 4;         //������ ����� TOL_LINE ������� ��� ��������� � ������ ����� ���������.
                private static int EOL_toc;             //����� ����� � ���. ������������ ��� ������������� ��� � �������� � TOC_LINE
                private int iTOC;                       //����� ������ � ��� - ������� ��������� �� ����� name
        private const string dirDBs  = "C:\\Users\\Pavel_Khrapkin\\Documents\\Pavel\\match\\matchDBs\\";    //��������!!!

        private const string F_1C = "1C.xlsx";
        private const string F_SFDC = "SFDC.xlsx";
        private const string F_ADSK = "ADSK.xlsm";
        private const string F_STOCK = "Stock.xlsx";
        private const string F_TMP   = "W_TMP.xlsm";

        public Document[] OpenDocs;                     //��������� ���������� ��� �������� � match
/*
        Stamp stamp;
 
        Body body;
        Header hdr;
        Summary smr;
        LeftCols lft;
        RightCols rght;
 */
        bool isTOCinitiated = false;
/*
 * ����������� ���������
 */
        public Document(string name) {
/* �������, ���� �����, ���������� ��������� ��� - ���������� ��� ������� ���������� */
//            if (!isTOCinitiated) {
            if (OpenDocs[1] == null) {
                TOCrow toc = new TOCrow();
                toc.name = TOC; toc.SheetN = TOC; toc.FileName = F_MATCH;
                Excel.Workbook db_match = fileOpen(F_MATCH);
                toc.EOLinTOC = Lib.EOL(db_match.Worksheets[TOC]);

                if (dirDBs != (string)db_match.Worksheets[TOC].cells[1, TOC_DIRDBS_COL].Value2) {
                    Box.Show("���� '" + F_MATCH + "' �������� �� ���������� �����!");
                }

//                WrTOC(TOC);    /* WrTOC - �����, ������������ ������ �� ���������� � ���� TOCmatch - ������� ����� */
            }
/* ������� �������� name � ��� �������� ��� ��������� �� ���� ����� */
            for (iTOC = TOC_LINE; iTOC <= EOL_toc; iTOC++) {
 //               if TOC.
            }
        }
        public bool isDocOpen(string name) {
            foreach (Document Doc in OpenDocs) {
                if (Doc.Name == name) return true;
            }
            return false;
        }

        public string Name {
            get { return TOCrow.name; }
            set { TOCrow.name = value; }
        }

        public bool CheckStamp() {
            return false;
        }

        private static Excel.Workbook  fileOpen(string name) {

            Excel.Application app = new Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb;
            try {
                wb = app.Workbooks.Open(dirDBs + name);
                return wb;
            } catch {
                return null;
            }
       }
    }
    class Stamp {
    }
    class Header {
    }
    class Body {
    }
    class Summary {
    }
    class LeftCols {
    }
    class RightCols {
    }
}