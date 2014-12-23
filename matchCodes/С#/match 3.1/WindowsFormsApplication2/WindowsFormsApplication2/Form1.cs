using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using match.Lib;
using Log = match.Lib.Log;
using Docs = match.Document.Document;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            Log.START("match v3.1.0.0");
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        { 
            Docs dicAcc = Docs.getDoc("SF_DicAccSyn");
            int i0 = dicAcc.Body.iEOL();
            Docs synAcc = Docs.NewSheet("DicAccSynonims");

            new Log("Исходный отчет " + dicAcc.name + " из " + i0 + "строк ЗАГРУЖЕН");
 
            for (int i = 2, k = 2; i <= i0 ; i++)
            {
                String CanonicAcc = dicAcc.Body.String(i, 3);
                string[] part = CanonicAcc.Split(new string[] { "<ИЛИ>" }, StringSplitOptions.None);
                foreach (var p in part) 
                {
                    new Log("\t" + i + ": " + k + "\t" + p + "\t" + CanonicAcc);
                    //synAcc.Body.set(k, 1) = p;
                    //synAcc.Body.set(k, 2) = part;
                    k++;
                };
            }
            synAcc.saveDoc();
        }
    }
}
