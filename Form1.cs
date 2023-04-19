using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using System.IO;
using System.Data.Odbc;

namespace ExportFK
{
    public partial class ExportDoc : Form
    {
        //string[] strNdoc;
        List<Int64> strNdoc = new List<Int64>();
        public ExportDoc()
        {

            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string strdt = dateTimePicker1.ToString();
            //           strdt = strdt.Substring(44, 10);
            int k = strdt.LastIndexOf("202")-6;
            strdt = strdt.Substring(k,10);

            string SqlString = @"select docs.docno, docs.docdate,partners.partner_name, docs.doc_id
                          from docs, partners
                          where docs.doctype = 16 and
                                docs.docdate = '" + strdt + @"' and
                                docs.partner_id = partners.partner_id";

            DataSet dsDocs = GetSelect(SqlString);
            listBox1.Items.Clear();
            
            int i = 0;
            foreach (DataRow row in dsDocs.Tables[0].Rows)
            {

                listBox1.Items.Add(FormatStrDoc(row.ItemArray[0].ToString(), 10) + FormatStrDoc(row.ItemArray[1].ToString().Substring(0,10),15) +
                                       FormatStrDoc(row.ItemArray[2].ToString(),55)); 

                //              strNdoc[i].Insert(i, row.ItemArray[3].ToString());
                strNdoc.Add((Int64)row.ItemArray[3]);
              i = i + 1;
            }


        }


        public static DataSet GetSelect(string strSQL)
        {
            string strCon = GetConnectionString();

            FbConnectionStringBuilder fb_con = new FbConnectionStringBuilder();
            fb_con.Charset = "WIN1251";
            fb_con.UserID = "SYSDBA";
            fb_con.Password = "masterkey";
            fb_con.Database = @strCon;
            fb_con.ServerType = 0;
            //создаем подключение
            var fb = new FbConnection(fb_con.ToString());
            fb.Open();
            FbDataAdapter myAdapter = new FbDataAdapter(strSQL, fb);
            DataSet ds = new DataSet();
            myAdapter.Fill(ds);
            fb.Close(); // по правилам хорошего тона ....
            return (ds);
        }

        public static string FormatStrDoc(string instr,int lenField)
        {
            string st_out = instr;
            if (instr.Length>lenField)
            {
                st_out = instr.Substring(0, lenField);
                return (st_out);
            }
            do
            {
                st_out = st_out + " ";
            }
            while (lenField != st_out.Length);
            return (st_out);
        }


        public static string GetConnectionString()
        {
            string path1 = "       ";
            string strConnect = "   ";
            if (System.IO.File.Exists(@"c:\IApteka\IApteka.ini"))
            { path1 = @"c:\IApteka\IApteka.ini"; }
            if (System.IO.File.Exists(@"D:\IApteka\IApteka.ini"))
            { path1 = @"D:\IApteka\IApteka.ini"; }
            if (System.IO.File.Exists(@"E:\IApteka\IApteka.ini"))
            { path1 = @"E:\IApteka\IApteka.ini"; }
            foreach (string line in System.IO.File.ReadLines(path1))
            {
                if (line.IndexOf("Path") == 0)
                {
                    strConnect = line.Substring(5, line.Length - 5);
                }

            }
            return (strConnect);
        }

        string  getGTIN(string P_ID)
        {
            string gtin = " ";
            string strSQL = @" select ms.sgtin
                               from   mark_sgtin ms
                               where  ms.iid ='"+ P_ID+@"'";
            DataSet dsSGTIN = GetSelect(strSQL);

            foreach (DataRow row in dsSGTIN.Tables[0].Rows)
            {
                gtin = row.ItemArray[0].ToString();
                gtin = gtin.Substring(0, 14);
                break;
            }
                return (gtin);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            int i= listBox1.SelectedIndex;
            string SqlString = @"select d.docno,        d.docdate,      m.med_name,      v.vendor_name,
                                        i.nds,          di.floatqtty,   i.divisor,       i.reg_price,
                                        i.vprice,       i.sprice,       i.rprice,        i.seria,
                                        i.reg_sert_num, i.valid_date,   i.gtd,           m.med_id,
                                        i.sert_num ,    k.country_name,  i.iid

                                 from   docs d, docitem di, items i, medicine m, vendor v, country k

                                 where  d.doc_id         = di.doc_id 
                                        and d.doc_id     = " + strNdoc[i].ToString() + @" 
                                        and i.iid        = di.iid
                                        and m.med_id     = i.med_id 
                                        and v.vendor_id  = m.vendor_id
                                        and k.country_id = v.country_id";

            DataSet dsDocs = GetSelect(SqlString);
            string Fileout = "";
            saveFileDialog1.DefaultExt = "DBF";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Fileout = saveFileDialog1.FileName;
                if (!File.Exists("dzo.dbf"))
                { MessageBox.Show("Нет файла шаблона выгрузки... "); return; }
            }

            File.Copy("dzo.dbf", @Fileout, true);

            OdbcConnection conn = new OdbcConnection();
  //          conn.ConnectionString = @"Dsn=dBASE Files;Driver ={Microsoft dBASE Driver (*.dbf)}; SourceType=dBASE; datasource = Fileout ";
            conn.ConnectionString = @"Dsn=Db4; Driver={Microsoft dBase Driver(*.dbf)};SourceType=dBASE Files;datasource=" + Fileout + "; encoding=cp866";

            int j  = 0;
            OdbcCommand com = conn.CreateCommand();
            conn.Open();

            foreach (DataRow row in dsDocs.Tables[0].Rows)
            {
                com.CommandText = @"INSERT INTO " + Fileout +
                @" ( DOCNUM,   DOCDATE, NAME,   MANUF,    REGPRC,   PRICE1,  PRICE3,  PRICE3N,  PRICE1N,
                     QUANTITY, SERIES,  NDS,     sumnds,  sumwonds, sumwnds, p_price, p_nds,    p_sumnds,
                     p_sum,    P_AMT,   CODEPST, CERTNUM, CNTR,     P_ID,    GTD,     LIFETIME, GTIN) VALUES ";

                //string str_ree = "0.00";
                string str_ree = row.ItemArray[7].ToString();
                if (str_ree.TrimEnd()=="") { str_ree = "0.00"; };
                str_ree = str_ree.Replace(',', '.');

                string str_price1 = row.ItemArray[8].ToString();                  // цена изготовителя б/ндс
                decimal price1N = 0;
                try
                { price1N = (decimal)row.ItemArray[8]; }
                catch { price1N = 0; }
                if (str_price1.TrimEnd() == "") { str_price1 = "0.00"; }
                str_price1 = str_price1.Replace(',', '.');

                string str_price3 = row.ItemArray[9].ToString();                 // цена поставщика б/ндс
                str_price3 = str_price3.Replace(',', '.');
                decimal price3N = (decimal)row.ItemArray[9];
                if (str_price3.TrimEnd() == "") { str_price3 = "0.00"; };

                string str_rprice = row.ItemArray[10].ToString();                 // цена продажи с/ндс
                str_rprice      = str_rprice.Replace(',', '.');
                decimal rprice  = (decimal)row.ItemArray[10];
                if (str_rprice.TrimEnd() == "") { str_rprice = "0.00"; };
                decimal sum     = -1*rprice * (decimal)row.ItemArray[5];             // сумма розница с ндс
                string str_sum  = sum.ToString();
                str_sum         = str_sum.Replace(',', '.');
                string seria    = row.ItemArray[11].ToString();                      // Серия
                string sertnum  = row.ItemArray[12].ToString();                      // Сертификат
                string godendo  = row.ItemArray[13].ToString();                      // Срок годности
                int t = godendo.LastIndexOf("202") - 6;
                godendo = godendo.Substring(t, 10);
                string gtd      = row.ItemArray[14].ToString();                    
                string kodtovar = row.ItemArray[15].ToString();                      // код товара 
                sertnum         = sertnum + row.ItemArray[16].ToString();            // 
                string country = row.ItemArray[17].ToString();                       // Страна
                string P_ID    = row.ItemArray[18].ToString();                      // ID партии товара

                decimal klv = -1*(decimal)row.ItemArray[5];
                string str_klv = klv.ToString();
                if (str_klv.TrimEnd() == "") { str_klv = "0.000"; };
                str_klv = str_klv.Replace(',', '.');

                
                string str_nds = row.ItemArray[4].ToString();
                decimal nds = (decimal)row.ItemArray[4];
                if (str_nds.TrimEnd() == "") { str_nds = "0"; };
                str_nds = str_nds.Replace(',', '.');

                price3N = price3N * (1 + nds / 100);
                string str_price3N = price3N.ToString();
                str_price3N = str_price3N.Replace(',', '.');

                price1N = price1N * (1 + nds / 100);
                string str_price1N = price1N.ToString();
                str_price1N = str_price1N.Replace(',', '.');


                decimal sumnds = (decimal)row.ItemArray[10] * nds / (nds+100);             // сумма ндс продажи ед 
                string str_sumnds = sumnds.ToString();
                str_sumnds = str_sumnds.Replace(',', '.');

                decimal sumwnds = (decimal)row.ItemArray[10] * (decimal)row.ItemArray[5];  //сумма продажи с ндс
                sumwnds = sumwnds * -1;
                string str_sumwnds = sumwnds.ToString();
                if (str_sumwnds.TrimEnd() == "") { str_sumwnds = "0.00"; };
                str_sumwnds = str_sumwnds.Replace(',', '.');

                decimal sumwonds = sumwnds - (sumnds * (decimal)row.ItemArray[5]);          // количество
  //              sumwonds = sumwonds * -1;
                string str_sumwonds = sumwonds.ToString();
                if (str_sumwonds.TrimEnd() == "") { str_sumwonds = "0.00"; };
                str_sumwonds = str_sumwonds.Replace(',', '.');
                /*
                                P_AMT = 

                                P_PCS
                */

                string gtin = getGTIN(P_ID);

                com.CommandText = com.CommandText + @"('" +
                                  row.ItemArray[0].ToString() + @"', '" +                       //DOCNUM
                                  row.ItemArray[1].ToString().Substring(0, 10) + @"', '" +      //DOCDATE
                                  row.ItemArray[2].ToString() + @"', '" +                       //NAME
                                  row.ItemArray[3].ToString() + @"', " +                       //MANUF
                                  str_ree      + ", " +                                          //REGPRC
                                  str_price1   + ", " +          // произв б/ндс                 //PRICE1
                                  str_price3   + ", " +          // поставщик б/ндс              //PRICE3
                                  str_price3N  + ", " +          // поставщик с ндс             //PRICE3N
                                  str_price1N  + ", " +          // производитель с ндс         //PRICE1N
                                  str_klv      + @",'" +                                         //QUANTITY
                                  seria        + @"', " +                                        //SERIES
                                  str_nds      + ", " +
                                  str_sumnds   + ", " +
                                  str_sumwonds + ", " +
                                  str_sumwnds  + ", " +
                                  str_price3N  + ", " +
                                  str_nds      + ", " +
                                  str_sumnds   + ", " +
                                  str_sum      + ","  +
                                  str_klv      + @", '" +
                                  kodtovar     + @"','" +
                                  sertnum      + @"','" +
                                  country      + @"','" +
                                  P_ID         + @"','" +
                                  gtd          + @"','" +
                                  godendo      + @"','" +
                                  gtin         + @"'" +
                                  ")";
                j = com.ExecuteNonQuery();

            }
            conn.Close();
            dsDocs.Dispose();
            MessageBox.Show("Готово!" );

        }


        private void ExportDoc_Load(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }
    }
}
