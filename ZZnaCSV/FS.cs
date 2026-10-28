using cdn_api;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ZZnaCSV
{
    public class FS
    {
        public int GIDNumer { get; set; }
        public int GIDTyp { get; set; }
        public int GIDFirma {get; set; }
        public short GIDLp { get; set; }

        public FS(int GIDNumer, int GIDTyp, int GIDFirma, short GIDLp)
        {
            this.GIDNumer = GIDNumer;
            this.GIDTyp = GIDTyp;
            this.GIDFirma = GIDFirma;
            this.GIDLp = GIDLp;
        }
        

        public bool czyPRO(string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string proQuery = "SELECT [TrN_TrNSeria] FROM [CDN].[TraNag] WHERE TrN_GIDFirma = " + GIDFirma + " and TrN_GIDTyp = " + GIDTyp + " and TrN_GIDLp = " + GIDLp + " and TrN_GIDNumer =" + GIDNumer;
                wynikQuery = ConnectToSQL(connection_string, proQuery);
                //MessageBox.Show(proQuery +"\nWynik: "+wynikQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (wynikQuery.Contains("PRO"))
                    return true;
                return false;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas sprawdzania czy faktura jest PRO: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        public string TowarLubUsluga(string pozycja, string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string typQuery = "SELECT TrE_TypTwr FROM [CDN].[TraElem] WHERE TrE_GIDTyp = "+GIDTyp+" AND TrE_GIDNumer = "+GIDNumer+" and TrE_GIDLp= "+pozycja;
                wynikQuery = ConnectToSQL(connection_string, typQuery);
                return wynikQuery;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas sprawdzania typu towaru: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return "";
            }
        }

        public bool czyJestUS(string pozycja, string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string usQuery = "SELECT count(*) FROM [CDN].[TwrKarty] WHERE Twr_Kod = (SELECT TrE_TwrKod FROM [CDN].[TraElem] WHERE TrE_GIDTyp =" + GIDTyp + " AND TrE_GIDNumer = " + GIDNumer + " and TrE_GIDLp = "+pozycja+")  +' US'";
                wynikQuery = ConnectToSQL(connection_string, usQuery);
                if (wynikQuery == "1")
                    return true;
                
                return false;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas szukania usługi: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false; 
            }
        }

        public bool czyJestTwr(string pozycja, string connection_string, string twrKod, string ilosc)
        {
            try { 
                string wynikQuery = "";
                twrKod = twrKod.TrimEnd();
                twrKod = twrKod.Substring(0, twrKod.Length - 3);
                //źle to nie będzie ten sam towar tylko towar z kodem bez US, czyli jesteś w dupie xd
                string zasobQuery = "SELECT SUM([TwZ_Ilosc]) FROM [CDN].[TwrZasoby] join CDN.TwrKarty as k on TwZ_TwrFirma = k.Twr_GIDFirma and TwZ_TwrNumer = k.Twr_GIDNumer and TwZ_TwrTyp = k.Twr_GIDTyp WHERE k.Twr_Kod = '"+twrKod+"' ";
                wynikQuery = ConnectToSQL(connection_string, zasobQuery);
                //MessageBox.Show("Twr_Kod: " + twrKod+"\nQuery: "+ zasobQuery+"\nwynik: "+wynikQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (wynikQuery != "")
                {
                    ilosc = ilosc.Substring(0, ilosc.IndexOf(","));
                    wynikQuery =wynikQuery.Substring(0,wynikQuery.IndexOf(","));
                    //MessageBox.Show("Ilosc: " + ilosc + "\nZasob: "+wynikQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (Convert.ToInt32(wynikQuery) >= Convert.ToInt32(ilosc))
                    {
                        //MessageBox.Show("Ilosc: " + wynikQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return true;
                    }
                    return false;
                }
            
                return false;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas szukania zasobu: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        public string czyUslugaBezUS(string pozycja, string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string kodQuery = "SELECT TrE_TwrKod FROM [CDN].[TraElem] WHERE TrE_GIDTyp = " + GIDTyp + " AND TrE_GIDNumer = " + GIDNumer + " and TrE_GIDLp = " + pozycja;
                wynikQuery = ConnectToSQL(connection_string, kodQuery);
                //MessageBox.Show("TwrKod: " + wynikQuery, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return wynikQuery;
               
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd w funkcji czyUslugaBezUS: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return "";
            }

        }

        public string numerFs(XLDokumentNagInfo_20202 nowyFS, string connection_string)
        {
            string wynikQuery = "";
            string nrQuery = "SELECT CDN.NumerDokumentu(TrN_GIDTyp, TrN_SpiTyp, TrN_TrNTyp, TrN_TrNNumer,TrN_TrNRok, TrN_TrNSeria, TrN_TrNMiesiac) FROM CDN.TraNag WHERE TrN_GIDTyp =" + nowyFS.GIDTyp + " AND TrN_GIDNumer = " + nowyFS.GIDNumer + " and TrN_GIDLp = " + nowyFS.GIDLp + " and TrN_GIDFirma = " + nowyFS.GIDFirma;
            wynikQuery = ConnectToSQL(connection_string, nrQuery);
            return wynikQuery;
        }

        public void updateAtrybut (string nrFaktury, string connection_string)
        {
            string wynikQuery = "";
            string updateQuery = "UPDATE [CDN].[Atrybuty] SET [Atr_Wartosc] = '"+nrFaktury+"' WHERE (Atr_ObiTyp="+ GIDTyp +" AND Atr_ObiNumer="+GIDNumer +" and Atr_AtkId=5)";
            wynikQuery = Update(updateQuery, connection_string);
        }

        public bool czyMaFS(string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string fsQuery = "SELECT Atr_Wartosc FROM [CDN].[Atrybuty] WHERE (Atr_ObiTyp=" + GIDTyp + " AND Atr_ObiNumer=" + GIDNumer + " and Atr_AtkId=5)";
                wynikQuery = ConnectToSQL(connection_string, fsQuery);
                if (wynikQuery == "" || wynikQuery.Contains("PRO"))
                    return false;

                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas sprawdzania czy istnieje faktura: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return true;
            }
           
        }

        public string znajdzJPK(string connection_string)
        {
            try
            {
                string wynikQuery = "";
                string jpkQuery = "SELECT [TrJ_Wartosc] FROM .[CDN].[TraJPK] WHERE [TrJ_TrNTyp] = "+GIDTyp+" and [TrJ_TrNFirma]="+ GIDFirma+" and [TrJ_TrNNumer]="+GIDNumer+" and [TrJ_TrNLp]="+GIDLp;
                wynikQuery = ConnectToSQL(connection_string, jpkQuery);
                return wynikQuery;
            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd podczas szukania jpk: " + e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return "";
            }
            
        }

        public int getDate(string connection_string, string date3)
        {
            string date1 = "1800-12-28";
            //string date3 = DateTime.Now.ToString("yyyy-MM-dd");
            string Datequery = "SELECT DATEDIFF(day, '" + date1 + "','" + date3 + "'); ";
            return Convert.ToInt32(ConnectToSQL(connection_string, Datequery));
        }

        public string ConnectToSQL(string connection_string, string query)
        {
            string str = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, "Query");
                    if (dataSet.Tables.Count == 0)
                    {
                        return "";
                    }
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                    {
                        str = row[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                str = str + "Bład SQL:" + ex.ToString() + "\n";
            }
            return str;
        }

        public string Update(string sql, string connection_string)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(connection_string))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    connection.Open();
                    adapter.UpdateCommand = new SqlCommand(sql, connection);
                    adapter.UpdateCommand.ExecuteNonQuery();
                    //MessageBox.Show("Row updated !! ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return sql;
        }

    }
}
