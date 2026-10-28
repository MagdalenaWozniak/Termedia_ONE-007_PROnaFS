using Hydra;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using ZZnaCSV;
using System.Collections.Generic;
using static ZZnaCSV.FS;
using System.Text;
using System;
using System.Data.SqlClient;
using System.Data;
using cdn_api;

[assembly: CallbackAssemblyDescription(Name: "ONE-007 PRO na FS",
                                           Description: "Przekształcanie proformy do faktury",
                                           Author: "OneSupport",
                                           Version: "1.0",
                                           SystemVersion: "2021.0.2",
                                           Created: "24-03-2021")]


namespace ZZnaCSV
{
    [SubscribeProcedure(Procedures.ListaDokumentow, "Proforma na fakturę")]
    public class ZSnaFS_Class : Callback
    {
        private int XLWersjaAPI = 20202;
        private int XLTrybWsadowy = 1;
        private int XLWinieta = -1;
        private int XLUtworzWlasnaSesje = 1;
        private int sesja = -1;
        private string XLSQLServer = "192.168.4.24";
        private string XLSQLDatabase = "ERPXL_Termedia_Test";//"ERPXL_PoloMarket"
        private string XLUser = "ADMIN"; //
        private string XLPass = "NE!w0rk12";//"Dupa2003!";
        private string connectionString = @"Data Source=192.168.4.24;Initial Catalog=ERPXL_Termedia_Test; User ID=onesupport;Password=a956!2A@";
        private ClaWindow RapButton;
        private ClaWindow zakladkaTabWewnetrzne;
        private ClaWindow filtrBtn;
        //private List<ListaZS> ZSy = new List<ListaZS>();
        private int listID;
        private bool bSpacja = false;
        long iNumer;
        long iTyp;
        long iFirma;
        long iLp;
        private ClaWindow listaPoz;
        public override void Init()
        {

            AddSubscription(true, 0, Events.JustAfterWindowOpening, new TakeEventDelegate(OnOpenWindow));
            AddSubscription(false, 0, Events.ResizeWindow, new TakeEventDelegate(Resize));
        }



        private bool OnOpenWindow(Procedures ProcId, int ControlId, Events Event)
        {

            ClaWindow Parent = GetWindow();
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Rectangle WndBnd = GetWindow().Bounds;
            zakladkaTabWewnetrzne = GetWindow().AllChildren["?TabSprzedaz"];
            listaPoz = GetWindow().AllChildren["?TabHan"];
            RapButton = zakladkaTabWewnetrzne.Children.Add(ControlTypes.button);
            RapButton.Visible = true;
            RapButton.TextRaw = "FS";
            RapButton.ToolTipRaw = "Przekształcanie proformy do faktury";
            //RapButton.IconRaw = "C:\\Onesupport\\PROnaFS\\proforma.ico";

            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 20, filtrBtn.Bounds.Y, 18, 15);
          //tu sprawdzic

            RapButton.OnAfterMouseDown += new TakeEventDelegate(BtnOnMouseDown);
            RapButton.OnBeforeAccepted += new TakeEventDelegate(zaczytaj_OnAfterAccepted);
            

            //pobrane z pola w zakładce Raporty
            return true;
        }

        public bool Resize(Procedures ProcId, int ControlId, Events Event)
        {
            filtrBtn = GetWindow().AllChildren["?BRW1::TotalWorking"];                            //pobiera współrzędne przycisku zapisz
            RapButton.Bounds = new Rectangle(filtrBtn.Bounds.X - 20, filtrBtn.Bounds.Y, 18, 15);

            return true;
        }

        bool zaczytaj_OnAfterAccepted(Procedures ProcedureId, int ControlId, Events Event)
        {
            bSpacja = true;
            Runtime.WindowController.PostEvent(listaPoz.Id, Events.NewSelection);
            iNumer = TraNag.TrN_GIDNumer;
            iFirma = TraNag.TrN_GIDFirma;
            iLp = TraNag.TrN_GIDLp;
            iTyp = TraNag.TrN_GIDTyp;
            

            return true;
        }

        public bool BtnOnMouseDown(Procedures ProcId, int ControlId, Events Event)
        {
            int dokumentID = 0;
            try
            {

                //ADODB._Recordset recordset = Runtime.WindowController.GetQueueMarked((int)ProcId, listID, GetCallbackThread());

               // MessageBox.Show(Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 18));
                if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 19) == "ERPXL_Termedia_TEST")
                {
                    //MessageBox.Show("Jestem tutaj", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    XLSQLDatabase = "Termedia TEST";
                    connectionString = @"Data Source=192.168.4.24;Initial Catalog=ERPXL_Termedia_Test; User ID=onesupport;Password=a956!2A@";
                }
                else if (Runtime.ConfigurationDictionary.Baza.ToString().Substring(0, 14) == "ERPXL_Termedia")
                {
                    XLSQLDatabase = "Termedia";
                    connectionString = @"Data Source=192.168.4.24;Initial Catalog=ERPXL_Termedia; User ID=onesupport;Password=a956!2A@";
                }
                else
                {
                    MessageBox.Show("Dodatek nie posiada funkcjonalności dla tej bazy.", "Komunikat",
                     MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;

                }

                #region logowanie
                XLLoginInfo_20202 nowyUzytkownik = new XLLoginInfo_20202();
                nowyUzytkownik.Wersja = XLWersjaAPI;
                nowyUzytkownik.Winieta = XLWinieta;
                nowyUzytkownik.TrybWsadowy = XLTrybWsadowy;
                nowyUzytkownik.UtworzWlasnaSesje = XLUtworzWlasnaSesje;
                nowyUzytkownik.SerwerKlucza = XLSQLServer;
                nowyUzytkownik.Baza = XLSQLDatabase;
                nowyUzytkownik.OpeIdent = XLUser;
                nowyUzytkownik.OpeHaslo = XLPass;
                nowyUzytkownik.PortHTTPSerweraKlucza = 1;
                nowyUzytkownik.RodzajSerweraKlucza = 0;
                nowyUzytkownik.SerwerKlucza = "localhost::5000185870";
                nowyUzytkownik.ProgramID = "ONE-007 PRO na FS";

                int blad = cdn_api.cdn_api.XLLogin(nowyUzytkownik, ref sesja);

                if (blad != 0)
                {
                    MessageBox.Show("Błąd logowania: "+blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                #endregion

                FS faktura = new FS(Convert.ToInt32(iNumer), Convert.ToInt32(iTyp), Convert.ToInt32(iFirma), (short)iLp);
                bool czyPro = faktura.czyPRO(connectionString);
                bool czyMaFs = faktura.czyMaFS(connectionString);
                if(czyPro && !czyMaFs)
                {
                    #region Otwieranie dokumentu
                    XLOtwarcieNagInfo_20202 nagDokument = new XLOtwarcieNagInfo_20202();
                    nagDokument.GIDFirma =faktura.GIDFirma;
                    nagDokument.GIDLp = faktura.GIDLp;
                    nagDokument.GIDNumer = faktura.GIDNumer;
                    nagDokument.GIDTyp = faktura.GIDTyp;
                    nagDokument.Wersja = XLWersjaAPI;
                    nagDokument.Tryb = 2;

                    
                    int blad2 = cdn_api.cdn_api.XLOtworzDokument(sesja, ref dokumentID, nagDokument);

                    if (blad2 != 0)
                    {
                        MessageBox.Show("Błąd otwierania dokumentu: " + blad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    #endregion

                    string LpQuery = "SELECT [TrE_GIDLp] FROM [CDN].[TraElem] WHERE TrE_GIDTyp=" + faktura.GIDTyp + " and TrE_GIDFirma=" + faktura.GIDFirma + " and TrE_GIDNumer=" + faktura.GIDNumer + " ORDER BY [TrE_GIDLp] DESC";
                    List<string> lp = SelectList(LpQuery, "CDN.TraElem", connectionString);

                    bool zlyTyp = false;
                    foreach (string pozycja in lp)
                    {
                        string typTwr=faktura.TowarLubUsluga(pozycja, connectionString);
                        int iTypTwr = Convert.ToInt32(typTwr);
                        //MessageBox.Show("Typ towaru: "+typTwr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        switch(iTypTwr)
                        {
                            case 1:
                                {
                                    //towar
                                    bool US =faktura.czyJestUS(pozycja, connectionString);
                                    if(US==false)
                                    {
                                        zlyTyp = true;
                                        MessageBox.Show("Przynajmniej dla jednej pozycji nie znaleziono usługi. \nAnulowano działanie dodatku", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    break;
                                }
                            case 4:
                                {
                                    break;
                                }
                            default:
                                {
                                    zlyTyp = true;
                                    MessageBox.Show("Przynajmniej jedna z pozycji posiada zły typ towaru. \nAnulowano działanie dodatku", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    break;
                                }
                        }
                        if (zlyTyp)
                            break;
                    }

                    string nrFaktury = "";


                    if (!zlyTyp)
                    {
                        nrFaktury = dodawanieDokumentu(faktura, lp);


                        string twrKod = "";
                        foreach (string pozycja in lp)
                        {
                            string typTwr = faktura.TowarLubUsluga(pozycja, connectionString);
                            int iTypTwr = Convert.ToInt32(typTwr);
                            string traElemQuery = "SELECT * FROM CDN.TraElem WHERE TrE_GIDTyp=" + faktura.GIDTyp + " and TrE_GIDFirma=" + faktura.GIDFirma + " and TrE_GIDNumer=" + faktura.GIDNumer + " and TrE_GIDLp = " + pozycja;
                            List<string> traElem = SelectListColumn(traElemQuery, "CDN.TraElem", connectionString);
                            //towar
                            //bool US = faktura.czyJestUS(pozycja, connectionString);
                            if (iTypTwr == 1)
                            {
                                twrKod = traElem[10].TrimEnd() + " US";
                                //MessageBox.Show("TwrKod z towaru: " + twrKod, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                #region usuwanie pozycji
                                XLUsunElemInfo_20202 usunElem = new XLUsunElemInfo_20202();
                                usunElem.Pozycja = Convert.ToInt32(pozycja);
                                usunElem.Wersja = XLWersjaAPI;
                                int usunBlad = cdn_api.cdn_api.XLUsunPozycje(dokumentID, usunElem);

                                if (usunBlad != 0)
                                    MessageBox.Show("Błąd usuwania pozycji: " + usunBlad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                #endregion

                                #region dodawanie pozycji
                                XLDokumentElemInfo_20202 nowyTowar = new XLDokumentElemInfo_20202();
                                nowyTowar.TowarKod = twrKod;
                                nowyTowar.Ilosc = (traElem[22] != null ? (traElem[22]) : "");
                                //nowyTowar.JmFormat = (traElem[23] != null ? Convert.ToInt32(traElem[23]) : 0);
                                nowyTowar.JmZ = (traElem[24] != null ? traElem[24] : "");
                                //nowyTowar.TypJm = (traElem[26] != null ? Convert.ToInt32(traElem[26]) : 0);
                                //nowyTowar.PrzeliczM = (traElem[27] != null ? Convert.ToInt32(traElem[27]) : 0);
                                //nowyTowar.PrzeliczL = (traElem[28] != null ? Convert.ToInt32(traElem[28]) : 0);
                                //nowyTowar.StawkaPod = (traElem[30] != null ? Convert.ToInt32(traElem[30]) : 0);
                                //nowyTowar.FlagaVAT = (traElem[31] != null ? Convert.ToInt32(traElem[31]) : 0);
                                nowyTowar.Waluta = (traElem[34] != null ? traElem[34] : "");
                                //nowyTowar.KursNr = (traElem[35] != null ? Convert.ToInt32(traElem[35]) : 0);
                                //nowyTowar.KursM = (traElem[36] != null ? Convert.ToInt32(traElem[36]) : 1);
                                //nowyTowar.KursL = (traElem[37] != null ? Convert.ToInt32(Convert.ToDecimal(traElem[37]) * 100) : 1);
                                //nowyTowar.Rabat = (traElem[40] != null ? Convert.ToInt32(Convert.ToDecimal(traElem[40]) * 100) : 0);
                                //nowyTowar.RecTyp = (traElem[59] != null ? Convert.ToInt32(traElem[59]) : 0);
                                //nowyTowar.RecFirma = (traElem[60] != null ? Convert.ToInt32(traElem[60]) : 0);
                                //nowyTowar.RecNumer = (traElem[61] != null ? Convert.ToInt32(traElem[61]) : 0);
                                //nowyTowar.RecLp = (traElem[62] != null ? Convert.ToInt32(traElem[62]) : 0);
                                nowyTowar.Cena = (traElem[63] != null ? traElem[63] : "");
                                /*nowyTowar.Rownanie = (traElem[64] != null ? Convert.ToInt32(traElem[64]) : 0);
                                nowyTowar.OdKsiegowych = (traElem[65] != null ? Convert.ToInt32(traElem[65]) : 0);
                                nowyTowar.PCN = (traElem[67] != null ? traElem[67] : "");
                                nowyTowar.PrecyzjaCeny = (traElem[68] != null ? Convert.ToInt32(traElem[68]) : 0);
                                nowyTowar.Kraj = (traElem[76] != null ? (traElem[76]) : "");
                                nowyTowar.CenaSpr = (traElem[89] != null ? Convert.ToInt32(traElem[89]) : 0);
                                nowyTowar.CChTyp = (traElem[91] != null ? Convert.ToInt32(traElem[91]) : 0);
                                nowyTowar.CChNumer = (traElem[92] != null ? Convert.ToInt32(traElem[92]) : 0);
                                nowyTowar.Cecha = (traElem[93] != null ? traElem[93] : "");
                                nowyTowar.PakietId = (traElem[97] != null ? Convert.ToInt32(traElem[97]) : 0);
                                nowyTowar.Gratis = (traElem[98] != null ? Convert.ToInt32(traElem[98]) : 0);
                                nowyTowar.PromocjaProgID = (traElem[99] != null ? Convert.ToInt32(traElem[99]) : 0);
                                nowyTowar.Zlom = (traElem[101] != null ? Convert.ToInt32(traElem[101]) : 0);*/
                                nowyTowar.Wersja = XLWersjaAPI;

                                int blad_towar = cdn_api.cdn_api.XLDodajPozycje(dokumentID, nowyTowar);
                                if (blad_towar != 0)
                                    MessageBox.Show("Błąd podczas dodawania pozycji API: " + blad_towar, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                #endregion
                            }

                            traElem.Clear();

                        }


                        faktura.updateAtrybut(nrFaktury, connectionString);

                        #region zamykanie dokumentu

                        lp.Clear();
                        XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                        closeDokument.Wersja = XLWersjaAPI;
                        closeDokument.Tryb = 1;
                        int closeBlad = cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
                        if (closeBlad != 0)
                            MessageBox.Show("Błąd podczas zamykania dokumentu PRO (API): " + closeBlad, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        #endregion
                    }

                }
                else
                {
                    MessageBox.Show("Podświetlona faktura nie jest Proformą i/lub posiada już przekształconą fakturę", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


                
               // MessageBox.Show("Dane: \nGIDNumer " + iNumer + ", GIDTyp " + iTyp + ", GIDFirma " + iFirma+", GIDLp "+iLp, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
              

                cdn_api.cdn_api.XLLogout(sesja);

            }
            catch {
                XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                closeDokument.Wersja = XLWersjaAPI;
                cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
            }
           
            MessageBox.Show("Zakończono dodawanie faktury", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return true;
        }

       
        public override void Cleanup()
        {

        }

        public string dodawanieDokumentu (FS faktura, List<string> lp)
        {

            int dokumentID = 0;
            try
            {
               // MessageBox.Show("Dodawanie faktury", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                string traNagQuery = "SELECT * FROM CDN.TraNag WHERE TrN_GIDTyp=" + faktura.GIDTyp + " and TrN_GIDFirma=" + faktura.GIDFirma + " and TrN_GIDNumer=" + faktura.GIDNumer + " and TrN_GIDLp = " + faktura.GIDLp;
                List<string> traNag = SelectListColumn(traNagQuery, "CDN.TraNag", connectionString);
                // MessageBox.Show("KntTyp = "+(traNag[16] != null ? (Convert.ToInt32(traNag[16])) : 0)+"\nKntFirma = "+(traNag[17] != null ? (Convert.ToInt32(traNag[17])) : 0)+ "\nKntNumer = "+(traNag[18] != null ? (Convert.ToInt32(traNag[18])) : 0)+ "\nKntLp = "+(traNag[19] != null ? (Convert.ToInt32(traNag[19])) : 0), "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DateTime data = DateTime.Now;
                string miesiac = data.ToString("MM");
                string rok = data.ToString("yyyy");
                string dataWys = data.ToString("yyyy-MM-dd");
                int dataDok = faktura.getDate(connectionString, dataWys);

                #region dodawanie dokumentu
                XLDokumentNagInfo_20202 nowyFS = new XLDokumentNagInfo_20202();
                nowyFS.GIDFirma = 0;
                nowyFS.GIDLp = 0;
                nowyFS.GIDNumer = 0;
                nowyFS.GIDTyp = 0;
                nowyFS.Typ = 2033;
                nowyFS.Tryb = 1;
                nowyFS.Wersja = XLWersjaAPI;
                nowyFS.Seria = "<PUS>";
                nowyFS.ZwrTyp = (traNag[4]!=null ? (Convert.ToInt32(traNag[4])) : 0);
                nowyFS.ZwrFirma = (traNag[5] != null ? (Convert.ToInt32(traNag[5])) : 0);
                nowyFS.ZwrNumer = (traNag[6] != null ? (Convert.ToInt32(traNag[6])) : 0);
                nowyFS.ZwrLp = (traNag[7] != null ? (Convert.ToInt32(traNag[7])) : 0);
                nowyFS.KntTyp = (traNag[16] != null ? (Convert.ToInt32(traNag[16])) : 0);
                nowyFS.KntFirma = (traNag[17] != null ? (Convert.ToInt32(traNag[17])) : 0);
                nowyFS.KntNumer = (traNag[18] != null ? (Convert.ToInt32(traNag[18])) : 0);
                nowyFS.KntLp = (traNag[19] != null ? (Convert.ToInt32(traNag[19])) : 0);
                nowyFS.AkwTyp = (traNag[24] != null ? (Convert.ToInt32(traNag[24])) : 0);
                nowyFS.AkwFirma = (traNag[25] != null ? (Convert.ToInt32(traNag[25])) : 0);
                nowyFS.AkwNumer = (traNag[26] != null ? (Convert.ToInt32(traNag[26])) : 0);
                nowyFS.AkwLp = (traNag[27] != null ? (Convert.ToInt32(traNag[27])) : 0);
                nowyFS.AdwTyp = (traNag[28] != null ? (Convert.ToInt32(traNag[28])) : 0);
                nowyFS.AdwFirma = (traNag[29] != null ? (Convert.ToInt32(traNag[29])) : 0);
                nowyFS.AdwNumer = (traNag[30] != null ? (Convert.ToInt32(traNag[30])) : 0);
                nowyFS.AdwLp = (traNag[31] != null ? (Convert.ToInt32(traNag[31])) : 0);
                nowyFS.ZamTyp = (traNag[36] != null ? (Convert.ToInt32(traNag[36])) : 0);
                nowyFS.ZamFirma = (traNag[37] != null ? (Convert.ToInt32(traNag[37])) : 0);
                nowyFS.ZamNumer = (traNag[38] != null ? (Convert.ToInt32(traNag[38])) : 0);
                nowyFS.ZamLp = (traNag[39] != null ? (Convert.ToInt32(traNag[39])) : 0);
                nowyFS.SaNTyp = (traNag[40] != null ? (Convert.ToInt32(traNag[40])) : 0);
                nowyFS.SaNFirma = (traNag[41] != null ? (Convert.ToInt32(traNag[41])) : 0);
                nowyFS.SaNNumer = (traNag[42] != null ? (Convert.ToInt32(traNag[42])) : 0);
                nowyFS.SaNLp = (traNag[43] != null ? (Convert.ToInt32(traNag[43])) : 0);
                nowyFS.KonTyp = (traNag[44] != null ? (Convert.ToInt32(traNag[44])) : 0);
                nowyFS.KonNumer = (traNag[46] != null ? (Convert.ToInt32(traNag[46])) : 0);
                nowyFS.Data = dataDok; //(traNag[53] != null ? (Convert.ToInt32(traNag[53])) : 0);
                nowyFS.DataSpr = dataDok; //(traNag[54] != null ? (Convert.ToInt32(traNag[54])) : 0);
                nowyFS.Rok = Convert.ToInt32(rok); //(traNag[56] != null ? (Convert.ToInt32(traNag[56])) : 0);
                nowyFS.Miesiac = Convert.ToInt32(miesiac); //(traNag[57] != null ? (Convert.ToInt32(traNag[57])) : 0);
                nowyFS.Rejestr= (traNag[62] != null ? (traNag[62]) : "");
                nowyFS.RokVAT = (traNag[63] != null ? (Convert.ToInt32(traNag[63])) : 0);
                nowyFS.MiesiacVAT = (traNag[64] != null ? (Convert.ToInt32(traNag[64])) : 0);
                nowyFS.Rozliczac = (traNag[70] != null ? (Convert.ToInt32(traNag[70])) : 0);
                //nowyFS.DokumentObcy = (traNag[71] != null ? (traNag[71]) : "");
                nowyFS.OdliczeniaVat = (traNag[75] != null ? (Convert.ToInt32(traNag[75])) : 0);
                nowyFS.RodzajZakupu = (traNag[76] != null ? (Convert.ToInt32(traNag[76])) : 0);
                nowyFS.ExpoNorm = (traNag[77] != null ? (Convert.ToInt32(traNag[77])) : 0);
                nowyFS.NB = (traNag[79] != null ? (traNag[79]) : "");
                nowyFS.Fiskalny = (traNag[81] != null ? (Convert.ToInt32(traNag[81])) : 0);
                nowyFS.Forma = (traNag[82] != null ? (Convert.ToInt32(traNag[82])) : 0);
                nowyFS.Rabat = (traNag[85] != null ? Convert.ToInt32((Convert.ToDecimal(traNag[85])*100)) : 0);
                nowyFS.SposobDst = (traNag[92] != null ? (traNag[92]) : "");
                nowyFS.Waluta = (traNag[117] != null ? (traNag[117]) : "");
                nowyFS.KursNr = (traNag[118] != null ? (Convert.ToInt32(traNag[118])) : 0);
                nowyFS.KursM = (traNag[119] != null ? (Convert.ToInt32(traNag[119])) : 1);
                nowyFS.KursL = (traNag[120] != null ? (Convert.ToInt32(Convert.ToDecimal(traNag[120]))*100) : 1);
                nowyFS.FRSID = ((traNag[125] != null && traNag[125]!="") ? (Convert.ToInt32(traNag[125].Trim())) : 0);
                nowyFS.Cecha = (traNag[126] != null ? (traNag[126]) : "");
                nowyFS.KnDTyp = (traNag[134] != null ? (Convert.ToInt32(traNag[134])) : 0);
                nowyFS.KnDFirma = (traNag[135] != null ? (Convert.ToInt32(traNag[135])) : 0);
                nowyFS.KnDNumer = (traNag[136] != null ? (Convert.ToInt32(traNag[136])) : 0);
                nowyFS.KnDLp = (traNag[137] != null ? (Convert.ToInt32(traNag[137])) : 0);
                nowyFS.URL = (traNag[142] != null ? (traNag[142]) : "");
                nowyFS.Termin = (traNag[144] != null ? (Convert.ToInt32(traNag[144])) : 0);
                nowyFS.KodRodzajuTransportu = (traNag[150] != null ? (traNag[150]) : "");
                nowyFS.KarNumer = (traNag[132] != null ? (Convert.ToInt32(traNag[132])) : 0);
                


                int bladDok = cdn_api.cdn_api.XLNowyDokument(sesja, ref dokumentID, nowyFS);
                if(bladDok!=0)
                    MessageBox.Show("Błąd podczas dodawania dokumentu API: " + bladDok, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                #endregion

                foreach (string pozycja in lp)
                {
                    string typTwr = faktura.TowarLubUsluga(pozycja, connectionString);
                    int iTypTwr = Convert.ToInt32(typTwr);
                    //MessageBox.Show("Typ towaru: "+typTwr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    string traElemQuery = "SELECT * FROM CDN.TraElem WHERE TrE_GIDTyp=" + faktura.GIDTyp + " and TrE_GIDFirma=" + faktura.GIDFirma + " and TrE_GIDNumer=" + faktura.GIDNumer + " and TrE_GIDLp = " + pozycja;
                    List<string> traElem = SelectListColumn(traElemQuery, "CDN.TraElem", connectionString);
                    string twrKod = traElem[10].TrimEnd(); ;

                    if(iTypTwr==4)
                    {
                     
                               
                                //usługa
                                string nazwaUslugi = faktura.czyUslugaBezUS(pozycja, connectionString);
                                 if (nazwaUslugi.Contains(" US"))
                                 {

                                     //MessageBox.Show("Zawiera US", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                     bool twr = faktura.czyJestTwr(pozycja, connectionString, nazwaUslugi, traElem[22]);
                                     if(twr)
                                     {
                                        
                                        twrKod = twrKod.Substring(0,twrKod.Length-3);
                                       // MessageBox.Show("TwrKod z usługi: " + twrKod, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                        //MessageBox.Show("Przynajmniej dla jednej pozycji usługowej nie znaleziono towaru w zasobie. \nAnulowano działanie dodatku", "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                    }
                                 }
                                else
                                {
                                    //MessageBox.Show("Nie zawiera US: "+nazwaUslugi, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                             
                            
                    }


                    #region dodawanie pozycji

                    XLDokumentElemInfo_20202 nowyTowar = new XLDokumentElemInfo_20202();
                    nowyTowar.TowarKod = twrKod;
                    nowyTowar.Ilosc = (traElem[22] != null ? (traElem[22]) : "");
                    //nowyTowar.JmFormat = (traElem[23] != null ? Convert.ToInt32(traElem[23]) : 0);
                    nowyTowar.JmZ = (traElem[24] != null ? traElem[24] : "");
                    //nowyTowar.TypJm = (traElem[26] != null ? Convert.ToInt32(traElem[26]) : 0);
                    //nowyTowar.PrzeliczM = (traElem[27] != null ? Convert.ToInt32(traElem[27]) : 0);
                    //nowyTowar.PrzeliczL = (traElem[28] != null ? Convert.ToInt32(traElem[28]) : 0);
                    //nowyTowar.StawkaPod = (traElem[30] != null ? Convert.ToInt32(traElem[30]) : 0);
                    //nowyTowar.FlagaVAT = (traElem[31] != null ? Convert.ToInt32(traElem[31]) : 0);
                    nowyTowar.Waluta = (traElem[34] != null ? traElem[34] : "");
                    //nowyTowar.KursNr = (traElem[35] != null ? Convert.ToInt32(traElem[35]) : 0);
                    //nowyTowar.KursM = (traElem[36] != null ? Convert.ToInt32(traElem[36]) : 1);
                    //nowyTowar.KursL = (traElem[37] != null ? Convert.ToInt32(Convert.ToDecimal(traElem[37])*100) : 1);
                    //nowyTowar.Rabat = (traElem[40] != null ? Convert.ToInt32(Convert.ToDecimal(traElem[40])*100) : 0);
                    //nowyTowar.RecTyp = (traElem[59] != null ? Convert.ToInt32(traElem[59]) : 0);
                    //nowyTowar.RecFirma = (traElem[60] != null ? Convert.ToInt32(traElem[60]) : 0);
                    //nowyTowar.RecNumer = (traElem[61] != null ? Convert.ToInt32(traElem[61]) : 0);
                    //nowyTowar.RecLp = (traElem[62] != null ? Convert.ToInt32(traElem[62]) : 0);
                    nowyTowar.Cena = (traElem[63] != null ? traElem[63] : "");
                   /* nowyTowar.Rownanie = (traElem[64] != null ? Convert.ToInt32(traElem[64]) : 0);
                    nowyTowar.OdKsiegowych = (traElem[65] != null ? Convert.ToInt32(traElem[65]) : 0);
                    nowyTowar.PCN = (traElem[67] != null ? traElem[67] : "");
                    nowyTowar.PrecyzjaCeny = (traElem[68] != null ? Convert.ToInt32(traElem[68]) : 0);
                    nowyTowar.Kraj = (traElem[76] != null ? (traElem[76]) : "");
                    nowyTowar.CenaSpr = (traElem[89] != null ? Convert.ToInt32(traElem[89]) : 0);
                    nowyTowar.CChTyp = (traElem[91] != null ? Convert.ToInt32(traElem[91]) : 0);
                    nowyTowar.CChNumer = (traElem[92] != null ? Convert.ToInt32(traElem[92]) : 0);
                    nowyTowar.Cecha = (traElem[93] != null ? traElem[93] : "");
                    nowyTowar.PakietId = (traElem[97] != null ? Convert.ToInt32(traElem[97]) : 0);
                    nowyTowar.Gratis = (traElem[98] != null ? Convert.ToInt32(traElem[98]) : 0);
                    nowyTowar.PromocjaProgID = (traElem[99] != null ? Convert.ToInt32(traElem[99]) : 0);
                    nowyTowar.Zlom = (traElem[101] != null ? Convert.ToInt32(traElem[101]) : 0);*/
                    nowyTowar.Wersja = XLWersjaAPI;

                    int blad_towar = cdn_api.cdn_api.XLDodajPozycje(dokumentID, nowyTowar);
                    if (blad_towar != 0)
                        MessageBox.Show("Błąd podczas dodawania pozycji API: " + blad_towar, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    #endregion


                    traElem.Clear();

                }

                string nrFaktury = faktura.numerFs(nowyFS, connectionString);

                if (bladDok == 0)
                {
                    #region dodawanie atrybutów

                    XLAtrybutInfo_20202 atr1 = new XLAtrybutInfo_20202();
                    atr1.GIDFirma = nowyFS.GIDFirma;
                    atr1.GIDLp = nowyFS.GIDLp;
                    atr1.GIDNumer = nowyFS.GIDNumer;
                    atr1.GIDTyp = nowyFS.GIDTyp;
                    atr1.GIDSubLp = 0;
                    atr1.Klasa = "Nr Proformy";
                    atr1.Wartosc = (traNag[71] != null ? (traNag[71]) : "");
                    atr1.Wersja = XLWersjaAPI;

                    int bladAtr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atr1);
                    if (bladAtr != 0)
                        MessageBox.Show("Błąd podczas dodawania atrybutu Nr Proformy: " + bladAtr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);


                    

                    XLAtrybutInfo_20202 atr2 = new XLAtrybutInfo_20202();
                    atr2.GIDFirma = nowyFS.GIDFirma;
                    atr2.GIDLp = nowyFS.GIDLp;
                    atr2.GIDNumer = nowyFS.GIDNumer;
                    atr2.GIDTyp = nowyFS.GIDTyp;
                    atr2.GIDSubLp = 0;
                    atr2.Klasa = "Nr Faktury Prenumeraty";
                    atr2.Wartosc = nrFaktury;
                    atr2.Wersja = XLWersjaAPI;

                    bladAtr = cdn_api.cdn_api.XLDodajAtrybut(sesja, atr2);
                    if (bladAtr != 0)
                        MessageBox.Show("Błąd podczas dodawania atrybutu Nr Proformy: " + bladAtr, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);


                    #endregion
                }


                #region dodawanie JPK
                string jpk = faktura.znajdzJPK(connectionString);
                if(jpk!="")
                {
                    XLJPKV7Info_20202 nowyJpk = new XLJPKV7Info_20202();
                    nowyJpk.Wartosc = jpk;
                    nowyJpk.Wersja = XLWersjaAPI;

                    int bladJpk = cdn_api.cdn_api.XLDodajKodJPKV7(dokumentID, nowyJpk);
                    if (bladJpk != 0)
                        MessageBox.Show("Błąd podczas dodawania atrybutu JPK: " + bladJpk, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                #endregion



                #region zamykanie dokumentu
                XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                closeDokument.Wersja = XLWersjaAPI;
                closeDokument.Tryb = 1;
                cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
                traNag.Clear();

                #endregion

                return nrFaktury;

            }
            catch (Exception e)
            {
                MessageBox.Show("Błąd dodawania faktury: "+e, "Komunikat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                XLZamkniecieDokumentuInfo_20202 closeDokument = new XLZamkniecieDokumentuInfo_20202();
                closeDokument.Wersja = XLWersjaAPI;
                closeDokument.Tryb = 1;
                cdn_api.cdn_api.XLZamknijDokument(dokumentID, closeDokument);
                return "";

            }

        }

        public List<string> SelectList(string query, string srcTableName, string connectionString)
        {
            List<string> str = new List<string>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                        str.Add(row[0].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

        public List<string> SelectListColumn(string query, string srcTableName, string connectionString)
        {
            List<string> str = new List<string>();
            string sttr = "";
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(query, connection);
                    DataSet dataSet = new DataSet();
                    adapter.Fill(dataSet, srcTableName);
                    foreach (DataColumn column in (InternalDataCollectionBase)dataSet.Tables[0].Columns)
                        foreach(DataRow row in (InternalDataCollectionBase)dataSet.Tables[0].Rows)
                        str.Add(row[column].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(sttr = sttr + "Bład SQL:" + ex.ToString() + "\n");
            }
            return str;
        }

    }
}
