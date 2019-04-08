using cdn_api;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hydra
{
    public partial class Api : Form
    {
        public static string komunikat = "";
        int sesjaId;
        int wersjaApi = 20182;
        int blad = 0;
        List<Gidy> listaDokumentow;

        public Api(List<Gidy> listaDokumentow)
        {
            InitializeComponent();
            this.listaDokumentow = listaDokumentow;
            komunikat = "";
            this.ControlBox = false;
        }

        private async void RozliczanieKaucji_Load(object sender, EventArgs e)
        {
            progressBar1.Value = 10;
            await UtworzKaucje(new Progress<Raport>(ReportProgress));
        }

        private bool Zaloguj()
        {
            //i tak loguje na obecną sesje 
            XLLoginInfo_20182 login = new XLLoginInfo_20182();
            login.Wersja = wersjaApi;
            login.Winieta = -1;
            login.TrybWsadowy = 1;
            login.Baza = "nazwaBazy";
            login.OpeIdent = "Login";
            login.OpeHaslo = "Haslo";
            login.UtworzWlasnaSesje = 1;
            login.ProgramID = "IdProgramu";
            blad = cdn_api.cdn_api.XLLogin(login, ref sesjaId);
            if (blad == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void ReportProgress(Raport raport)
        {
            richTextBox1.SelectionColor = raport.sukces ? Color.Black : Color.Red;
            richTextBox1.AppendText(raport.komunikat + "\n");
            richTextBox1.ScrollToCaret();
            if (progressBar1.Value + raport.wartosc <= 100)
                progressBar1.Value += raport.wartosc;
        }


        async Task UtworzKaucje(IProgress<Raport> progress)
        {

            if (Zaloguj())
            {
                await Task.Delay(1);
                progress.Report(new Raport("Zalogowano", true, 10));

                int dzielnik = 50 / (listaDokumentow.Count());
                for (int i = 0; i < listaDokumentow.Count; i++)
                {
                    int gidNumer = int.Parse(listaDokumentow[i].GIDNumer);
                    int gidTyp = int.Parse(listaDokumentow[i].GIDTyp);
                    int gidFirma = int.Parse(listaDokumentow[i].GIDFirma);
                    int gidLp = int.Parse(listaDokumentow[i].GIDLp);

                    DataTable dataTable = new DataTable();//"Wczytane dane z bazy danych";

                    DataTable konTable = new DataTable();//"Wczytane dane z bazy danych";

                    XLDokumentNagInfo_20182 nowyDokument = new XLDokumentNagInfo_20182();
                    nowyDokument.Wersja = wersjaApi;
                    nowyDokument.ZwrNumer = gidNumer;
                    nowyDokument.ZwrTyp = gidTyp;
                    nowyDokument.ZwrFirma = gidFirma;
                    nowyDokument.ZwrLp = gidLp;
                    nowyDokument.Typ = 2008; //WKK
                    nowyDokument.Tryb = 2;
                    nowyDokument.Korekta = 1;
                    nowyDokument.FRSID = int.Parse(dataTable.Rows[0]["Trn_FrsId"].ToString());
                    nowyDokument.MagazynD = dataTable.Rows[0]["mag_kod"].ToString();
                    nowyDokument.KonNumer = int.Parse(konTable.Rows[0]["TrN_KonNumer"].ToString());
                    nowyDokument.KonTyp = int.Parse(konTable.Rows[0]["TrN_KonTyp"].ToString());
                    int dokumentId = 0;
                    int wynik = cdn_api.cdn_api.XLNowyDokument(sesjaId, ref dokumentId, nowyDokument);
                    if (wynik == 0)
                    {
                        await Task.Delay(1);
                        progress.Report(new Raport("Utworzono dokument ", true, dzielnik));
                        foreach (DataRow row in dataTable.Rows) //Dodawanie pozycji
                        {
                            XLDokumentElemInfo_20182 pozycja = new XLDokumentElemInfo_20182();
                            pozycja.Wersja = wersjaApi;
                            pozycja.Ilosc = row["Ilosc"].ToString();
                            pozycja.TowarKod = row["twr_kod"].ToString();
                            pozycja.TypKorekty = 1;
                            pozycja.GIDLpOrg = int.Parse(row["OrgidPozycji"].ToString());
                            wynik = cdn_api.cdn_api.XLDodajPozycje(dokumentId, pozycja);
                            if (wynik == 0)
                            {
                                await Task.Delay(1);
                                progress.Report(new Raport("Dodano pozycje!", true, 0));
                            }
                            else
                            {
                                await Task.Delay(1);
                                progress.Report(new Raport("Błąd dodawania pozycji " + KomunikatBledu(wynik, 2), false, dzielnik));
                            }
                        }

                        XLZamkniecieDokumentuInfo_20182 zamkniecie = new XLZamkniecieDokumentuInfo_20182();
                        zamkniecie.Wersja = wersjaApi;
                        zamkniecie.Tryb = 0;
                        wynik = cdn_api.cdn_api.XLZamknijDokument(dokumentId, zamkniecie);
                        if (wynik != 0)
                        {
                            await Task.Delay(1);
                            progress.Report(new Raport("Błąd podczas zamykania dokumentu " + KomunikatBledu(wynik, 7), false, dzielnik));
                        }
                    }
                    else
                    {
                        await Task.Delay(1);
                        progress.Report(new Raport("Pojawil się błąd podczas tworzenia dokumentu " + KomunikatBledu(wynik, 1), false, dzielnik));
                    }
                    await Task.Delay(1);
                    progress.Report(new Raport("Zakończono dodawanie pozycji", true, dzielnik));
                }
            }
            else
            {
                await Task.Delay(1);
                progress.Report(new Raport("Błąd logowania " + blad, false, 10));
            }
            DialogResult = DialogResult.OK;
            progressBar1.Value = 100;
            MessageBox.Show("Zakończono");
            this.Close();
        }

        private string KomunikatBledu(long nrBledu, int nrFunkcja)
        {
            XLKomunikatInfo_20182 komunikat = new XLKomunikatInfo_20182();
            komunikat.Wersja = wersjaApi;
            komunikat.Blad = (int)nrBledu;
            komunikat.Funkcja = nrFunkcja;
            int blad = cdn_api.cdn_api.XLOpisBledu(komunikat);
            if (blad == 0)
            {
                return komunikat.OpisBledu;
            }
            else
            {
                return "Funkcja: " + nrFunkcja + ", Błąd nr " + nrBledu;
            }
        }

        public class Raport
        {
            public string komunikat { get; set; }
            public bool sukces { get; set; }
            public int wartosc { get; set; }

            public Raport(string komunikat, bool blad, int wartosc)
            {
                this.komunikat = komunikat;
                this.wartosc = wartosc;
                this.sukces = blad;
            }
        }


    }
}
