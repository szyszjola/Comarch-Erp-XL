using ADODB;
using Hydra;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

[assembly: CallbackAssemblyDescription("NazwaDodatku", "Opis dodatku", "Twórca", "Wersja", "Zalecana wersja erp xl", "Data")]
namespace HydraRozliczKaucjeWKA
{
    [SubscribeProcedure((Procedures)Procedures.ListaDokumentow, "NazwaDodatku")]
    public class Class1 : Callback
    {
        ClaWindow oknoGlowne, przycisk;

        public override void Cleanup()
        {
        }

        public override void Init()
        {
            AddSubscription(true, 0, Events.OpenWindow, new TakeEventDelegate(zaladujPrzycisk));
        }

        private bool zaladujPrzycisk(Procedures ProcedureId, int ControlId, Events Event)
        {
            oknoGlowne = GetWindow();
            przycisk = oknoGlowne.Children["?TabKaucje"].Children.Add(ControlTypes.button);
            przycisk.Visible = true;
            przycisk.TextRaw = "Rozlicz";
            oknoGlowne.OnBeforeSized += oknoProgramu_OnBeforeSized;
            przycisk.OnBeforeAccepted += przycisk_OnBeforeAccepted;
            return true;
        }

        private bool przycisk_OnBeforeAccepted(Procedures ProcedureId, int ControlId, Events Event)
        {
            List<Gidy> listaGid = listaWybranychGid("?Handlowe", Procedures.ListaDokumentow);
            if (listaGid.Count > 0)
            {

                if (listaGid.All(x => x.GIDTyp.Equals("2000")))
                {
                    #region forms
                    RozliczanieKaucji rozliczanieKaucji = new RozliczanieKaucji((listaGid));
                    rozliczanieKaucji.ShowDialog();
                    #endregion

                    if (rozliczanieKaucji.DialogResult == DialogResult.OK)
                    {
                        Runtime.WindowController.PostEvent(0, Events.FullRefresh);
                    }
                }
                else
                {
                    MessageBox.Show("Wybrane dokumenty nie są WKA!");
                }
            }
            else
            {
                MessageBox.Show("Nie wybrano dokumentów!");
            }
            return true;
        }

        private bool oknoProgramu_OnBeforeSized(Procedures ProcedureId, int ControlId, Events Event)
        {
            System.Drawing.Rectangle oknoOrg = GetWindow().Bounds;
            przycisk.Bounds = new System.Drawing.Rectangle(180, oknoOrg.Height - 20, 40, 13);
            return true;
        }

        private List<Gidy> listaWybranychGid(string nazwaListy, Procedures procId)
        {
            List<Gidy> listaGID = new List<Gidy>();

            int listaId = GetWindow().AllChildren[nazwaListy].Id;
            _Recordset recordset = Runtime.WindowController.GetQueueMarked((int)procId, listaId, GetCallbackThread());
            try
            {
                //jesli nie jest nic zaznaczone to recordset == null
                if (recordset != null && recordset.RecordCount > 0)
                {
                    string fieldName;
                    recordset.MoveFirst();
                    while (recordset.EOF == false)
                    {
                        ADODB.Fields fields = recordset.Fields;
                        Gidy g = new Gidy();

                        for (int i = 0; i < fields.Count; i++)
                        {
                            fieldName = fields[i].Name;
                            if (fieldName == "TYP")
                            {
                                g.GIDTyp = fields[i].Value.ToString();
                            }

                            if (fieldName == "FIRMA")
                            {
                                g.GIDFirma = fields[i].Value.ToString();
                            }

                            if (fieldName == "NUMER")
                            {
                                g.GIDNumer = fields[i].Value.ToString();
                            }

                            if (fieldName == "LP")
                            {
                                g.GIDLp = fields[i].Value.ToString();
                            }
                        }

                        listaGID.Add(g);
                        recordset.MoveNext();
                    }
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return listaGID;
        }


    }
}
