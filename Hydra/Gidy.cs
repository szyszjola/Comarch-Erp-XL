using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hydra
{
    public class Gidy
    {
        public string GIDTyp, GIDFirma, GIDNumer, GIDLp;

        public Gidy()
        {
            GIDTyp = "";
            GIDFirma = "";
            GIDNumer = "";
            GIDLp = "";
        }

        override public string ToString()
        {
            return "Typ:" + GIDTyp + ", Firma:" + GIDFirma + ", Numer:" + GIDNumer + ", LP:" + GIDLp;
        }
    }
}
