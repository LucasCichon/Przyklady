using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Diagnostics;

using System.Data.Common;
using System.Configuration;

namespace Import
{
    class Program
    {
        static void Main(string[] args)
        {
            OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");
            Debug.WriteLine("zalogowano");
            List<Rejestr> ListaRejestrow = ImportClass.SzukaniePlikow();


            foreach (Rejestr r in ListaRejestrow)
            {
                if( ImportClass.czyIstniejeRaport(r.Data, r.Numer) == true)
                {
                ImportClass.ImportZPliku(r);
                }
            }

            OptimaCOM.O_Logout();
        }
    }
}
