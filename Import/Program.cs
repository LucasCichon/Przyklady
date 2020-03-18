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
        public static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {

            OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");

            List<Rejestr> ListaRejestrow = ImportClass.SzukaniePlikow();

            //int NumerNr;
            for(int i=0; i<ListaRejestrow.Count; i++)
            {
                int NumerNr = ImportClass.ZwrocNumerNr(ListaRejestrow[i]);

                    if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                        {
                            ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
                        }
                    else
                    {
                        ImportClass.NowyRaport(ListaRejestrow[i]);
                        if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                            {
                                ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
                            }
                    }
            }
            OptimaCOM.O_Logout();
        }
    }
}
