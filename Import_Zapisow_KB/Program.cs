using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CDNBase;
using System.Runtime.InteropServices;


namespace Import_Zapisow_KB
{
    class Program
    {
		protected static IApplication Application = null;
		protected static ILogin Login = null;

		#region Logowanie i wylogowanie
		// Przyklad 1. - Logowanie do O! bez wyświetlania okienka logowania
		static protected void LogowanieAutomatyczne()
		{
			string Operator = "ADMIN";
			string Haslo = "";      // operator nie ma hasła
			string Firma = "Firma_Demo";  // nazwa firmy


			object[] hPar = new object[] {
						 1,  1,   0,  0,  1,   1,  0,    0,   0,   0,   0,   0,   1,   1,  1,   0,  0, 0 }; // do jakich modułów się logujemy
			/* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP, CRMP
			 */

			// katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
			//System.Environment.CurrentDirectory = @"C:\Gotowe.Optima.2015.0.1";//@"C:\Program Files\OPTIMA.NET";	
			System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";//@"C:\Program Files\OPTIMA.NET";

			// tworzymy nowy obiekt apliakcji
			Application = new Application();

			// Jeśli proces nie ma dostępu do klucza w rejstrze 
			// HKCU\Software\CDN\CDN OPT!MA\CDN OPT!MA\Login\KonfigConnectStr
			// np. gdy pracuje jako aplikacji w IIS 
			// ciąg połączeniowy (ConnectString) podajemy bezpośrednio :
			// Application.KonfigConnectStr = "NET:CDN_KNF_Konfiguracja_DW,SERWERSQL,NT=0";


			// blokujemy
			Login = Application.LockApp(256, 5000, null, null, null, null);
			//Login =  Application.LockApp(1, 5000, null, null, null, null);


			// logujemy się do podanej Firmy, na danego operatora, do podanych modułów
			Login = Application.Login(Operator, Haslo, Firma, hPar[0], hPar[1], hPar[2], hPar[3], hPar[4], hPar[5], hPar[6], hPar[7], hPar[8], hPar[9], hPar[10], hPar[11], hPar[12], hPar[13], hPar[14], hPar[15], hPar[16], hPar[17]);


			//  Logowanie z pobraniem ustawienia modułów z karty Operatora
			//	Login = Application.Login(Operator, Haslo, Firma, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
			//

			// tu jesteśmy zalogowani do O!
			Console.WriteLine("Jesteśmy zalogowani do O!");
			Console.ReadLine();
			
		}

		protected static void Wylogowanie()
		{
			// niszczymy Login
			Login = null;
			// odblokowanie (wylogowanie) O!
			Application.UnlockApp();
			// niszczymy obiekt Aplikacji
			Application = null;
		}
        #endregion

        #region Dodaj Zapis KB

		protected static void OdczytZPliku()
		{

		}
        #endregion
        //protected static void DodajZapisKB() 
        //{
        //	//Tworzymy obiekt sesji
        //	AdoSession Sesja = Login.CreateSession();

        //	var zapis = Sesja.CreateObject("CDN.ZapisyKB").AddNew();
        //	var rNumerator = zapis.Numerator;
        //	try
        //	{
        //		//zmienne na czerwono to zmienne z funkcji parsującej dane z pliku
        //		var DokDef = Sesja.CreateObject("CDN.DefinicjeDokumentów").Item("DDf_Symbol='" + Seria + "'");
        //		rNumerator.DefinicjaDokumentu = DokDef;
        //	}
        //	catch (Exception e) { Console.WriteLine("blad"); }

        //	var rRachunek = Sesja.CreateObject("CDN.Rachunki").Item("Bra_Akronim = '" + Rachunek + "'");
        //	zapis.Rachunek = rRachunek;
        //	rNumerator.Rejestr = rRachunek.Symbol;

        //	var rRaport = Sesja.CreateObject("CDN.RaportyKB").Item("BRp_NumerPelny = '" + Raport + "'");
        //	zapis.RaportKB = rRaport;
        //	zapis.DataDok = Data;
        //	zapis.Kwota = Wartosc;
        //	zapis.Kierunek = Kier;

        //	var Kontrahent = Sesja.CreateObject("CDN.Kontrahenci").Item("Knt_KOD ='" + Knt + "'");
        //	zapis.Podmiot = Kontrahent;
        //	try 
        //	{ 
        //		Sesja.Save(); 
        //	}
        //	catch (Exception e)
        //	{
        //		Console.WriteLine("Błąd podczas zapisu " + e.Message);
        //	}
        //}
        //      #endregion

        [STAThread]
        static void Main(string[] args)
        {
			// tu będziemy sobie wywoływać nasze przykłady.
			try
			{
				LogowanieAutomatyczne();
				//	DodanieRejestru();
				//	DodajKontrahenta();

				//DodanieTowaru();
				//DodanieFaktury();
				//        DodanieDokumentuOBD();
				//        DodanieZleceniaSerwisowego();            

				//	DogenerujKorekteFA();
				//	Dogenerowanie_WZ_Do_FA();
				//  UtworzeniePK();
			}
			catch (COMException e)
			{
				if (e.InnerException != null)
					System.Console.WriteLine("Błąd: " + e.InnerException.Message);
				else
					System.Console.WriteLine("Błąd: " + e.Message);
			}
			catch (Exception e)
			{
				if (e.InnerException != null)
					System.Console.WriteLine("Błąd: " + e.InnerException.Message);
				else
					System.Console.WriteLine("Błąd: " + e.Message);
			}





			finally
			{
				Wylogowanie();
			}





			Console.ReadLine();

		}
    }
}
