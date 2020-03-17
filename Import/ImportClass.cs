using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CDNBase;

using System.Data.Common;
using System.Configuration;

namespace Import
{
    class ImportClass
    {
        public static bool czyIstniejeRaport(DateTime data, string Numer, Rejestr rejestr)
        {
            string provider = ConfigurationManager.AppSettings["provider"];
            string connectionString = ConfigurationManager.AppSettings["connectionString"];
            DbProviderFactory factory = DbProviderFactories.GetFactory(provider); //to pozwala na słanie zapytań do bazy danych
            using (DbConnection connection = factory.CreateConnection())
            {
                if (connection == null)
                {
                    Debug.WriteLine("Connection Error");
                    Console.ReadLine();
                    return false;
                }
                connection.ConnectionString = connectionString;
                connection.Open();
                DbCommand command = factory.CreateCommand();
                if (command == null)
                {
                    Debug.WriteLine("Command Error");
                    Console.ReadLine();
                    return false;
                }
                command.Connection = connection;
                command.CommandText = "Select * From [CDN_Firma_Demo].[CDN].[BnkRaporty]  "; // +" Where BRp_DataDok = "+ data;


                List<Raport> raporty = new List<Raport>();
                
                using (DbDataReader dataReader = command.ExecuteReader())
                {
                    DateTime dataS;
                    while (dataReader.Read())
                    {
                        //Debug.WriteLine(
                        //    $"{dataReader["BRp_NumerPelny"]} " +    //musi zawierać rejestr
                        //    $"{dataReader["BRp_DataDok"]} " +
                        //    $"{dataReader["BRp_DataZam"]}"
                        //    );
                        string dataOtw = dataReader["BRp_DataDok"].ToString();
                        dataS = DateTime.Parse(dataOtw.Substring(0, 10));

                        raporty.Add(new Raport
                        {
                            NazwaRejestru = dataReader["BRp_NumerPelny"].ToString(),
                            //DataOtwarciaString = dataReader["BRp_DataDok"].ToString(),
                            //DataZamknieciaString = dataReader["BRp_DataZam"].ToString(),
                            DataOtwarcia = dataS
                        });
                        
                    }
                
                } // zapełnienie listy wszystkich raportów z daty

                //foreach(Raport r in raporty) 
                //{
                //    Debug.WriteLine(r.NazwaRejestru);
                //    Debug.WriteLine(r.DataOtwarcia);
                  
                //} //Wypisanie wszystkich raportów

                //Filtrowanie listy raportówpo numerze rejestru
                var rap = from Raport in raporty
                          where Raport.NazwaRejestru.Contains(Numer) && Raport.DataOtwarcia == data
                          select Raport;
                if (rap.Any())
                {
                    //Debug.WriteLine("***Raport Istnieje***");
                    //foreach (Raport r in rap)
                    //{
                    //    Debug.WriteLine(r.NazwaRejestru);
                    //    Debug.WriteLine(r.DataOtwarcia);
                    //}
                    return true;
                } //return true    ...wypisanie przefiltrowanych raportów
                else
                {
                    Debug.WriteLine("Brak Raportu...");
                    //Debug.WriteLine("Tworzenie nowego raportu");
                    //try
                    //{
                        
                    //    NowyRaport(rejestr);
                       
                    //}
                    //catch (Exception e)
                    //{
                    //    Debug.WriteLine("Raport nie mógł zostać utworzony: " + e.Message);
                    //}

                    //tak robimy metode zeby zwrócilo true !!
                    //Miejsce na metode utworzenia nowego raportu !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    return false;
                }
                
            }
        }

        public static void ImportZPliku( Rejestr rejestr)
        {
            
            try
            {

                //tablica lini z pliku
                string[] lines = File.ReadAllLines(@"C:\Banki\"+rejestr.Numer+@"\"+rejestr.Data.Year+@"\"+rejestr.DataWpisana+@"\"+rejestr.PelnaNazwa);
                //string[] lines = File.ReadAllLines(@"C:\Banki\1572\2020\03.2020\1572_02032020");
                //ilość zapisów do wykonania
                int iloscZapisow = 0;
                int iloscLinii = 0;
                int dlugoscZapisu; 
            
                foreach (string line in lines)
                {
                    if (line.Contains("4:")) 
                    {
                        iloscZapisow++; 
                    }
                    iloscLinii++;
                }
                //Obliczenie długośći jednego zapisuKB
                dlugoscZapisu = iloscLinii / iloscZapisow;
                //tworzymy tablice obiektów
                Zapis[] zapisy = new Zapis[iloscZapisow];

                //Debug.WriteLine("Ilość lini: " + iloscZapisow);
           
                //int ktoryZapis = 0; //    do obiektów
                int aktualnaLinia = 0;
                int koniecPetli = dlugoscZapisu;
                int KtoryZapisOpis = 1;
                for (int j = 0; j < iloscZapisow; j++) //Petla zmieniająca zapis
                {
                    Zapis zapis = new Zapis();
                    //Debug.WriteLine("*******" + KtoryZapisOpis + "******");
                
                    for(int i = aktualnaLinia; i<koniecPetli; i++  ) //pętla poszczególnych linii
                    {
                        //Debug.WriteLine(lines[i]); // sprawdzanie
                        if (lines[i].Contains(":25:"))
                        {
                            //Debug.WriteLine("***"+lines[i].Substring(7,26));
                            zapis.Konto = lines[i].Substring(7, 26);
                        }
                        if (lines[i].Contains(":28:"))
                        {
                            //Debug.WriteLine("***"+lines[i].Substring(9));
                            zapis.Data = (DateTime.Parse(lines[i].Substring(9)));
                        }
                        if (lines[i].Contains(":60F:"))
                        {
                            //Debug.WriteLine("***" + lines[i].Substring(15));
                            if (lines[i].Substring(15) == "0,00")
                            {
                                //Debug.WriteLine("***" + lines[i + 1].Substring(15));
                                zapis.Wartosc = decimal.Parse(lines[i + 1].Substring(15));
                                zapis.Symbol = 1;
                            }
                            else
                            {
                                zapis.Wartosc = decimal.Parse(lines[i].Substring(15));
                                zapis.Symbol = -1;
                            }
                        }
                        if (lines[i].Contains("^20"))
                        {
                            //Debug.WriteLine("opis: "+lines[i].Substring(3));
                            zapis.Opis = lines[i].Substring(3);
                        }
                        zapisy[j] = zapis;
                    }
                    aktualnaLinia = aktualnaLinia + dlugoscZapisu;
                    koniecPetli = aktualnaLinia + dlugoscZapisu;
                    KtoryZapisOpis++;
                }
            

                Debug.WriteLine("OBIEKTY");
                foreach(Zapis zapis in zapisy)
                {
                    try { //Debug.WriteLine(zapis.Konto); 
                    }
                    catch(Exception e) { Debug.WriteLine(e.Message); }
                    try { //Debug.WriteLine(zapis.Data); 
                    }
                    catch(Exception e) { Debug.WriteLine(e.Message); }
                    try { //Debug.WriteLine(zapis.Wartosc); 
                    }
                    catch (Exception e) { Debug.WriteLine(e.Message); }
                    try { //Debug.WriteLine(zapis.Symbol); 
                    }
                    catch (Exception e) { Debug.WriteLine(e.Message); }
                    try { //Debug.WriteLine(zapis.Opis); 
                    }
                    catch (Exception e) { Debug.WriteLine(e.Message); }

                    //Zapis do Optimy!!!
                    ZapisKB(zapis, rejestr);
                }
            }
            catch(Exception e)
            {
                Debug.WriteLine("Plik nie został znalezniony" + e.Message);
            }
        }

        public static int ZwrocID(Rejestr rejestr)
        {
            int ID=1;
            //int NumerNr=0;

            string provider = ConfigurationManager.AppSettings["provider"];
            string connectionString = ConfigurationManager.AppSettings["connectionString"];
            DbProviderFactory factory = DbProviderFactories.GetFactory(provider); //to pozwala na słanie zapytań do bazy danych
            using (DbConnection connection = factory.CreateConnection())
            {
                if (connection == null)
                {
                    Debug.WriteLine("Connection Error");
                    Console.ReadLine();
                    return 0;
                }
                connection.ConnectionString = connectionString;
                connection.Open();
                DbCommand command = factory.CreateCommand();
                if (command == null)
                {
                    Debug.WriteLine("Command Error");
                    Console.ReadLine();
                    return 0;
                }
                command.Connection = connection;
                command.CommandText = "Select TOP(1) [BRp_BRaID],[BRp_NumerNr],[BRp_NumerString],[BRp_DataDok],[BRp_DataZam],[BRp_Zamkniety],[BRp_NumerPelny] From [CDN_Firma_Demo].[CDN].[BnkRaporty] where [BRp_NumerString] LIKE '%" + rejestr.Numer + "%' "; // +" Where BRp_DataDok = "+ data;

                using (DbDataReader dataReader = command.ExecuteReader())
                {
                    DateTime dataS;
                    while (dataReader.Read())
                    {
                        Object SqlId = dataReader["BRp_BRaID"];
                        int ParseID = (int)SqlId;
                        ID = ParseID;
                        //Object SqlNumerNr = dataReader["BRp_NumerNr"];
                    }
                }
                return ID;
            }
        }

        public static int ZwrocNumerNr(Rejestr rejestr)
        {
            int NumerNr=0;
            if (czyIstniejeRaport(rejestr.Data, rejestr.Nazwa, rejestr) == true)
            {
                string provider = ConfigurationManager.AppSettings["provider"];
                string connectionString = ConfigurationManager.AppSettings["connectionString"];
                DbProviderFactory factory = DbProviderFactories.GetFactory(provider); //to pozwala na słanie zapytań do bazy danych
                using (DbConnection connection = factory.CreateConnection())
                {
                    if (connection == null)
                    {
                        Debug.WriteLine("Connection Error");
                        Console.ReadLine();
                        return 0;
                    }
                    connection.ConnectionString = connectionString;
                    connection.Open();
                    DbCommand command = factory.CreateCommand();
                    if (command == null)
                    {
                        Debug.WriteLine("Command Error");
                        Console.ReadLine();
                        return 0;
                    }
                    command.Connection = connection;
                    command.CommandText = "Select TOP(1) [BRp_NumerNr],[BRp_NumerString],[BRp_DataDok],[BRp_DataZam]From [CDN_Firma_Demo].[CDN].[BnkRaporty] where [BRp_NumerString] LIKE '%" + rejestr.Numer + "%' AND [BRp_DataDok] = '" + rejestr.Data.ToString("yyyy-MM-dd")+" 00:00:00.000'";//.ToString("yyyy-mm-dd") + ""; // +" Where BRp_DataDok = "+ data;

                    

                    using (DbDataReader dataReader = command.ExecuteReader())
                    {
                        DateTime dataS;
                        while (dataReader.Read())
                        {
                            Object SqlNumerNr = dataReader["BRp_NumerNr"];
                            int ParseNumerNr = (int)SqlNumerNr;
                            NumerNr = ParseNumerNr;

                        }
                    }
                }
            }
            return NumerNr;
        }

        public static void NowyRaport(Rejestr rejestr)
        {
            var RachunekID = ZwrocID(rejestr);
            try
            {
            CDNBase.AdoSession oSession = OptimaCOM.oLogin.CreateSession();
            var rRaport = oSession.CreateObject("CDN.RaportyKB").AddNew();

                //OP_KASBOLib.RaportKB raport = (OP_KASBOLib.RaportKB)oSession.CreateObject("CDNHeal.Raport", null).Item("RejestrRej = "+rejestr.Nazwa).AddNew();
            var rRejestr = oSession.CreateObject("CDN.Rachunki").Item("BRa_BRaID=" + RachunekID); // zeby metoda działała, w bazie danych musi być przynajmniej jeden raport !!!
                rRaport.Rachunek = rRejestr;
            rRaport.DataOtw = rejestr.Data;
            rRaport.DataZam = rejestr.Data;

            //czyIstniejeRaport(rejestr.Data, rejestr.Numer, rejestr);

            oSession.Save();
                Debug.WriteLine("Dodawanie Raportu zakończone powodzeniem!"); //Wcześniejszy raport nie ma ustalonej daty zamknięcia.
            }
            catch(Exception e)
            {
                Debug.WriteLine("Błąd tworzenia sesji podczas dodawania RAPORTU: " + e.Message); //Błąd tworzenia sesji podczas dodawania RAPORTUWcześniejszy raport nie ma ustalonej daty zamknięcia.
            }
        }

        public static void ZapisKB( Zapis zapis, Rejestr rejestr)
        {
            try
            {
                CDNBase.AdoSession oSession = OptimaCOM.oLogin.CreateSession();
                
                var rZapis = oSession.CreateObject("CDN.ZapisyKB").AddNew();
                var rNumerator = rZapis.Numerator; //

                try
                {
                    var rDokDef = oSession.CreateObject("CDN.DefinicjeDokumentow").Item("DDf_DDfID = 5");//.Item("DDf_Symbol=''"); //seria
                    rNumerator.DefinicjaDokumentu = rDokDef;

                    Debug.WriteLine("Definicja dokumentu UDANA!");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Błąd podczas tworzenia definicji dokumentu: "+e.Message);
                } //DefinicjeDokumentów

                try
                {
                var rRachunek = oSession.CreateObject("CDN.Rachunki").Item("Bra_Akronim = '" + rejestr.Nazwa+ "'"); //Rejestr
                rZapis.Rachunek = rRachunek;
                    Debug.WriteLine("pomyślne ustawienie rachunku");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Błąd podczas ustawiania rachunku" + Environment.NewLine+e.Message);
                } // Ustawienie rachunku
               
                try 
                {
                    //Numer raportu... jak to rozegrać...? 1/2020 to jest numer otwartego raportu w kolejnośći tworzenia.. to chyba powinna być zmienna
                    //DateTime data = rejestr.Data.AddDays(-1);

                    if (czyIstniejeRaport(rejestr.Data, rejestr.Numer, rejestr) == false)
                        {
                            NowyRaport(rejestr);
                        Debug.WriteLine("NowyRaport() ZapisKB()......................................................................");
                        };


                        int NumerNr = ZwrocNumerNr(rejestr); //moge to wrzucic jako parametr do funkcji... int wywolasc gdzies przed wywolanie tej funkcji
                        Debug.WriteLine("NumerNr() ZapisKB()......................................................NIEPOTRZEBNE ZAPYTANIA DO BAZY...");
                    

                    var rRaport = oSession.CreateObject("CDN.RaportyKB").Item("BRp_NumerPelny = '" + "RKB/"+NumerNr+"/"+rejestr.Data.Year+"/"+rejestr.Nazwa + "'");
                    //Debug.WriteLine("pomyślne dodanie Numeru");
                    rZapis.RaportKB = rRaport;
                      
                    rZapis.DataDok = rejestr.Data; //z ta data byl problem !!!!!
                    rZapis.Kwota = zapis.Wartosc;
                    //rZapis.NumerObcy = zapis.Konto;
                    rZapis.Opis = zapis.Opis;
   
                   // Debug.WriteLine("pomyślne dodanie danych bloku: RaportKB, DataDok, Kwota");
                   // var rSeria = oSession.CreateObject("OP_KASBOLib.ZapisKB").Item("Seria1 = KASA");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Błąd podczas ustawiania raportu: "+e.Message);
                } //RaportKB, Data, Kwota, NUMER OBCY !!! 

                rZapis.DefinicjaDokumentu = rNumerator.DefinicjaDokumentu;
                rZapis.Kierunek = zapis.Symbol; //nie wiem czemu 2
                                     //rZapis.Seria = "KASA";
                
                //OP_KASBOLib.ZapisKB zapis = oSession.CreateObject(OP_KASBOLib.ZapisKB)
                
                try
                {
                    var rKontrahent = oSession.CreateObject("CDN.Kontrahenci").Item("Knt_KOD = '" +"!NIEOKREŚLONY!"+"'");
                    rZapis.Podmiot = rKontrahent;
                    //Debug.WriteLine("Kontrahent: POMYŚLNIE");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Kontrahent NIE POMYŚLNIE... "+ e.Message);
                } //Kontrahent

                oSession.Save();
                //Debug.WriteLine("sesja udana");
            }
            catch(Exception e)
            {
                Debug.WriteLine("Błąd tworzenia sesji: "+e.Message);
            }
        } 

        public static List<Rejestr> SzukaniePlikow() //zwraca pelna liste rejestrów... lacznie z tymi w których nie ma pliku
        {
            string[] directories = Directory.GetDirectories(@"C:\Banki", ".", SearchOption.AllDirectories);

            List<string> rejestry = new List<string>(); //Lista rejestrów 1234
            List<string> rejestrIPlik = new List<string>(); //Lista pelnych nazw plików 1234_02032020
            List<Rejestr> ListaRejestrow = new List<Rejestr>();
            List<Rejestr> ListaRejestrowOdfiltrowana = new List<Rejestr>();
            string Path = ""; //nazwa pliku bez rejestru

            //foreach (string d in directories)
            //{
            //    if (d.Contains("."))
            //    {
            //        Debug.WriteLine(d);
            //    }
            //} //wypisanie katalogów z datami miesiac/rok

            string data;    //Podajemy Datę
            Console.WriteLine("Podaj Datę:");
            data = Console.ReadLine();

            foreach (string d in directories)   //zapisuje rejestry które posiadają folder o podanej dacie MM.yyyy
            {
                if (d.Contains(data.Substring(3)))
                {
                    Debug.WriteLine(d);
                    rejestry.Add(d.Substring(9, 4));
                }
            } // jeżeli w spisice katalogów zawiera się takalog o dacie miesiac/rok ( podanej w formie DD:MM:YYYY) do listy rejestrów dodany jest rejestr

            //mechanizm przerabiania daty na ścieżke pliku
            Char[] charsToRemove = { '.' };
            string[] pth = data.Split(charsToRemove);
            for (int i = 0; i < pth.Length; i++)
            {
                Path = Path + pth[i];
            } // tworzenie zmiennej Path, czyli data bez kropek

            foreach (string s in rejestry)
            {
                rejestrIPlik.Add(s + "_" + Path);
            } // tworzenie pełnej listy pełnych nazw plików

            /* foreach (string s in rejestrIPlik)
             {
                 Debug.WriteLine("Ścieżka:" + s);
             } //wypisanie pełnych nazw */


            for (int i = 0; i < rejestrIPlik.Count; i++) //zrobić filtrowanie rejetrów do tych które zawierają plik !!!
            {

                ListaRejestrow.Add(new Rejestr
                {
                    PelnaNazwa = rejestrIPlik[i],
                    Nazwa = rejestry[i],
                    Numer = rejestry[i],
                    Data = DateTime.Parse(data), // dodawanie daty do obiektu Rejestr. Data jest dodawana na podstawie nazwy pliku
                    DataWpisana = data.Substring(3, 7)
                });
            }

            //foreach(Rejestr r in ListaRejestrow) // filtrowanie aktywnych rejestrów po tym czy plik istnieje
            //{
            //   // string path = @"‪C:\Banki\"+r.Nazwa+@"\"+r.Data.Year+ @"\"+r.Data.ToString("MM.yyyy")+@"\"+r.PelnaNazwa+@"";
            //    string path = @"‪C:\Banki\1572\2020\03.2020\1572_07032020";
            //    if (File.Exists(path))
            //    {
            //        ListaRejestrowOdfiltrowana.Add(r);
            //    }
            //}
            //foreach (Rejestr r in ListaRejestrow) //Wypisanie Listy Rejestrów
            //{
            //    Debug.WriteLine("***Rejestr***");
            //    Debug.WriteLine("Pelna nazwa: " + r.PelnaNazwa);
            //    Debug.WriteLine("Nazwa: " + r.Nazwa);
            //    Debug.WriteLine("Numer: " + r.Numer);
            //    Debug.WriteLine("Data: " + r.Data.ToString());
            //    Debug.WriteLine("Data wpisana: " + r.DataWpisana);
            //}


            return ListaRejestrow;
        } // funkcja zwraca listę rejestrów
    }
}
