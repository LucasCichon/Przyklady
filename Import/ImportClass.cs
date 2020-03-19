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
using System.Windows.Forms;




namespace Import
{
     class ImportClass : Form1
    {
        //Program.logUri("takie tam");
        //public Form1 myForm = new Form1();
        public Form1 form;
        public ImportClass(Form1 form) {
            this.form = form;
        }

        public static List<Rejestr> SzukaniePlikow(string Data) 
        {
            string[] directories = Directory.GetDirectories(@"C:\Banki", ".", SearchOption.AllDirectories);
            string[] directoriesRejestr = Directory.GetDirectories(@"C:\Banki", ".", SearchOption.TopDirectoryOnly);

            List<string> rejestry = new List<string>(); //Lista rejestrów 1234
            List<string> rejestrIPlik = new List<string>(); //Lista pelnych nazw plików 1234_02032020
            List<Rejestr> ListaRejestrow = new List<Rejestr>();
            //List<Rejestr> ListaRejestrowOdfiltrowana = new List<Rejestr>();
            string Path = ""; //nazwa pliku bez rejestru, czyli data bez kropek
            
            foreach(string s in directoriesRejestr)
            {
                //Console.WriteLine(s);
                rejestry.Add(s.Substring(9, 4));
            }
           
            string data;    //Podajemy Datę
            Console.WriteLine("Podaj Datę:"); // miejsce na zczytanie daty z okienka
            data = Data;
            //data = Console.ReadLine();

            #region stara metoda wyszukiwania rejestrów
            //try
            //{
            //    if (data.Length == 10)
            //    {
            //        foreach (string d in directories)   //zapisuje rejestry które posiadają folder o podanej dacie MM.yyyy
            //        {
            //            if (d.Contains(data.Substring(3)))
            //            {
            //                Debug.WriteLine(d);
            //                rejestry.Add(d.Substring(9, 4));
            //            }
                        
            //        } // jeżeli w spisice katalogów zawiera się takalog o dacie "miesiac.rok" ( podanej w formie MM:YYYY) do listy rejestrów dodany jest rejestr

            //    }
            //    else
            //    {
            //        Console.WriteLine("zły format daty. (dd.MM.yyyy)");
            //        SzukaniePlikow();
            //    }
            //}
            //catch(Exception e) 
            //    { 
            //        Debug.WriteLine("Błąd na poziomie zapełniania listy rejestró które posiadają folder o podanej dacie MM.yyyy" + e.Message); 
            //    }

            //if (!rejestry.Any())
            //    {
            //        Console.WriteLine("Błędne Dane, bądź plik nie istnieje");
            //        SzukaniePlikow();
            //    }
            #endregion  // stara metoda wyszukiwania rejestró //stara metoda wyszukiwania rejestrów stara metoda przerabiania 

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
        } // funkcja zwraca listę rejestrów ...  lacznie z tymi w których nie ma pliku ...UWAGA!!! rejestry są brane jeżeli folder (MM.yyyy) istnieje.

        public static void ImportZPliku(int NumerNr, Rejestr rejestr)
        {
            int numerNr = NumerNr;
            
            try
            {
                //tablica lini z pliku
                string[] lines = File.ReadAllLines(Properties.Settings.Default.Katalog_z_wyciagami+rejestr.Numer+@"\"+rejestr.Data.Year+@"\"+rejestr.DataWpisana+@"\"+rejestr.PelnaNazwa);
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
                          
                            //zapis.Opis = lines[i].Substring(3);
                            //string s1 = EncodingDoWindows1250(lines[i].Substring(3));
                            //string s2 = ZamianaPolskiegoZnaku(s1);
                            zapis.Opis = ZamianaPolskiegoZnaku(lines[i].Substring(3));
                        }
                        zapisy[j] = zapis;
                    }
                    aktualnaLinia = aktualnaLinia + dlugoscZapisu;
                    koniecPetli = aktualnaLinia + dlugoscZapisu;
                    KtoryZapisOpis++;
                }

                try  
                {
                    foreach (Zapis zapis in zapisy)
                    {
                        try
                        { //Debug.WriteLine(zapis.Konto); 
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }
                        try
                        { //Debug.WriteLine(zapis.Data); 
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }
                        try
                        { //Debug.WriteLine(zapis.Wartosc); 
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }
                        try
                        { //Debug.WriteLine(zapis.Symbol); 
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }
                        try
                        { //Debug.WriteLine(zapis.Opis); 
                        }
                        catch (Exception e) { Debug.WriteLine(e.Message); }

                        //Zapis do Optimy!!!
                        ZapisKB(numerNr, zapis, rejestr);

                    } 
                    Form1.log.Debug("Poprawny zapis rejestrów do Raportu: RKB/" + NumerNr + "/" + rejestr.Data.Year.ToString() + "/" + rejestr.Numer);
                    Debug.WriteLine("Poprawny zapis rejestrów do Raportu: RKB/" + NumerNr + "/" + rejestr.Data.Year.ToString() + "/" + rejestr.Numer);
                }
                catch(Exception e)
                {
                    Form1.log.Error("Niepoprawny zapis rejestrów do Raportu: RKB/"+NumerNr+"/"+rejestr.Data.Year.ToString()+"/"+rejestr.Numer+", "+e.Message  );
                    Debug.WriteLine("Niepoprawny zapis rejestrów do Raportu: RKB/"+NumerNr+"/"+rejestr.Data.Year.ToString()+"/"+rejestr.Numer+", "+e.Message  );
                } // zapisywanie zapisówKB
            }
            catch(Exception e)
            {
                Form1.log.Error("Plik nie został znalezniony" + e.Message);
                Debug.WriteLine("Plik nie został znalezniony" + e.Message);
            }
        } //Parsowanie danych z pliku. 

        public static bool czyIstniejeRaport(DateTime data, string Numer, Rejestr rejestr)
        {
            string provider = ConfigurationManager.AppSettings["provider"];
            string connectionString = ConfigurationManager.AppSettings["connectionString"];
            DbProviderFactory factory = DbProviderFactories.GetFactory(provider); //to pozwala na słanie zapytań do bazy danych
            using (DbConnection connection = factory.CreateConnection())
            {
                if (connection == null)
                {
                    Form1.log.Error("Connection Error");
                    Debug.WriteLine("Connection Error");
                    Console.ReadLine();
                    return false;
                }
                connection.ConnectionString = connectionString;
                connection.Open();
                DbCommand command = factory.CreateCommand();
                if (command == null)
                {
                    Form1.log.Error("Commant Error");
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

                //Filtrowanie listy raportówpo numerze rejestru i Dacie otwarcia
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
                } //return true    
                else
                {
                    Debug.WriteLine("Brak Raportu. Raport jest tworzony.");
                    Form1.log.Debug("Brak Raportu.");
                    try
                    {
                        NowyRaport(rejestr); //jeżeli nie ma raportu, raport jest tworzony od razu i zwracana jest wartosć true
                        return true;
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine("Raport nie mógł zostać utworzony: " + e.Message);
                        Form1.log.Error("Raport nie mógł zostać utworzony: " + e.Message);
                    }

                    return false;
                    //tak robimy metode zeby zwrócilo true !!
                    //Miejsce na metode utworzenia nowego raportu !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                } //tworzenie nowego raportu - return true / catch exeption return false
                
            }
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
                Form1.log.Debug("Dodawanie Raportu zakończone powodzeniem!"); //Wcześniejszy raport nie ma ustalonej daty zamknięcia.
                Debug.WriteLine("Dodawanie Raportu zakończone powodzeniem!"); //Wcześniejszy raport nie ma ustalonej daty zamknięcia.
            }
            catch(Exception e)
            {
                Form1.log.Error("Błąd tworzenia sesji podczas dodawania RAPORTU: "+ Environment.NewLine + e.Message); //Błąd tworzenia sesji podczas dodawania RAPORTUWcześniejszy raport nie ma ustalonej daty zamknięcia.
                Debug.WriteLine("Błąd tworzenia sesji podczas dodawania RAPORTU: "+ Environment.NewLine + e.Message); 
            }
        } // Tworzenie nowego raportu

        public static void ZapisKB(int numerNr,Zapis zapis, Rejestr rejestr)
        {
            int NumerNr = numerNr;
            try
            {
                CDNBase.AdoSession oSession = OptimaCOM.oLogin.CreateSession();

                var rZapis = oSession.CreateObject("CDN.ZapisyKB").AddNew();
                var rNumerator = rZapis.Numerator; //

                try
                {
                    //.Item("CDN_PRINTERCODEPAGE = 1250")
                    var rDokDef = oSession.CreateObject("CDN.DefinicjeDokumentow").Item("DDf_DDfID = 5");//.Item("DDf_Symbol=''"); //seria
                    rNumerator.DefinicjaDokumentu = rDokDef;
                    
                    //Debug.WriteLine("Definicja dokumentu UDANA!");
                }
                catch (Exception e)
                {
                    Form1.log.Debug("Błąd podczas tworzenia definicji dokumentu: " + e.Message);
                } //DefinicjeDokumentów

                try
                {
                    var rRachunek = oSession.CreateObject("CDN.Rachunki").Item("Bra_Symbol = '" + rejestr.Nazwa + "'"); //Rejestr//musi byc BRa_Symbol bo Akronimy mogą mieć inną długość
                    rZapis.Rachunek = rRachunek;
                    //Debug.WriteLine("pomyślne ustawienie rachunku");
                }
                catch (Exception e)
                {
                    Form1.log.Debug("Błąd podczas ustawiania rachunku: " + Environment.NewLine + e.Message);
                } // Ustawienie rachunku

                try
                {
                    if (czyIstniejeRaport(rejestr.Data, rejestr.Numer, rejestr) == false)
                    {
                        NowyRaport(rejestr);
                        ZapisKB(NumerNr, zapis, rejestr);
                        Debug.WriteLine("NowyRaport() ZapisKB()......................................................................");
                    };

                    //int NumerNr = ZwrocNumerNr(rejestr); //moge to wrzucic jako parametr do funkcji... int wywolasc gdzies przed wywolanie tej funkcji
                    //Debug.WriteLine("NumerNr() ZapisKB()......................................................NIEPOTRZEBNE ZAPYTANIA DO BAZY...");
                    var rRaport = oSession.CreateObject("CDN.RaportyKB").Item("BRp_NumerPelny = '" + "RKB/" + NumerNr + "/" + rejestr.Data.Year + "/" + rejestr.Nazwa + "'");
                    //Debug.WriteLine("pomyślne dodanie Numeru");
                    rZapis.RaportKB = rRaport;

                    rZapis.DataDok = rejestr.Data; 
                    rZapis.Kwota = zapis.Wartosc;
                    //rZapis.NumerObcy = zapis.Konto;
                    rZapis.Opis = zapis.Opis;

                    // Debug.WriteLine("pomyślne dodanie danych bloku: RaportKB, DataDok, Kwota");
                }
                catch (Exception e)
                {
                    Debug.WriteLine("Błąd podczas ustawiania raportu: " + e.Message);
                } //RaportKB, Data, Kwota, NUMER OBCY !!! 

                rZapis.DefinicjaDokumentu = rNumerator.DefinicjaDokumentu;
                rZapis.Kierunek = zapis.Symbol; 
                
                try
                {
                    var rKontrahent = oSession.CreateObject("CDN.Kontrahenci").Item("Knt_KOD = '" + "!NIEOKREŚLONY!" + "'");
                    rZapis.Podmiot = rKontrahent;
                    //Debug.WriteLine("Kontrahent: POMYŚLNIE");
                }
                catch (Exception e)
                {
                    Debug.WriteLine("Kontrahent nie dodany pomyślnie: " + e.Message);
                } //Kontrahent

                oSession.Save();
                //Debug.WriteLine("sesja udana");
            }
            catch (Exception e)
            {
                Debug.WriteLine("Błąd tworzenia sesji: " + e.Message);
            }
        }  //tworzenie zapisu KB ( bez niepotrzebnych zapytań do bazy)

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
                        Form1.log.Debug("Connection Error");
                        Console.ReadLine();
                        return 0;
                    }
                    connection.ConnectionString = connectionString;
                    connection.Open();
                    DbCommand command = factory.CreateCommand();
                    if (command == null)
                    {
                        Form1.log.Debug("Connection Error");
                        Console.ReadLine();
                        return 0;
                    }
                    command.Connection = connection;
                    command.CommandText = "Select TOP(1) [BRp_NumerNr],[BRp_NumerString],[BRp_DataDok],[BRp_DataZam]From [CDN_Firma_Demo].[CDN].[BnkRaporty] where [BRp_NumerString] LIKE '%" + rejestr.Numer + "%' AND [BRp_DataDok] = '" + rejestr.Data.ToString("yyyy-MM-dd")+" 00:00:00.000'";//.ToString("yyyy-mm-dd") + ""; // +" Where BRp_DataDok = "+ data;

                    

                    using (DbDataReader dataReader = command.ExecuteReader())
                    {
                        while (dataReader.Read())
                        {
                            Object SqlNumerNr = dataReader["BRp_NumerNr"];
                            int ParseNumerNr = (int)SqlNumerNr;
                            NumerNr = ParseNumerNr;

                        }
                    }
                }
            }
            else
            {
                NowyRaport(rejestr);
                ZwrocNumerNr(rejestr);
            }
            return NumerNr;
        } // funkcja zwrada numer raportu bankowego

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
        } //funkcja zwtara id rejestru, potrzebne do utworzenia nowego raportu

        public static string ZamianaPolskiegoZnaku(string s)
        {
            string getStr = s;
            string returnStr="";
            for(int i=0; i<getStr.Length; i++)
            {
                if(getStr[i]== '�'||getStr[i]=='?')
                {
                    //returnStr.Remove(i, 1);
                   // returnStr.Insert(i, "Ł");
                    returnStr = returnStr + "Ł";
                }
                else
                {
                   // returnStr.Insert(i, getStr[i].ToString());
                    returnStr = returnStr + getStr[i];
                }
            }
            //Console.WriteLine(returnStr); //wypisanie zmienionego opisu

            return returnStr;
        }

        //#region nieużywane metody
        //public static void ZapisKB( Zapis zapis, Rejestr rejestr)
        //{
        //    try
        //    {
        //        CDNBase.AdoSession oSession = OptimaCOM.oLogin.CreateSession();
                
        //        var rZapis = oSession.CreateObject("CDN.ZapisyKB").AddNew();
        //        var rNumerator = rZapis.Numerator; //

        //        try
        //        {
        //            var rDokDef = oSession.CreateObject("CDN.DefinicjeDokumentow").Item("DDf_DDfID = 5");//.Item("DDf_Symbol=''"); //seria
        //            rNumerator.DefinicjaDokumentu = rDokDef;

        //            //Debug.WriteLine("Definicja dokumentu UDANA!");
        //        }
        //        catch(Exception e)
        //        {
        //            Program.log.Debug("Błąd podczas tworzenia definicji dokumentu: "+e.Message);
        //        } //DefinicjeDokumentów

        //        try
        //        {
        //        var rRachunek = oSession.CreateObject("CDN.Rachunki").Item("Bra_Symbol= '" + rejestr.Nazwa+ "'"); //Rejestr
        //        rZapis.Rachunek = rRachunek;
        //            //Debug.WriteLine("pomyślne ustawienie rachunku");
        //        }
        //        catch(Exception e)
        //        {
        //            Program.log.Debug("Błąd podczas ustawiania rachunku" + Environment.NewLine+e.Message);
        //        } // Ustawienie rachunku
               
        //        try 
        //        {
        //            if (czyIstniejeRaport(rejestr.Data, rejestr.Numer, rejestr) == false)
        //                {
        //                    NowyRaport(rejestr);
        //                Debug.WriteLine("NowyRaport() ZapisKB()......................................................................");
        //                };


        //                int NumerNr = ZwrocNumerNr(rejestr); //moge to wrzucic jako parametr do funkcji... int wywolasc gdzies przed wywolanie tej funkcji
        //                Debug.WriteLine("NumerNr() ZapisKB()......................................................NIEPOTRZEBNE ZAPYTANIA DO BAZY...");
                    

        //            var rRaport = oSession.CreateObject("CDN.RaportyKB").Item("BRp_NumerPelny = '" + "RKB/"+NumerNr+"/"+rejestr.Data.Year+"/"+rejestr.Nazwa + "'");
        //            //Debug.WriteLine("pomyślne dodanie Numeru");
        //            rZapis.RaportKB = rRaport;
                      
        //            rZapis.DataDok = rejestr.Data; //z ta data byl problem !!!!!
        //            rZapis.Kwota = zapis.Wartosc;
        //            //rZapis.NumerObcy = zapis.Konto;
        //            rZapis.Opis = zapis.Opis;
   
        //           // Debug.WriteLine("pomyślne dodanie danych bloku: RaportKB, DataDok, Kwota");
        //           // var rSeria = oSession.CreateObject("OP_KASBOLib.ZapisKB").Item("Seria1 = KASA");
        //        }
        //        catch(Exception e)
        //        {
        //            Debug.WriteLine("Błąd podczas ustawiania raportu: "+e.Message);
        //        } //RaportKB, Data, Kwota, NUMER OBCY !!! 

        //        rZapis.DefinicjaDokumentu = rNumerator.DefinicjaDokumentu;
        //        rZapis.Kierunek = zapis.Symbol; //nie wiem czemu 2
        //                             //rZapis.Seria = "KASA";
                
        //        //OP_KASBOLib.ZapisKB zapis = oSession.CreateObject(OP_KASBOLib.ZapisKB)
                
        //        try
        //        {
        //            var rKontrahent = oSession.CreateObject("CDN.Kontrahenci").Item("Knt_KOD = '" +"!NIEOKREŚLONY!"+"'");
        //            rZapis.Podmiot = rKontrahent;
        //            //Debug.WriteLine("Kontrahent: POMYŚLNIE");
        //        }
        //        catch(Exception e)
        //        {
        //            Debug.WriteLine("Kontrahent NIE POMYŚLNIE... "+ e.Message);
        //        } //Kontrahent

        //        oSession.Save();
        //        //Debug.WriteLine("sesja udana");
        //    }
        //    catch(Exception e)
        //    {
        //        Debug.WriteLine("Błąd tworzenia sesji: "+e.Message);
        //    }
        //}  //tworzenie zapisu KB

        //public static void ImportZPliku(Rejestr rejestr)
        //{
        //    //int numerNr = NumerNr;

        //    try
        //    {

        //        //tablica lini z pliku
        //        string[] lines = File.ReadAllLines(@"C:\Banki\" + rejestr.Numer + @"\" + rejestr.Data.Year + @"\" + rejestr.DataWpisana + @"\" + rejestr.PelnaNazwa);
        //        //string[] lines = File.ReadAllLines(@"C:\Banki\1572\2020\03.2020\1572_02032020");
        //        //ilość zapisów do wykonania
        //        int iloscZapisow = 0;
        //        int iloscLinii = 0;
        //        int dlugoscZapisu;

        //        foreach (string line in lines)
        //        {
        //            if (line.Contains("4:"))
        //            {
        //                iloscZapisow++;
        //            }
        //            iloscLinii++;
        //        }
        //        //Obliczenie długośći jednego zapisuKB
        //        dlugoscZapisu = iloscLinii / iloscZapisow;
        //        //tworzymy tablice obiektów
        //        Zapis[] zapisy = new Zapis[iloscZapisow];

        //        //Debug.WriteLine("Ilość lini: " + iloscZapisow);

        //        //int ktoryZapis = 0; //    do obiektów
        //        int aktualnaLinia = 0;
        //        int koniecPetli = dlugoscZapisu;
        //        int KtoryZapisOpis = 1;
        //        for (int j = 0; j < iloscZapisow; j++) //Petla zmieniająca zapis
        //        {
        //            Zapis zapis = new Zapis();
        //            //Debug.WriteLine("*******" + KtoryZapisOpis + "******");

        //            for (int i = aktualnaLinia; i < koniecPetli; i++) //pętla poszczególnych linii
        //            {
        //                //Debug.WriteLine(lines[i]); // sprawdzanie
        //                if (lines[i].Contains(":25:"))
        //                {
        //                    //Debug.WriteLine("***"+lines[i].Substring(7,26));
        //                    zapis.Konto = lines[i].Substring(7, 26);
        //                }
        //                if (lines[i].Contains(":28:"))
        //                {
        //                    //Debug.WriteLine("***"+lines[i].Substring(9));
        //                    zapis.Data = (DateTime.Parse(lines[i].Substring(9)));
        //                }
        //                if (lines[i].Contains(":60F:"))
        //                {
        //                    //Debug.WriteLine("***" + lines[i].Substring(15));
        //                    if (lines[i].Substring(15) == "0,00")
        //                    {
        //                        //Debug.WriteLine("***" + lines[i + 1].Substring(15));
        //                        zapis.Wartosc = decimal.Parse(lines[i + 1].Substring(15));
        //                        zapis.Symbol = 1;
        //                    }
        //                    else
        //                    {
        //                        zapis.Wartosc = decimal.Parse(lines[i].Substring(15));
        //                        zapis.Symbol = -1;
        //                    }
        //                }
        //                if (lines[i].Contains("^20"))
        //                {
        //                    //Debug.WriteLine("opis: "+lines[i].Substring(3));
        //                    zapis.Opis = lines[i].Substring(3);
        //                }
        //                zapisy[j] = zapis;
        //            }
        //            aktualnaLinia = aktualnaLinia + dlugoscZapisu;
        //            koniecPetli = aktualnaLinia + dlugoscZapisu;
        //            KtoryZapisOpis++;
        //        }


        //        Debug.WriteLine("OBIEKTY");
        //        foreach (Zapis zapis in zapisy)
        //        {
        //            try
        //            { //Debug.WriteLine(zapis.Konto); 
        //            }
        //            catch (Exception e) { Debug.WriteLine(e.Message); }
        //            try
        //            { //Debug.WriteLine(zapis.Data); 
        //            }
        //            catch (Exception e) { Debug.WriteLine(e.Message); }
        //            try
        //            { //Debug.WriteLine(zapis.Wartosc); 
        //            }
        //            catch (Exception e) { Debug.WriteLine(e.Message); }
        //            try
        //            { //Debug.WriteLine(zapis.Symbol); 
        //            }
        //            catch (Exception e) { Debug.WriteLine(e.Message); }
        //            try
        //            { //Debug.WriteLine(zapis.Opis); 
        //            }
        //            catch (Exception e) { Debug.WriteLine(e.Message); }

        //            //Zapis do Optimy!!!
        //            //ZapisKB(zapis, rejestr);
        //            ZapisKB( zapis, rejestr);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine("Plik nie został znalezniony" + e.Message);
        //    }
        //} //Parsowanie danych z pliku.
        
        //public static string EncodingDoWindows1250(string toEncode)
        //{
           
        //    string inputStr= toEncode;

        //    //byte[] bytes = new byte[encode.Length * sizeof(char)];
        //    //System.Buffer.BlockCopy(encode.ToCharArray(), 0, bytes, 0, bytes.Length);

        //    //Encoding w1252 = Encoding.GetEncoding(1252);
        //    //Encoding utf8 = Encoding.GetEncoding(65001);
        //    //byte[] output = Encoding.Convert(utf8, w1252, bytes);
        //    //w1252.GetString(output);

        //    //return w1252.ToString();


        //    // get the correct encodings 
        //    var srcEncoding = Encoding.UTF8; // utf-8
        //    var destEncoding = Encoding.GetEncoding(1252); // windows-1252

        //    // convert the source bytes to the destination bytes
        //    var destBytes = Encoding.Convert(srcEncoding, destEncoding, srcEncoding.GetBytes(inputStr));

        //    // process the byte[]
        //    //File.WriteAllBytes("myFile", destBytes); // write it to a file OR ...
        //    var destString = destEncoding.GetString(destBytes); // ... get the string
            
        //    return destString;
        //}

        //#endregion 
    }
}
