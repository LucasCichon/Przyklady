using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CDNBase;

namespace Import
{
    class ImportClass
    {
        public static void ImportZPliku()
        {
            //tablica lini z pliku
            string[] lines = File.ReadAllLines(@"C:\Banki\1572\2020\03.2020\5041_02032020");
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

            Debug.WriteLine("Ilość lini: " + iloscZapisow);
           
            //int ktoryZapis = 0; //    do obiektów
            int aktualnaLinia = 0;
            int koniecPetli = dlugoscZapisu;
            int KtoryZapisOpis = 1;
            for (int j = 0; j < iloscZapisow; j++) //Petla zmieniająca zapis
            {
                Zapis zapis = new Zapis();
                Debug.WriteLine("*******" + KtoryZapisOpis + "******");
                
                for(int i = aktualnaLinia; i<koniecPetli; i++  ) //pętla poszczególnych linii
                {
                    //Debug.WriteLine(lines[i]); // sprawdzanie
                    if (lines[i].Contains(":25:"))
                    {
                        Debug.WriteLine("***"+lines[i].Substring(7,26));
                        zapis.Konto = lines[i].Substring(7, 26);
                    }
                    if (lines[i].Contains(":28:"))
                    {
                        Debug.WriteLine("***"+lines[i].Substring(9));
                        zapis.Data = (DateTime.Parse(lines[i].Substring(9)));
                    }
                    if (lines[i].Contains(":60F:"))
                    {
                        Debug.WriteLine("***" + lines[i].Substring(15));
                        if (lines[i].Substring(15) == "0,00")
                        {
                            Debug.WriteLine("***" + lines[i + 1].Substring(15));
                            zapis.Wartosc = decimal.Parse(lines[i + 1].Substring(15));
                            zapis.Symbol = 1;
                        }
                        else
                        {
                            zapis.Wartosc = decimal.Parse(lines[i].Substring(15));
                            zapis.Symbol = -1;
                        }
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
                try { Debug.WriteLine(zapis.Konto); }
                catch(Exception e) { Debug.WriteLine(e.Message); }
                try { Debug.WriteLine(zapis.Data); }
                catch(Exception e) { Debug.WriteLine(e.Message); }
                try { Debug.WriteLine(zapis.Wartosc); }
                catch (Exception e) { Debug.WriteLine(e.Message); }
                try { Debug.WriteLine(zapis.Symbol); }
                catch (Exception e) { Debug.WriteLine(e.Message); }

                //Zapis do Optimy!!!
                ZapisKB(zapis);
            }
        }

        public static void ZapisKB(Zapis zapis)
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
                var rRachunek = oSession.CreateObject("CDN.Rachunki").Item("Bra_Akronim = '" + "KASA" + "'"); //Rejestr
                rZapis.Rachunek = rRachunek;
                    Debug.WriteLine("pomyślne ustawienie rachunku");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Błąd podczas ustawiania rachunku" + Environment.NewLine+e.Message);
                } // Ustawienie rachunku
               
                try 
                {
                    var rRaport = oSession.CreateObject("CDN.RaportyKB").Item("BRp_NumerPelny = '" + "RKB/1/2020/KASA" + "'");
                    Debug.WriteLine("pomyślne dodanie Numeru");
                    rZapis.RaportKB = rRaport;
                      
                    rZapis.DataDok = zapis.Data;
                    rZapis.Kwota = zapis.Wartosc;
                    //rZapis.NumerObcy = zapis.Konto;
                    rZapis.Opis = "Opis z C#";
   
                    Debug.WriteLine("pomyślne dodanie danych bloku: RaportKB, DataDok, Kwota");
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
                    Debug.WriteLine("Kontrahent: POMYŚLNIE");
                }
                catch(Exception e)
                {
                    Debug.WriteLine("Kontrahent NIE POMYŚLNIE... "+ e.Message);
                } //Kontrahent

                oSession.Save();
                Debug.WriteLine("sesja udana");
            }
            catch(Exception e)
            {
                Debug.WriteLine("Błąd tworzenia sesji: "+e.Message);
            }
        }
    }
}
