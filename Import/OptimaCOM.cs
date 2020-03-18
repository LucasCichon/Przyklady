using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CDNBase;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Import
{  
    public class OptimaCOM
    {
        protected static IApplication oApp = null;
        public static ILogin oLogin = null;

        public static string _oOper = string.Empty;
        public static DateTime OdokDate = DateTime.Now;


        public static void O_Login(string oOper, string oPass, string oBase = "WIKANYS")
        {
            Program.log.Info("OptimaCOM.O_Login() - Logowanie w systemie ERP Optima.");
            //_oOper = oOper;
            Debug.WriteLine("Próba logowania...");
            try
            {
                System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
                oApp = new CDNBase.Application();
                oLogin = oApp.Login(oOper, oPass, oBase);
                Program.log.Info("zalogowano!!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("błąd");
                Program.log.Info("OptimaCOM.O_Login() - " + ex.Message);
            }
        }

        public static void O_Logout()
        {
            Program.log.Info("OptimaCOM.O_Logout() - Wylogowanie z systemu ERP Optima.");
            Debug.WriteLine("Próba wylogowania...");
            try
            {
                oLogin = null;

                oApp.UnlockApp();
                oApp = null;
                Debug.WriteLine("Wylogowano pomyślnie!!!");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Błąd podczas wylogowywania: " +ex.Message);
                Program.log.Info("OptimaCOM.O_Logout() - " + ex.Message);
            }
        }

        //    protected static string O_API_ErrorMessage(System.Exception e)
        //    {
        //        Common.log.Info("OptimaCOM.O_API_ErrorMessage() - Generowanie wiadomości o błędzie.");

        //        StringBuilder mess = new StringBuilder();
        //        if (e != null)
        //        {
        //            mess.Append(e.Message);
        //            while (e.InnerException != null)
        //            {
        //                mess.Append(e.InnerException.Message);
        //                e = e.InnerException;
        //            }
        //        }
        //        return mess.ToString();
        //    }

        //public static void O_DodanieTowaru(string twrKod, string twrVendor, string twrNazwa, string twrKodDost, string twrEan, double twrVat)
        //{
        //    Common.log.Info($@"OptimaCOM.O_DodanieTowaru() - Dodanie towaru: {twrKod} | {twrNazwa}");
        //    try
        //    {
        //        CDNBase.AdoSession oSession = oLogin.CreateSession();

        //        CDNTwrb1.Towary Towary = (CDNTwrb1.Towary)oSession.CreateObject("CDN.Towary", null);
        //        CDNTwrb1.ITowar Towar = (CDNTwrb1.ITowar)Towary.AddNew(null);

        //        Towar.Kod = StringExtensions.Left(twrKod, 40).ToUpper();
        //        if (Database.GetProductName(twrKod) != string.Empty && Database.GetProductName(twrKod) != null)
        //        {
        //            Towar.Nazwa = Database.GetProductName(twrKod);
        //        }
        //        else
        //        {
        //            Towar.Nazwa = twrNazwa;
        //        }
        //        Towar.NumerKat = twrVendor;
        //        Towar.DostawcaID = 2;
        //        Towar.KodDostawcy = twrKodDost;
        //        Towar.EAN = twrEan;
        //        if (twrNazwa.ToLower().Contains("iiyama"))
        //        {
        //            Towar.PrdID = 1;
        //            Towar.ProducentKod = twrVendor;
        //        }
        //        Towar.Flaga = 1;
        //        Towar.JM = "szt";
        //        switch (twrVat)
        //        {
        //            case 23:
        //                Towar.StawkaVat = "23.00 %";
        //                Towar.StawkaVatZak = "23.00 %";
        //                break;
        //            case 22:
        //                Towar.StawkaVat = "22.00 %";
        //                Towar.StawkaVatZak = "22.00 %";
        //                break;
        //            case 8:
        //                Towar.StawkaVat = "8.00 %";
        //                Towar.StawkaVatZak = "8.00 %";
        //                break;
        //            case 7:
        //                Towar.StawkaVat = "7.00 %";
        //                Towar.StawkaVatZak = "7.00 %";
        //                break;
        //            case 6:
        //                Towar.StawkaVat = "6.00 %";
        //                Towar.StawkaVatZak = "6.00 %";
        //                break;
        //            case 5:
        //                Towar.StawkaVat = "5.00 %";
        //                Towar.StawkaVatZak = "5.00 %";
        //                break;
        //            case 4:
        //                Towar.StawkaVat = "4.00 %";
        //                Towar.StawkaVatZak = "4.00 %";
        //                break;
        //            case 3:
        //                Towar.StawkaVat = "3.00 %";
        //                Towar.StawkaVatZak = "3.00 %";
        //                break;
        //            case 0:
        //                Towar.StawkaVat = "0.00 %";
        //                Towar.StawkaVatZak = "0.00 %";
        //                break;

        //            default:
        //                Towar.StawkaVat = "23.00 %";
        //                Towar.StawkaVatZak = "23.00 %";
        //                break;
        //        }

        //        Towar.JMCalkowite = 1;
        //        Towar.EdycjaNazwy = 1;
        //        Towar.EdycjaOpisu = 1;

        //        CDNBase.ICollection Grupy = (CDNBase.ICollection)(oSession.CreateObject("CDN.TwrGrupy", null));
        //        CDNTwrb1.TwrGrupa Grupa;
        //        if (twrNazwa.ToLower().Contains("iiyama"))
        //        {
        //            Grupa = (CDNTwrb1.TwrGrupa)Grupy["twg_kod = 'IIYAMA'"];
        //        }
        //        else
        //        {
        //            Grupa = (CDNTwrb1.TwrGrupa)Grupy["twg_kod = 'Grupa Główna'"];
        //        }

        //        Towar.TwGGIDNumer = Grupa.GIDNumer;
        //        CDNTwrb1.TwrGrupa GrupaZListy = (CDNTwrb1.TwrGrupa)Towar.Grupy.AddNew(null);
        //        GrupaZListy = Grupa;

        //        oSession.Save();
        //    }
        //    catch (Exception comErr)
        //    {
        //        Console.WriteLine(comErr.Message);
        //        //Common.log.Info("OptimaCOM.DodanieTowaru() - " + O_API_ErrorMessage(comErr));
        //    }
        //}

        //    public static async void O_DodanieKontrahenta(Int32 order_id)
        //    {
        //        Common.log.Info($@"OptimaCOM.O_DodanieKontrahenta({order_id})");

        //        Models.Linker.invoice invoice = Database.GetInvoiceData(order_id);

        //        int _payment_method_days = 0;
        //        int.TryParse(Database.GetPaymentMethod(order_id).ToLower().Replace("przelew", "").Replace("dni", "").Trim(), out _payment_method_days);
        //        if (_payment_method_days > 0)
        //        {
        //            //
        //        }
        //        int _payment_method_cod = Database.GetPaymentMethodCod(order_id);
        //        string _order_source = Database.GetOrderSource(order_id);

        //        if (invoice != null)
        //        {
        //            try
        //            {
        //                string country_code = Database.GetDeliveryCountryCode(order_id);

        //                CDNBase.AdoSession oSession = oLogin.CreateSession();

        //                OP_KASBOLib.Banki Banki = (OP_KASBOLib.Banki)oSession.CreateObject("CDN.Banki", null);
        //                CDNHeal.Kategorie Kategorie = (CDNHeal.Kategorie)oSession.CreateObject("CDN.Kategorie", null);

        //                CDNHeal.Kontrahenci Kontrahenci = (CDNHeal.Kontrahenci)oSession.CreateObject("CDN.Kontrahenci", null);
        //                CDNHeal.IKontrahent Kontrahent = (CDNHeal.IKontrahent)Kontrahenci.AddNew(null);

        //                CDNBase.ICollection FormyPlatnosci = (CDNBase.ICollection)(oSession.CreateObject("CDN.FormyPlatnosci", null));
        //                OP_KASBOLib.FormaPlatnosci FPl = (OP_KASBOLib.FormaPlatnosci)FormyPlatnosci[2]; // Przelew

        //                CDNHeal.IAdres Adres = Kontrahent.Adres;
        //                CDNHeal.INumerNIP NumerNIP = Kontrahent.NumerNIP;

        //                if (country_code == "PL")
        //                {
        //                    NumerNIP.UstawNIP("PL", "", 1);
        //                    Adres.Kraj = "Polska";

        //                    Kontrahent.Export = 0;
        //                }
        //                else if (country_code == "CZ")
        //                {
        //                    NumerNIP.UstawNIP("CZ", "", 1);
        //                    Adres.Kraj = "Czechy";

        //                    Kontrahent.Export = 3;
        //                    Kontrahent.KodTransakcji = "11";
        //                }
        //                else if (country_code == "SK")
        //                {
        //                    NumerNIP.UstawNIP("SK", "", 1);
        //                    Adres.Kraj = "Słowacja";

        //                    Kontrahent.Export = 3;
        //                    Kontrahent.KodTransakcji = "11";
        //                }
        //                Adres.KodPocztowy = invoice.postcode.Trim();
        //                Adres.Miasto = invoice.city.Trim();
        //                Adres.Ulica = invoice.address.Trim();

        //                Kontrahent.Akronim = order_id.ToString();
        //                if (invoice.fullname.Trim().Length > 50)
        //                {
        //                    Kontrahent.Nazwa1 = invoice.fullname.Trim().Substring(0, 50);
        //                    Kontrahent.Nazwa2 = invoice.fullname.Trim().Substring(50, invoice.fullname.Trim().Length - 50);

        //                    Common.log.Debug("OptimaCOM.O_DodanieKontrahenta() ->");
        //                    Common.log.Debug(invoice.fullname.Trim().Substring(0, 50));
        //                    Common.log.Debug(invoice.fullname.Trim().Substring(50, invoice.fullname.Trim().Length - 50));
        //                }
        //                else
        //                {
        //                    Kontrahent.Nazwa1 = invoice.fullname.Trim();
        //                }

        //                Kontrahent.Finalny = 1;
        //                Kontrahent.PodatekVAT = 0;

        //                if (invoice.nip.Trim() != "")
        //                {
        //                    if (country_code == "PL")
        //                    {
        //                        NumerNIP.UstawNIP("PL", invoice.nip.Trim().Replace(" ", "").Replace("-", "").ToString().Trim(), 1);
        //                    }
        //                    else if (country_code == "CZ")
        //                    {
        //                        NumerNIP.UstawNIP("CZ", invoice.nip.Trim().Replace(" ", "").Replace("-", "").ToString().Trim(), 1);
        //                    }
        //                    else if (country_code == "SK")
        //                    {
        //                        NumerNIP.UstawNIP("SK", invoice.nip.Trim().Replace(" ", "").Replace("-", "").ToString().Trim(), 1);
        //                    }


        //                    Kontrahent.Akronim = invoice.nip.Trim();
        //                    if (invoice.company.Trim().Length > 50)
        //                    {
        //                        Kontrahent.Nazwa1 = invoice.company.Trim().Substring(0, 50);
        //                        Kontrahent.Nazwa2 = invoice.company.Trim().Substring(50, invoice.company.Trim().Length - 50);

        //                        Common.log.Debug("OptimaCOM.O_DodanieKontrahenta() ->");
        //                        Common.log.Debug(invoice.company.Trim().Substring(0, 50));
        //                        Common.log.Debug(invoice.company.Trim().Substring(50, invoice.company.Length - 50));
        //                    }
        //                    else
        //                    {
        //                        Kontrahent.Nazwa1 = invoice.company.Trim();
        //                    }
        //                    Kontrahent.Finalny = 0;
        //                    Kontrahent.PodatekVAT = 1;
        //                }

        //                Common.log.Debug($@"OptimaCOM.O_DodanieKontrahenta({order_id}) - " + Kontrahent.Akronim);

        //                Models.Linker.invoice invoice_data = Database.GetInvoiceData(order_id);

        //                Kontrahent.KrajISO = country_code;
        //                Kontrahent.Telefon = invoice_data.phone.Trim();
        //                Kontrahent.Email = invoice_data.email.Trim();
        //                Kontrahent.FormaPlatnosci = FPl;
        //                Kontrahent.TerminPlat = 1;
        //                Kontrahent.Termin = _payment_method_days;

        //                oSession.Save();
        //            }
        //            catch (Exception comErr)
        //            {
        //                if (!comErr.ToString().Contains("Nie można wstawić wiersza zduplikowanego klucza w obiekcie CDN.Kontrahenci o unikatowym indeksie KntKod"))
        //                {
        //                    Common.log.Info($@"OptimaCOM.O_DodanieKontrahenta({order_id}) - " + O_API_ErrorMessage(comErr));

        //                    await BaseLinker.setOrderFields(order_id, "extra_field_2", "ERP.Kontrahent()");
        //                    await BaseLinker.setOrderStatus(order_id, int.Parse(Properties.Settings.Default.BaseLinker_Status_Error.Trim()));
        //                }
        //            }
        //        }
        //    }

        //    public static async void O_DodanieDokRO(Int32 order_id)
        //    {
        //        Common.log.Info($@"OptimaCOM.O_DodanieDokRO({order_id})");
        //        try
        //        {
        //            string country_code = Database.GetDeliveryCountryCode(order_id);
        //            Models.BaseLinker.order order = Database.GetOrderData(order_id);
        //            List<Models.BaseLinker.product> products = Database.GetOrderProducts(order_id);
        //            Models.Linker.invoice invoice = Database.GetInvoiceData(order_id);

        //            string _customer = order.invoice_nip != string.Empty ? order.invoice_nip : order_id.ToString();
        //            decimal delivery_price = Database.GetDeliveryPrice(order_id);
        //            int payment_method_days = 0;
        //            int.TryParse(Database.GetPaymentMethod(order_id).ToLower().Replace("przelew", "").Replace("dni", "").Trim(), out payment_method_days);
        //            if (payment_method_days > 0)
        //            {
        //                //
        //            }
        //            Common.log.Debug("payment_method_days:" + payment_method_days);

        //            if (Database.chkIfReservationExists(order_id) == 0)
        //            {
        //                CDNBase.AdoSession oSession = oLogin.CreateSession();

        //                CDNHlmn.DokumentyHaMag Dokumenty = (CDNHlmn.DokumentyHaMag)oSession.CreateObject("CDN.DokumentyHaMag", null);
        //                CDNHlmn.IDokumentHaMag RO = (CDNHlmn.IDokumentHaMag)Dokumenty.AddNew(null);

        //                CDNBase.ICollection Kontrahenci = (CDNBase.ICollection)(oSession.CreateObject("CDN.Kontrahenci", null));
        //                CDNHeal.IKontrahent Kontrahent = (CDNHeal.IKontrahent)Kontrahenci["Knt_Kod='" + _customer.Trim() + "'"];

        //                CDNBase.ICollection waluty = (CDNBase.ICollection)(oSession.CreateObject("CDN.Waluty", null));

        //                CDNBase.ICollection FormyPlatnosci = (CDNBase.ICollection)(oSession.CreateObject("CDN.FormyPlatnosci", null));

        //                OP_KASBOLib.FormaPlatnosci FPl;
        //                Common.log.Debug("dMetodaPlatnosci: " + order.payment_method.Trim());
        //                int CDN_PaymentMethodId = Database.CDN_GetPaymentCode(order.payment_method.ToLower());

        //                FPl = (OP_KASBOLib.FormaPlatnosci)FormyPlatnosci[Database.CDN_GetPaymentCode(order.payment_method.ToLower())];
        //                if (order.order_source == "allegro")
        //                {
        //                    FPl = (OP_KASBOLib.FormaPlatnosci)FormyPlatnosci[Database.CDN_GetPaymentCode("przelew Allegro")];
        //                    CDN_PaymentMethodId = Database.CDN_GetPaymentCode("przelew Allegro");
        //                }
        //                if (order.payment_method_cod == 1)
        //                {
        //                    FPl = (OP_KASBOLib.FormaPlatnosci)FormyPlatnosci[Database.CDN_GetPaymentCode("pobranie")];
        //                    CDN_PaymentMethodId = Database.CDN_GetPaymentCode("pobranie");
        //                }

        //                Common.log.Debug("Forma płatności id: " + CDN_PaymentMethodId);

        //                //if (order.payment_method.ToLower().Contains("Płatność przy odbiorze".ToLower()) || order.payment_method.ToLower().Contains("Pobranie".ToLower()))
        //                //{
        //                //    FPl = (OP_KASBOLib.FormaPlatnosci)FormyPlatnosci[Database.CDN_GetPaymentCode("pobranie")];
        //                //}

        //                CDNHeal.DefinicjeDokumentow DefinicjeDokumentow = (CDNHeal.DefinicjeDokumentow)(oSession.CreateObject("CDN.DefinicjeDokumentow", null));
        //                CDNHeal.DefinicjaDokumentu DefinicjaDokumentu = (CDNHeal.DefinicjaDokumentu)DefinicjeDokumentow["DDf_DDfID=1"];

        //                OP_KASBOLib.INumerator numerator = (OP_KASBOLib.INumerator)RO.Numerator;

        //                RO.Rodzaj = 308000;
        //                RO.TypDokumentu = 308;
        //                RO.MagazynZrodlowyID = 1;

        //                RO.Bufor = 1;
        //                RO.Podmiot = (CDNHeal.IPodmiot)Kontrahent;
        //                RO.FormaPlatnosci = FPl;
        //                RO.DataDok = Common.UnixTimeStampToDateTime(order.date_add);
        //                RO.DataWys = Common.UnixTimeStampToDateTime(order.date_confirmed);
        //                RO.Uwagi = Database.GetOrderUserLogin(order_id);

        //                #region Atrybuty
        //                CDNHlmn.AtrybutHaMag Atrybut = (CDNHlmn.AtrybutHaMag)oSession.CreateObject("CDN.AtrybutHaMag", null);
        //                CDNTwrb1.IDokAtrybut Atrybut1 = (CDNTwrb1.IDokAtrybut)RO.Atrybuty.AddNew(null);
        //                CDNTwrb1.IDokAtrybut Atrybut2 = (CDNTwrb1.IDokAtrybut)RO.Atrybuty.AddNew(null);
        //                CDNTwrb1.IDokAtrybut Atrybut3 = (CDNTwrb1.IDokAtrybut)RO.Atrybuty.AddNew(null);
        //                CDNTwrb1.IDokAtrybut Atrybut4 = (CDNTwrb1.IDokAtrybut)RO.Atrybuty.AddNew(null);
        //                CDNTwrb1.IDokAtrybut Atrybut5 = (CDNTwrb1.IDokAtrybut)RO.Atrybuty.AddNew(null);

        //                Atrybut1.Kod = "BL_NR_ZAMÓWIENIA";
        //                Atrybut1.Wartosc = order_id.ToString().Trim();
        //                Atrybut2.Kod = "BL_ŹRÓDŁO";
        //                Atrybut2.Wartosc = order.order_source.ToString().Trim();
        //                Atrybut3.Kod = "BL_SPEDYTOR";
        //                Atrybut3.Wartosc = order.delivery_package_module.ToString().Trim();
        //                Atrybut4.Kod = "BL_NR_LIST_PRZEWOZ";
        //                Atrybut4.Wartosc = order.delivery_package_nr.ToString().Trim();
        //                Atrybut5.Kod = "BL_ID_PŁATNOŚCI";
        //                await BaseLinker.getOrderPaymentsHistory(order.order_id);
        //                if (Common.bl_payment_history != null)
        //                {
        //                    if (Common.bl_payment_history.Count > 0)
        //                    {
        //                        string payment_id = string.Empty;
        //                        foreach (var pmt in Common.bl_payment_history)
        //                        {
        //                            if (pmt.external_payment_id != string.Empty)
        //                            {
        //                                payment_id = pmt.external_payment_id;
        //                            }
        //                        }
        //                        Atrybut5.Wartosc = payment_id;
        //                    }
        //                    else
        //                    {
        //                        Atrybut5.Wartosc = string.Empty;
        //                    }
        //                }
        //                else
        //                {
        //                    Atrybut5.Wartosc = string.Empty;
        //                }
        //                #endregion

        //                RO.TypNB = 2;

        //                if (country_code == "PL")
        //                {
        //                    CDNHeal.Waluta waluta = (CDNHeal.Waluta)waluty["WNa_Symbol='PLN'"];

        //                    RO.Waluta = waluta;
        //                }

        //                List<Int32> lp = new List<int>();
        //                CDNBase.ICollection Pozycje = RO.Elementy;
        //                foreach (var poz in products)
        //                {
        //                    Common.log.Debug(Environment.NewLine + "order_product_id:" + poz.order_product_id + Environment.NewLine + "sku:'" + poz.sku + "'" + Environment.NewLine + "auction_id:" + poz.auction_id);

        //                    if (lp.Contains(poz.order_product_id) == false)
        //                    {
        //                        int err_auction = 0;
        //                        int err_sku = 0;

        //                        if (poz.auction_id != string.Empty)
        //                        {
        //                            Common.log.Debug($@"CDN_GetProductCodeFromAuctionId({poz.auction_id}): " + Database.CDN_GetProductCodeFromAuctionId(poz.auction_id));
        //                            Common.log.Debug($@"chkIfTwrIsProduct({poz.auction_id}): " + Database.chkIfTwrIsProduct(poz.auction_id));

        //                            if (Database.chkIfTwrIsProduct(poz.auction_id) == 1)
        //                            {
        //                                err_auction = 0;
        //                            }
        //                            else
        //                            {
        //                                if (Database.CDN_GetProductCodeFromAuctionId(poz.auction_id) == string.Empty)
        //                                {
        //                                    err_auction = 1;
        //                                }
        //                            }

        //                            if (Database.CDN_CheckIfProductExists(poz.sku, "sku") == 0)
        //                            {
        //                                err_sku = 1;
        //                            }
        //                            else
        //                            {
        //                                err_sku = 0;
        //                            }
        //                        }
        //                        else
        //                        {
        //                            err_auction = 1;

        //                            if (Database.CDN_CheckIfProductExists(poz.sku, "sku") == 0)
        //                            {
        //                                err_sku = 1;
        //                            }
        //                            else
        //                            {
        //                                err_sku = 0;
        //                            }
        //                        }

        //                        Common.log.Debug(Environment.NewLine + "auction_id:" + poz.auction_id + Environment.NewLine + "err_auction:" + err_auction + Environment.NewLine + "sku:" + poz.sku + Environment.NewLine + "err_sku:" + err_sku);

        //                        if (err_auction == 0)
        //                        {
        //                            if (Database.chkIfTwrIsProduct(poz.auction_id) == 1)
        //                            {
        //                                List<string> rec = Database.getRecipeIds(poz.auction_id);
        //                                int recipe_count = rec.Count - 1;
        //                                int fp = 1;

        //                                foreach (string twr in rec)
        //                                {
        //                                    CDNHlmn.IElementHaMag Pozycja = (CDNHlmn.IElementHaMag)Pozycje.AddNew(null);
        //                                    Common.log.Debug("twr.id:" + twr);
        //                                    Pozycja.TowarKod = twr;
        //                                    Pozycja.Ilosc = poz.quantity;
        //                                    if (country_code == "PL")
        //                                    {
        //                                        Pozycja.Stawka = 23;
        //                                        if (fp == 1)
        //                                        {
        //                                            Pozycja.Cena0 = decimal.Parse(poz.price_brutto.ToString().Replace(".", ",")) - recipe_count;
        //                                        }
        //                                        else
        //                                        {
        //                                            Pozycja.Cena0 = 1;
        //                                        }
        //                                    }
        //                                    else
        //                                    {
        //                                        if (invoice.nip != string.Empty)
        //                                        {
        //                                            Pozycja.Stawka = 0;
        //                                            if (fp == 1)
        //                                            {
        //                                                Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ",")) - recipe_count;
        //                                            }
        //                                            else
        //                                            {
        //                                                Pozycja.Cena0WD = 1;
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            Pozycja.Stawka = 23;
        //                                            if (fp == 1)
        //                                            {
        //                                                Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ",")) - recipe_count;
        //                                            }
        //                                            else
        //                                            {
        //                                                Pozycja.Cena0WD = 1;
        //                                            }
        //                                        }
        //                                    }
        //                                    fp++;
        //                                }
        //                                lp.Add(poz.order_product_id);
        //                            }
        //                            else
        //                            {
        //                                CDNHlmn.IElementHaMag Pozycja = (CDNHlmn.IElementHaMag)Pozycje.AddNew(null);
        //                                Pozycja.TowarKod = Database.CDN_GetProductCodeFromAuctionId(poz.auction_id);
        //                                Pozycja.Ilosc = poz.quantity;
        //                                if (country_code == "PL")
        //                                {
        //                                    Pozycja.Stawka = 23;
        //                                    Pozycja.Cena0 = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                }
        //                                else
        //                                {
        //                                    if (invoice.nip != string.Empty)
        //                                    {
        //                                        Pozycja.Stawka = 0;
        //                                        Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                    }
        //                                    else
        //                                    {
        //                                        Pozycja.Stawka = 23;
        //                                        Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                    }
        //                                }
        //                                lp.Add(poz.order_product_id);
        //                            }
        //                        }
        //                        else
        //                        {
        //                            if (err_sku == 0)
        //                            {
        //                                CDNHlmn.IElementHaMag Pozycja = (CDNHlmn.IElementHaMag)Pozycje.AddNew(null);
        //                                Pozycja.TowarKod = poz.sku;
        //                                Pozycja.Ilosc = poz.quantity;

        //                                if (country_code == "PL")
        //                                {
        //                                    Pozycja.Stawka = 23;
        //                                    Pozycja.Cena0 = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                }
        //                                else
        //                                {
        //                                    if (invoice.nip != string.Empty)
        //                                    {
        //                                        Pozycja.Stawka = 0;
        //                                        Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                    }
        //                                    else
        //                                    {
        //                                        Pozycja.Stawka = 23;
        //                                        Pozycja.Cena0WD = decimal.Parse(poz.price_brutto.ToString().Replace(".", ","));
        //                                    }
        //                                }
        //                                lp.Add(poz.order_product_id);
        //                            }
        //                            else
        //                            {
        //                                Common.log.Error("OptimaCOM.O_DodanieDokRO() -> Brak towaru.");
        //                                await BaseLinker.setOrderFields(order_id, "extra_field_2", StringExtensions.Left("Optima-" + "Brak towaru", 50));
        //                                await BaseLinker.setOrderStatus(order_id, int.Parse(Properties.Settings.Default.BaseLinker_Status_Error.Trim()));
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        //
        //                    }
        //                }

        //                Common.log.Debug("dPobranie: " + order.payment_method_cod);
        //                if (delivery_price > 0)
        //                {
        //                    CDNHlmn.IElementHaMag Pozycja = (CDNHlmn.IElementHaMag)Pozycje.AddNew(null);
        //                    Pozycja.TowarKod = "KOSZT DOSTAWY";
        //                    Pozycja.Ilosc = 1;
        //                    if (country_code == "PL")
        //                    {
        //                        Pozycja.Stawka = 23;
        //                        Pozycja.Cena0 = delivery_price;
        //                    }
        //                    else
        //                    {
        //                        if (invoice.nip != string.Empty)
        //                        {
        //                            Pozycja.Stawka = 0;
        //                            Pozycja.Cena0WD = delivery_price;
        //                        }
        //                        else
        //                        {
        //                            Pozycja.Stawka = 23;
        //                            Pozycja.Cena0WD = delivery_price;
        //                        }
        //                    }
        //                }
        //                oSession.Save();

        //                Common.log.Info($@"OptimaCOM.O_DodanieDokRO({order_id}) - Dodanie dokumentu RO: Utworzono dokument " + RO.NumerPelny);
        //                await BaseLinker.setOrderStatus(order_id, int.Parse(Properties.Settings.Default.BaseLinker_Status_Complete.Trim()));
        //            }
        //            else
        //            {
        //                Common.log.Error($@"OptimaCOM.O_DodanieDokRO({order_id}) -> " + "Dokument już istnieje pod TrNId:" + Database.chkIfReservationExists(order_id));
        //            }
        //        }
        //        catch (Exception comErr)
        //        {
        //            Common.log.Info($@"OptimaCOM.O_DodanieDokRO({order_id}) - " + O_API_ErrorMessage(comErr));
        //            try
        //            {
        //                if (comErr.ToString().Contains("Nie można wstawić wiersza zduplikowanego klucza w obiekcie CDN.TraNag o unikatowym indeksie TrNNumer"))
        //                {
        //                    Common.log.Error(StringExtensions.Left("ERP.RO()-" + "Dokument już istnieje.", 50));
        //                }
        //                else
        //                {
        //                    Common.log.Error("ERP.RO()-> " + comErr.Message);
        //                    await BaseLinker.setOrderFields(order_id, "extra_field_2", StringExtensions.Left("ERP.RO()-" + "Bład dokumentu", 50));
        //                    await BaseLinker.setOrderStatus(order_id, int.Parse(Properties.Settings.Default.BaseLinker_Status_Error.Trim()));
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Common.log.Error($@"OptimaCOM.O_DodanieDokRO({order_id}) - " + ex.Message);
        //            }
        //        }
        //    }
    }
    

//    Pozdrawiam serdecznie


//Michał Koźlik  |  Kierownik działu ERP
//tel.  77 442 71 61
//kom. 510 98 27 28




//HSI Sp. z o.o.ul.Technologiczna 2, 45-839 Opole
//NIP: 754-033-48-43 REGON: 008305006 GIOŚ: E0009567Z Sąd Rejonowy w Opolu, VIII Wydział Gospodarczy KRS: 0000198842
//Kapitał zakładowy: 100 000,00 PLN Rok założenia: 1989 Bank BPH S.A. 28 1060 0076 0000 3200 0033 4948


//Treść tej wiadomości zawiera informacje przeznaczone tylko dla adresata.Jeżeli nie jesteście Państwo jej adresatem, bądź otrzymaliście ją przez pomyłkę, prosimy o powiadomienie o tym nadawcy oraz trwałe jej usunięcie.


}
