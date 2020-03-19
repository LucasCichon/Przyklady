using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Data.Common;
using System.Configuration;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace Import
{
    public partial class Form1 : Form
    {

        private MonthCalendar Kalendarz;
        private Button BtnImportZapisow;
        private ListBox listBoxLog;
        private ContextMenuStrip contextMenuStrip1;
        private System.ComponentModel.IContainer components;
        private Button btnClose;
        public static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Form1()
        {
            InitializeComponent();
            Trace.Listeners.Add(new ListBoxTraceListener(listBoxLog));
        }

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            

            //OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");

            //List<Rejestr> ListaRejestrow = ImportClass.SzukaniePlikow(DateTime.Now.ToString());

            ////int NumerNr;
            //for (int i = 0; i < ListaRejestrow.Count; i++)
            //{
            //    int NumerNr = ImportClass.ZwrocNumerNr(ListaRejestrow[i]);
            //    ImportClass.ImportZPlikuKonversja(NumerNr, ListaRejestrow[i]);

                //if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                //{
                //    ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
                //}
                //else
                //{
                //    ImportClass.NowyRaport(ListaRejestrow[i]);
                //    if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                //    {
                //        ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
                //    }
                //}
            //}
            //OptimaCOM.O_Logout();
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.Kalendarz = new System.Windows.Forms.MonthCalendar();
            this.BtnImportZapisow = new System.Windows.Forms.Button();
            this.listBoxLog = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnClose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Kalendarz
            // 
            this.Kalendarz.BackColor = System.Drawing.SystemColors.Window;
            this.Kalendarz.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Kalendarz.Location = new System.Drawing.Point(18, 18);
            this.Kalendarz.MaxSelectionCount = 1;
            this.Kalendarz.Name = "Kalendarz";
            this.Kalendarz.TabIndex = 0;
            // 
            // BtnImportZapisow
            // 
            this.BtnImportZapisow.BackColor = System.Drawing.SystemColors.Window;
            this.BtnImportZapisow.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.BtnImportZapisow.Location = new System.Drawing.Point(18, 192);
            this.BtnImportZapisow.Name = "BtnImportZapisow";
            this.BtnImportZapisow.Size = new System.Drawing.Size(124, 38);
            this.BtnImportZapisow.TabIndex = 1;
            this.BtnImportZapisow.Text = "Import";
            this.BtnImportZapisow.UseVisualStyleBackColor = false;
            this.BtnImportZapisow.Click += new System.EventHandler(this.BtnImportZapisow_Click);
            // 
            // listBoxLog
            // 
            this.listBoxLog.FormattingEnabled = true;
            this.listBoxLog.HorizontalScrollbar = true;
            this.listBoxLog.Location = new System.Drawing.Point(299, 18);
            this.listBoxLog.Name = "listBoxLog";
            this.listBoxLog.Size = new System.Drawing.Size(406, 212);
            this.listBoxLog.TabIndex = 2;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.SystemColors.Window;
            this.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnClose.Location = new System.Drawing.Point(161, 192);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(126, 38);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "Wyjdź";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // Form1
            // 
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(726, 250);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.listBoxLog);
            this.Controls.Add(this.BtnImportZapisow);
            this.Controls.Add(this.Kalendarz);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Name = "Form1";
            this.Text = "Importowanie Zapisów KB";
            this.ResumeLayout(false);

        }

        private void BtnImportZapisow_Click(object sender, EventArgs e)
        {
            OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");

            List<Rejestr> ListaRejestrow = ImportClass.SzukaniePlikow(Kalendarz.SelectionRange.Start.ToShortDateString());

            //DateTime data = Kalendarz.selec

            //int NumerNr;
            for (int i = 0; i < ListaRejestrow.Count; i++)
            {
                int NumerNr = ImportClass.ZwrocNumerNr(ListaRejestrow[i]);

                if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                {
                    ImportClass.ImportZPlikuKonversja(NumerNr, ListaRejestrow[i]);
                }
                else
                {
                    ImportClass.NowyRaport(ListaRejestrow[i]);
                    if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
                    {
                        ImportClass.ImportZPlikuKonversja(NumerNr, ListaRejestrow[i]);
                    }
                }
            }
            OptimaCOM.O_Logout();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
