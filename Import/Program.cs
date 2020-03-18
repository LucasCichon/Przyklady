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
        private Label label1;

        public static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public Form1()
        {
            InitializeComponent();
        }

        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            //OptimaCOM.O_Login("ADMIN", "", "Firma_Demo");

            //List<Rejestr> ListaRejestrow = ImportClass.SzukaniePlikow();

            ////int NumerNr;
            //for(int i=0; i<ListaRejestrow.Count; i++)
            //{
            //    int NumerNr = ImportClass.ZwrocNumerNr(ListaRejestrow[i]);

            //        if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
            //            {
            //                ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
            //            }
            //        else
            //        {
            //            ImportClass.NowyRaport(ListaRejestrow[i]);
            //            if (ImportClass.czyIstniejeRaport(ListaRejestrow[i].Data, ListaRejestrow[i].Numer, ListaRejestrow[i]) == true)
            //                {
            //                    ImportClass.ImportZPliku(NumerNr, ListaRejestrow[i]);
            //                }
            //        }
            //}
            //OptimaCOM.O_Logout();
        }

        private void InitializeComponent()
        {
            this.Kalendarz = new System.Windows.Forms.MonthCalendar();
            this.BtnImportZapisow = new System.Windows.Forms.Button();
            this.listBoxLog = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Kalendarz
            // 
            this.Kalendarz.Location = new System.Drawing.Point(18, 18);
            this.Kalendarz.MaxSelectionCount = 1;
            this.Kalendarz.Name = "Kalendarz";
            this.Kalendarz.TabIndex = 0;
            // 
            // BtnImportZapisow
            // 
            this.BtnImportZapisow.Location = new System.Drawing.Point(90, 229);
            this.BtnImportZapisow.Name = "BtnImportZapisow";
            this.BtnImportZapisow.Size = new System.Drawing.Size(105, 36);
            this.BtnImportZapisow.TabIndex = 1;
            this.BtnImportZapisow.Text = "Import";
            this.BtnImportZapisow.UseVisualStyleBackColor = true;
            this.BtnImportZapisow.Click += new System.EventHandler(this.BtnImportZapisow_Click);
            // 
            // listBoxLog
            // 
            this.listBoxLog.FormattingEnabled = true;
            this.listBoxLog.Location = new System.Drawing.Point(299, 18);
            this.listBoxLog.Name = "listBoxLog";
            this.listBoxLog.Size = new System.Drawing.Size(472, 277);
            this.listBoxLog.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 202);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "label1";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(794, 313);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxLog);
            this.Controls.Add(this.BtnImportZapisow);
            this.Controls.Add(this.Kalendarz);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Name = "Form1";
            this.Text = "Importowanie Zapisów KB";
            this.ResumeLayout(false);
            this.PerformLayout();

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

        public void logURI(/*string OutputLog, string Information, string JOB,*/ string s)
        {
            try
            {
                listBoxLog.Items.Add(s);
                //listBox1.BeginUpdate();
                //listBox1.Items.Add("0");
                //listBox1.Items[0] = DateTime.Now.ToString() + " : " + JOB + " " + Information;
                //listBox1.Items.Add("1");
                //listBox1.EndUpdate();
                //textBox1.Text = JOB;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            label1.Text = Kalendarz.SelectionRange.Start.ToShortDateString();
            //label1.Text = Kalendarz.SelectionRange.ToString();
        }
    }
}
