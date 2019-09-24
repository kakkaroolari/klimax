using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using Ookii.Dialogs.Wpf;
using ClosedXML.Excel;

namespace excel2kli
{
    public partial class Form1 : Form
    {
        private const string WUFIDT = "d.M.yy_H.00";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dlg = new VistaOpenFileDialog
            {
                CheckFileExists = true,
            };
            var ret = dlg.ShowDialog();
            if (ret.HasValue && ret.Value /* == System.Windows.Forms.DialogResult.OK*/)
            {
                string filePath = dlg.FileName;
                Debug.WriteLine($"[ERITE] exceli faili: {filePath}");

                LisaaTekstia("Raahaa sarakkeet oikeille paikoilleen.");
                LisaaTekstia("| aika | sade [Ltr/m²h] | säteily [W/m²] | T_ulko [°C] | RH_ulko (0..1) | T_sisä [°C] | RH_ulko (0..1) | paine [hPa] |");

                //Open the Excel file using ClosedXML.
                using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    //Create a new DataTable.
                    DataTable dt = new DataTable();

                    //Loop through the Worksheet rows.
                    bool firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        //Use the first row to add columns to DataTable.
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            //Add rows to DataTable.
                            dt.Rows.Add();
                            int i = 0;
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                i++;
                            }
                        }
                    }
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    dataGridView1.DataSource = dt;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.richTextBox1.Text = "Avaa exceli filu napista tai raahaa se tähän päälle.";
        }

        private void LisaaTekstia(string tekstia)
        {
            this.richTextBox1.Text += Environment.NewLine + tekstia;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            // set the current caret position to the end
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            // scroll it automatically
            richTextBox1.ScrollToCaret();
        }

        // jyysta klikki
        private void button2_Click(object sender, EventArgs e)
        {
            // input check
            var taulukko = dataGridView1.DataSource as DataTable;
            var riveja = dataGridView1.Rows.Count;

            if (null == taulukko || 0 == riveja)
            {
                NaytaVirhe("Laatikosta ei tule dataa.");
            }

            // aika
            var rivilista = new List<klimalli>();
            int sarakeNro = 0;
            
            try
            {
                for (int riviNro = 0; riviNro < riveja; riviNro++)
                {
                    var rivi = dataGridView1.Rows[riviNro];
                    sarakeNro = 0;
                    var malli = new klimalli
                    {
                        jarjestysnumero = riviNro + 1,
                        aika = Convert.ToDateTime(rivi.Cells[sarakeNro++].Value),
                        sade = Convert.ToDecimal(rivi.Cells[sarakeNro++].Value),
                        sateily = Convert.ToDecimal(rivi.Cells[sarakeNro++].Value),
                        T_e = Convert.ToDecimal(rivi.Cells[sarakeNro++].Value),
                        RH_e = ToAbsoluteValueBetween0And1(rivi.Cells[sarakeNro++].Value),
                        T_i = Convert.ToDecimal(rivi.Cells[sarakeNro++].Value),
                        RH_i = ToAbsoluteValueBetween0And1(rivi.Cells[sarakeNro++].Value),
                    };
                    rivilista.Add(malli);
                }
            }
            catch(Exception ex)
            {
                NaytaVirhe($"Lukuvirhe sarakkeessa {sarakeNro}, '{ex.Message}'");
            }

            // TODO: nayta virhe jos eri kuin yhden tunnin aikavali
            var alku = rivilista[0].aika;
            var loppu = rivilista[rivilista.Count-1].aika;

            // write file
            var kliifile = GetTempFilePathWithExtension(".kli");
            try
            {
                //Open the File
                using (FileStream fs = File.Create(kliifile))
                using (StreamWriter sw = new StreamWriter(fs, Encoding.ASCII))
                {
                    sw.WriteLine($"$WUFI$\t{alku.ToString(WUFIDT)} - {loppu.ToString(WUFIDT)}");
                    sw.WriteLine("Azimut : 90, Neigung : 90  Abgeleitet von IBP1991.WET");
                    sw.WriteLine(@"C:\WIN16APP\WUFI\WET\IBP1991.WET                                                47.9   11.7   680.0  90     90     1      0      21     1      0.45   0.15   1      3.06_12.0       16.08_12.0      0  0  1  1  ");

                    foreach(var klirivi in rivilista)
                    {
                        sw.WriteLine(klirivi);
                    }

                    //close the file
                    sw.Close();
                }
                LisaaTekstia($"Tallennettiin tonne: '{kliifile}'");
                System.Diagnostics.Process.Start("notepad.exe", kliifile);
            }
            catch (Exception ex)
            {
                NaytaVirhe("Sori, tallennus ei toimi: " + ex.Message);
            }
            finally
            {
                Debug.WriteLine("Executing finally block.");
                Form1_Load(null, null);
            }
        }

       private decimal ToAbsoluteValueBetween0And1(object value)
       {
          var inPercent = Convert.ToDecimal(value);
          var asAbs = inPercent / 100;
          return Math.Max(0, Math.Min(asAbs, 1));
       }

       private void NaytaVirhe(string tekstia)
        {
            //
            // Dialog box with exclamation icon.
            //
            MessageBox.Show(tekstia,
                "Eeppinen virhe",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation,
                MessageBoxDefaultButton.Button1);
        }

        private static string GetTempFilePathWithExtension(string extension)
        {
            var path = Path.GetTempPath();
            var fileName = Guid.NewGuid().ToString() + extension;
            return Path.Combine(path, fileName);
        }
    }
}
