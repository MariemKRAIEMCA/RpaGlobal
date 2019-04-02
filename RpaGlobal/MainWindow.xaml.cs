using Pechkin;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace RpaGlobal
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        private MyDbContext db = new MyDbContext();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            Calcule cal = db.Calcule.Last<Calcule>();
            var RpaList = new List<RPA>();
            string[] dirs = Directory.GetDirectories(@"U:\DIL_INNOVATION_ET_RELAIS_CROISSANCES\INNO_Aurelien\00_Factory\Outils_PtfProjet\Rpa_Excel_Pilotage\");
            int SommeGlobal = 0;
            foreach (string dir in dirs)
            {
                //recuperer la date de creation
                DateTime dateCreation = System.IO.File.GetCreationTime(dir);
                if(dateCreation > cal.LastDate)
                {
                    string RpaName = dir.Substring(20).ToLower();
                    var rpaTable = db.Rpa.SqlQuery("SELECT * FROM Rpa WHERE Nom=@Nom", new SqlParameter("@Nom", RpaName)).ToList();
                    RPA rpa = rpaTable[1]; 
                    //ouverture de fichier Excel et recup de nb
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(dir);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Synthèse"];
                    range = xlWorkSheet.UsedRange;
                    int nb = range.Cells[Int32.Parse(rpa.ligne), Int32.Parse(rpa.colonne)];
                    int SommeRpa = nb * Int32.Parse(rpa.minutes);
                    //fin de calcule

                    SommeRpa += rpa.SommeMin;
                    SommeGlobal += SommeRpa;
                    rpa.SommeMin = SommeRpa;
                    db.Entry(rpa).State = EntityState.Modified;
                    cal.LastDate = DateTime.Now;
                    db.Entry(cal).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }

            cal.Sommes += SommeGlobal;
            db.Entry(cal).State = EntityState.Modified;
            db.SaveChanges();

            RpaList = db.Rpa.ToList();
            /********************************************Creation du PDF*************************************/
            string html = "<html><body><h2 style = 'text-align:center'>récapitulatif</h2>" +
                "<h3>Somme en jourH pour Rpa Factory" + cal.Sommes / 468 + "</h3>" +
                "<table class='table table-striped' style='width: 100%;max-width: 100%;margin-bottom: 20px;margin-top: 20%;border-spacing:0;border-collapse: collapse;'>" +
                "<thead style = 'display: table-header-group;vertical-align: middle;border-color: inherit;' >" +
                "<tr style='display: table-row;vertical-align: inherit;border-color:inherit;'>" +
                "<th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Rpa Nom</th>" +
                " <th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Somme(min)</th>" +
                " <th style = 'vertical-align: line-height: 1.42857143;padding: 8px;bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='tableauTitre'>Unité(J,S,M)</th>";
            foreach (var rpa in RpaList)
            {
                html += "<tr style='display: table-row;vertical-align: inherit;border-color:inherit;'>" +
                    "<td style='padding:8px;line-height:1.42857143;border-top:1px; border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + rpa.Nom + "</td>" +
                    "<td style='padding:8px;line-height:1.42857143;border-top:1px; border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + rpa.SommeMin + "</td>" +
                    "<td style='padding:8px;line-height:1.42857143;border-top:1px; border-left:3px; solid #ddd;vertical-align: bottom;border-bottom: 2px solid #ddd;font-size: 15px;text-align: center;' class='elementTableau'>" + rpa.periode + "</td>";
            }
            html+= "</tbody></table><footer style='position:absolute;bottom:0;width:100%;height:60px;'></footer></body></html>";
            //byte[] pdfContent = new SimplePechkin(new GlobalConfig()).Convert(html);

            

















        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            RPA rpa = new RPA();
            rpa.Nom = RpaNom.Text;
            rpa.ligne = Ligne.Text;
            rpa.colonne = Col.Text;
            rpa.minutes = Min.Text;
            rpa.SommeMin = 0;
            var radiobtn = sender as RadioButton;
            if(Periode.IsChecked == true)
            {
                rpa.periode = "j";
            }else if(Periode.IsChecked == true)
            {
                rpa.periode = "s";
            }
            else
            {
                rpa.periode = "m";
            }
            db.Rpa.Add(rpa);
            db.SaveChanges();
            System.Windows.MessageBox.Show("RPA bien ajouté");
        }
    }
}
