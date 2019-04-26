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
            
            
            var calTab = db.Calcules.ToList();
            Calcule cal = calTab[0];
            var RpaList = new List<RPA>();
            string[] dirs = Directory.GetDirectories(@"U:\DIL_INNOVATION_ET_RELAIS_CROISSANCES\INNO_Aurelien\00_Factory\Outils_PtfProjet\Rpa_Excel_Pilotage\");
            int SommeGlobal = 0;
            foreach (string dir in dirs)
            {
                string[] fils = Directory.GetFiles(dir, "C*");
                foreach (string file in fils)
                {
                    if (File.Exists(file))
                    {
                    
                        string RpaName = new DirectoryInfo(System.IO.Path.GetDirectoryName(file)).Name;
                        DateTime dateCreation = System.IO.File.GetCreationTime(file);
                        if (dateCreation > cal.LastDate)
                        {
                        
                            var rpaTable = db.RPAs.SqlQuery("SELECT *  FROM RPAs WHERE Nom=@Nom", new SqlParameter("@Nom", RpaName.ToLower())).ToList();

                            RPA rpa = rpaTable[0];
                            //ouverture de fichier Excel et recup de nb
                            xlApp = new Excel.Application();
                            xlWorkBook = xlApp.Workbooks.Open(file);
                            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets["Synthèse"];
                            range = xlWorkSheet.UsedRange;
                            Double nb1 = (range.Cells[Int32.Parse(rpa.ligne), Int32.Parse(rpa.colonne)] as Excel.Range).Value;
                            int nb = Convert.ToInt32(nb1);
                            int SommeRpa = nb * Int32.Parse(rpa.minutes);
                            //fin de calcule

                            SommeGlobal += SommeRpa;
                            SommeRpa += rpa.SommeMin;
                            rpa.SommeMin = SommeRpa;
                            db.Entry(rpa).State = EntityState.Modified;

                            db.SaveChanges();
                        }
                    }
                }
                //recuperer la date de creation

            }
            cal.LastDate = DateTime.Now;
            db.Entry(cal).State = EntityState.Modified;
            
            cal.Sommes += SommeGlobal;
            double nbj = cal.Sommes / 468.0;
            db.Entry(cal).State = EntityState.Modified;
            db.SaveChanges();

            RpaList = db.RPAs.ToList();
            /********************************************Creation du PDF*************************************/
            string html = "<html><body><img src='file:///P:/logoCA.jpg' Height='105' Width='110'/> <h2 style ='text-align:center; font-family= \'Trebuchet MS\', \'Lucida Sans Unicode\', \'Lucida Grande\', \'Lucida Sans\', Arial, sans-serif;color=darkslategrey;'>Récapitulatif</h2>" +
                "<h1 style = 'text-align:center; color:SeaGreen ' ;font-family='\'Trebuchet MS_', \'Lucida Sans Unicode\', \'Lucida Grande\', \'Lucida Sans\', Arial, sans-serif' >" +  Math.Round(nbj, 1, MidpointRounding.ToEven) + " J/H</h1>" +
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
            byte[] pdfContent = new SimplePechkin(new GlobalConfig()).Convert(html);
            try
            {
                File.WriteAllBytes(@"U:\DIL_INNOVATION_ET_RELAIS_CROISSANCES\INNO_Aurelien\00_Factory\Outils_PtfProjet\Sommes.pdf", pdfContent);
                
            }catch(InvalidCastException )
            {
                System.Windows.MessageBox.Show("Merci de fermer le pdf ");
            }


            System.Windows.MessageBox.Show("Calcule effectué ");
















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
            db.RPAs.Add(rpa);
            db.SaveChanges();
            System.Windows.MessageBox.Show("RPA bien ajouté");
        }

        private void RpaNom_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged_2(object sender, TextChangedEventArgs e)
        {

        }

        private void Min_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void RadioButton_Checked_1(object sender, RoutedEventArgs e)
        {

        }

        private void Periode1_Checked(object sender, RoutedEventArgs e)
        {

        }


    }
}
