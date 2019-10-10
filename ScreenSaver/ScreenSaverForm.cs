using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Net;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using PdfSharp.Pdf.Security;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;

namespace ScreenSaver
{
    public enum OrientationPage { Portrait, Paysage};

    public partial class ScreenSaverForm : Form
    {
        //Variables 
        private System.Drawing.Point mouseLocation;                //Variable pour la localisation de la souris
        private List<FileName> list_files;      //Déclaration de la liste des documents à afficher
        private Random rand = new Random();         //Random
        private int intervale = 5000;
        private WebBrowser webBrowser1,webBrowser2;
        private int compteur = 0;
        private string folderDoc = @"E:\Projet Sanofi\Projet Ecran Communication\Dossier de test\";
        private string foldertmp = @"E:\Projet Sanofi\Projet Ecran Communication\Dossier de test\tmp\";
        private System.Drawing.Rectangle rec_active_screen;
        private int ToCenterWebWidth = 50, ToCenterWebWidth2 = 793;
        private int ToCenterWebHeight = 0, ToCenterWebHeight2 = 0;
        private int width_PDF;
        private int height_PDF;
        private Timer tmFading = new Timer();
        private int aa = 255;

        /// <summary>
        /// Methode qui SUPPRIME le dossier temporaire
        /// lorsque les documents ont été affichés
        /// les documents sont également détruits
        /// </summary>
        /// <param FileNoExtension="pathfilenametmp"></param>
        private void Destroy_Temp_Folder(string pathfilenametmp)
        {
            if (Directory.Exists(pathfilenametmp))
                Directory.Delete(pathfilenametmp, true);
        }


        /// <summary>
        /// Constructeur simple de la form
        /// </summary>
        public ScreenSaverForm()
        {
            InitializeComponent();
        }



        /// <summary>
        /// Constructeur avec un rectangle de la form
        /// </summary>
        /// <param FileNoExtension="bounds"></param>
        public ScreenSaverForm(System.Drawing.Rectangle bounds)
        {
            InitializeComponent();
            rec_active_screen = Screen.FromControl(this).Bounds;
         
            this.Controls.Add(this.webBrowser1);    
            this.Controls.Add(this.webBrowser2);

           // MessageBox.Show(Directory.Exists(foldertmp).ToString());
            if (Directory.Exists(foldertmp))
                Destroy_Temp_Folder(foldertmp);
        }
   


        /// <summary>
        /// Permet de récupérer la liste des fichiers dans le répertoire 
        /// des documents à afficher
        /// </summary>
        private List<FileName> Get_File_From_Path()
        {       
            list_files = new List<FileName>();
    
            var listFile = Directory.EnumerateFiles(folderDoc);

            if(listFile.Count() > 0)
            {
                foreach (var name in listFile)
                {
                    list_files.Add(new FileName { name = Path.GetFileName(name) });
                }
            }
        
            return list_files;
        }

   

        /// <summary>
        /// Methode principale 
        /// qui appelle la fonction de la liste des fichiers
        /// et affiche à l'écran chaque document
        /// </summary>
        private void Construit_Liste_Et_Affiche()
        {
          

            Get_File_From_Path();//Génère la liste des fichiers dans le répertoire

            tmFading.Enabled = true;
            tmFading.Tick += new EventHandler(FadeTimer_Tick);      
            tmFading.Start();

            //Charge chaque document de la liste, tant que la liste n'est pas vide
            if (list_files.Count > 0)
            {
                moveTimer.Tick += new EventHandler(moveTimer_Tick);
                moveTimer.Start();
            }
            else
            {
                MessageBox.Show("Il n'y a pas de fichier dans le répertoire !");
            }                      
        }



        /// <summary>
        /// Permet de charger l'écran de veille
        /// et d'afficher document par document à l'écran
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void ScreenSaverForm_Load(object sender, EventArgs e)
        {
            //Cursor.Hide();//Cache le curseur de la souris
            TopMost = true;//Affiche l'écran de veille en premier plan
          
            Construit_Liste_Et_Affiche();
        }



        /// <summary>
        /// Charge les documents 1 par 1 suivant avec un minuteur
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void moveTimer_Tick(object sender, EventArgs e)
        {
            SelectionTypeOfDocument();//Permet de séléctionner le type de traitement du document en fonction de son extension         
            compteur++;
            tmFading.Interval = intervale;
        }

        /// <summary>
        /// Effets de fondu à l'ouverture des documents
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void FadeTimer_Tick(object sender, EventArgs e)
        {        
           
            if (aa <= 0)
            {
                MessageBox.Show(aa.ToString());
                tmFading.Enabled = false;
                aa = 255;
                FadePanel.BackColor = System.Drawing.Color.FromArgb(aa, 0, 0, 0);

            }

            else
            {
                aa -= 5;
                FadePanel.BackColor = System.Drawing.Color.FromArgb(aa, 0, 0, 0);
            }
 
        }



        /// <summary>
        /// Permet de supprimer les fichiers dans le répertoire tmp
        /// </summary>
        /// <param FileNoExtension="pathfilenametmp"></param>
        private void Delete_Files_In_Tmp(string pathfilenametmp)
        {
            string[] filePaths = Directory.GetFiles(pathfilenametmp);
            foreach (string filePath in filePaths)
            {
                File.Delete(filePath);                         
            }                   
        }

 

        /// <summary>
        /// Methode qui affiche une page 
        /// dans le webbrowser
        /// </summary>
        /// <param name="fileName"></param>
        private void Affichage1(string fileName)
        {           
            //Charge le document dans le WebBrowser
            webBrowser1.Url = new Uri(fileName);

            //Le rend visible
            webBrowser1.Visible = true;
        }



        /// <summary>
        /// Methode qui affiche 2 pages
        /// dans le webbrowser
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileName2"></param>
        private void Affichage2(string fileName1,string fileName2)
        {           
            //On affiche les 2 pages à l'écran
            webBrowser1.Url = new Uri(fileName1);
            webBrowser2.Url = new Uri(fileName2);
           
            //Paramètres visibles pour les webrowsers
            webBrowser1.Visible = true;
            webBrowser2.Visible = true;
        }

      

        /// <summary>
        /// Methode qui CREE un dossier temporaire
        /// pour stocker les fichiers à une page
        /// et les afficher à l'écran
        /// </summary>
        /// <param FileNoExtension="filename"></param>
        private void CreateTmpFolder(string pathfilenametmp)
        {                     
            if(!Directory.Exists(pathfilenametmp))      
              Directory.CreateDirectory(pathfilenametmp);        
        }



        /// <summary>
        /// Fonction qui fait une copie du document
        /// à visionner dans le dossier tmp
        /// </summary>
        /// <param FileNoExtension="filepath"></param>
        /// <param FileNoExtension="filename"></param>
        private void CopiePDFToTmp(string fileorigin, string filedest)
        {
            //Produit une copie du document dans le dossier "tmp"      
            File.Copy(fileorigin, filedest, true);
        }



        /// <summary>
        /// Methode qui split un PDF avec plusieurs pages en plusieurs 
        /// fichiers avec 1 page chacun
        /// </summary>
        /// <param FileNoExtension="filename"></param>
        /// <param FileNoExtension="filepath"></param>
        private List<string> DecoupagePDF(string filename)
        {
            filename = Path.GetFileName(filename);
            string file_name1 = "", file_name2 = "";

            List<string> list_of_Split_Files = new List<string>();

            //Si le repertoire "tmp" n'existe pas, il est créé
            CreateTmpFolder(foldertmp);

            //Suprime les fichiers "fantomes"
            new List<string>(Directory.GetFiles(foldertmp)).ForEach(file => { if (file.ToUpper().Contains("$$".ToUpper())) File.Delete(file); });

            if (!File.Exists(foldertmp + filename))
            {
                //Copie du document dans le répertoire tmp
                CopiePDFToTmp(folderDoc + filename, foldertmp + filename);
            }
            PdfDocument outputDocument1 = new PdfDocument();
            PdfDocument outputDocument2 = new PdfDocument();
           
            // Ouvre le fichier dupliqué
            PdfDocument inputDocument = PdfReader.Open(foldertmp + filename, PdfDocumentOpenMode.Import);

            //Extrait le nom du fichier sans son extension
            string name = Path.GetFileNameWithoutExtension(filename);
           
            if (!File.Exists(String.Format("{0} - Page {1}.pdf", foldertmp + name, 1)))
            {
                //Génère les nouveaux documents à partir du nouveau fichier
             
                outputDocument1.ViewerPreferences.HideMenubar = true;
                outputDocument1.ViewerPreferences.HideToolbar = true;
                outputDocument1.ViewerPreferences.HideWindowUI = true;

                outputDocument1.AddPage(inputDocument.Pages[0]);
                
                //Ajoute une page et la sauvegarde

                file_name1 = String.Format("{0} - Page {1}.pdf", foldertmp +  name, 1);

                outputDocument1.Save(file_name1);
                list_of_Split_Files.Add(file_name1);
            }

            if (!File.Exists(String.Format("{0} - Page {1}.pdf", foldertmp +  name, 2)))
            {
                //Génère les nouveaux documents à partir du nouveau fichier
             
                outputDocument2.ViewerPreferences.HideMenubar = true;
                outputDocument2.ViewerPreferences.HideToolbar = true;
                outputDocument2.ViewerPreferences.HideWindowUI = true;
                //Ajoute une page et la sauvegarde
                outputDocument2.AddPage(inputDocument.Pages[1]);
                file_name2 = String.Format("{0} - Page {1}.pdf", foldertmp + name, 2);
         
                outputDocument2.Save(file_name2);
                list_of_Split_Files.Add(file_name2);
            }

            return list_of_Split_Files;
        }



        /// <summary>
        /// Methode pour transformer des documents PDF
        /// </summary>
        /// <param FileNoExtension="totalPages"></param>
        /// <param FileNoExtension="filepath"></param>
        /// <param FileNoExtension="filename"></param>
        private void TraitementPDF(string filename)
        {
            //Longueur et largeur d'un PDF en rapport avec la résolution de l'écran MIMIO (1080 x 1920)
             width_PDF = 750;
             height_PDF = 1080;
            string fileNoextension = "";
            ToCenterWebWidth = 435;
            string fileToOpen = filename;
            string copyOfFile = foldertmp + Path.GetFileName(filename);
           // MessageBox.Show(copyOfFile);
            // Ouvre le fichier dupliqué
            PdfDocument inputDocument = PdfReader.Open(fileToOpen, PdfDocumentOpenMode.Import);

            int totalPages = inputDocument.PageCount;

            switch (totalPages)
            {
                case 1:
                   
                    //793 pix = 21 cm et 1123 pix = 29,70 cm, taille d'un format A4
                    webBrowser1.SetBounds(ToCenterWebWidth, ToCenterWebHeight, width_PDF, height_PDF);
                    if (!File.Exists(copyOfFile))
                    CopiePDFToTmp(fileToOpen, copyOfFile);
                    filename = Path.GetFileNameWithoutExtension(filename);
                    fileNoextension = String.Format("{0}Copy.pdf", filename);
                    copyOfFile = foldertmp + fileNoextension;              
                    PdfDocument outputDocument1 = new PdfDocument();

                    if (!File.Exists(copyOfFile))
                    {
                        outputDocument1.ViewerPreferences.HideMenubar = true;
                        outputDocument1.ViewerPreferences.HideToolbar = true;
                        outputDocument1.ViewerPreferences.HideWindowUI = true;
                        outputDocument1.AddPage(inputDocument.Pages[0]);
                       
                        //Sauvgarde la copie du document d'origine
                        outputDocument1.Save(copyOfFile);
                    }
            
                    Affichage1(copyOfFile);
             
                    break;

                //Plus d'une page
                default:
                  
                    ToCenterWebWidth = 50;
                    List<string> list_files_splited = new List<string>();
                           
                    list_files_splited = DecoupagePDF(filename);  //On appelle la fonction qui subdivise les pages en plusieurs documents

                    //On stock chacune des pages dans une variable qui les localise
                    string doc1 = list_files_splited[0];
                    string doc2 = list_files_splited[1];

                    //793 pix = 21 cm et 1123 pix = 29,70 cm, taille d'un format A4
                    webBrowser1.SetBounds(ToCenterWebWidth, ToCenterWebHeight, width_PDF, height_PDF);
                    webBrowser2.SetBounds(ToCenterWebWidth2, ToCenterWebHeight2, width_PDF, height_PDF);
         
                    Affichage2(doc1, doc2);

                    break;
            }
        }



        /// <summary>
        /// Methode pour transforme des documents PDF
        /// </summary>
        /// <param FileNoExtension="totalPages"></param>
        private void TraitementURL(string filename)
        {     
            webBrowser1.SetBounds(0,0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            string LinkToOpen = folderDoc + filename;
      
            using (WebClient client = new WebClient())
            {
                //Récupère le contenu du lien .URL et "Split" l'adresse web qui nous interresse
                string s = client.DownloadString(LinkToOpen);
                var avantURL = s.Split(new[] { "URL=" }, StringSplitOptions.None)[1];
                var apresURL = avantURL.Split(new[] { "\n" }, StringSplitOptions.None)[0];

                Affichage1(apresURL);
            }                     
        }



        /// <summary>
        /// Methode qui transforme les fichiers PPT
        /// </summary>
        /// <param FileNoExtension="filePath"></param>
        /// <param FileNoExtension="fileName"></param>
        private void TraitementPPT(string fileName)
        {
            //Convertion du PPT ou PPTX en PDF
          try
            {
                Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();

                Microsoft.Office.Interop.PowerPoint.Presentation presentation = ppApp.Presentations.Open(folderDoc + fileName,
                    Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);

                string fileToConvert = foldertmp + string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(fileName));
 
                //Export du fichier sous format pdf dans tmp
                if (!File.Exists(fileToConvert))
                {
                        presentation.ExportAsFixedFormat(fileToConvert,
                  PpFixedFormatType.ppFixedFormatTypePDF, PpFixedFormatIntent.ppFixedFormatIntentPrint, Microsoft.Office.Core.MsoTriState.msoFalse,
                  PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst, PpPrintOutputType.ppPrintOutputSlides, Microsoft.Office.Core.MsoTriState.msoFalse,
                  null, PpPrintRangeType.ppPrintAll, "", false, true, true, true, false, System.Reflection.Missing.Value);
                }

                //Ouverture du nouveau PDF généré à partir du PPT ou PPTX
                TraitementPDF(foldertmp + string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(fileName)));

                presentation.Close();
                presentation = null;
                ppApp = null;

                GC.Collect();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }        
        }



        /// <summary>
        /// Methode qui transforme les fichiers XLS ou XLSX
        /// </summary>
        /// <param FileNoExtension="filePath"></param>
        /// <param FileNoExtension="fileName"></param>
        private void TraitementXLS(string fileName)
        {
            try
            {
                //Convertion du fichier XLS ou XLSX en fichier PDF
                Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application()
                {
                    Visible = false,
                    ScreenUpdating = false
                };

                Workbook excelDocument = appExcel.Workbooks.Open(folderDoc + fileName);

                string fileToOpen = foldertmp + string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(fileName));

                if (!File.Exists(fileToOpen))
                {
                    excelDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, fileToOpen);
                }

                excelDocument.Close();
                appExcel.Quit();

                TraitementPDF(foldertmp + string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(fileName)));

                GC.Collect();
            } 
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }           
        }



        /// <summary>
        /// Permet de séléctionner le traitement qui convient 
        /// en fonction de l'extension du document
        /// </summary>     
        private void SelectionTypeOfDocument()
        {          
            webBrowser1.Visible = false;
            webBrowser2.Visible = false;

            CreateTmpFolder(foldertmp);//Si le dossier tmp n'existe pas, on le créer
           
                
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            switch (Path.GetExtension(list_files[compteur].name))
                {
                    case ".pdf":
                        TraitementPDF(folderDoc + list_files[compteur].name);
                        break;          

                    case ".xls":
                        TraitementXLS(list_files[compteur].name);
                        break;

                    case ".xlsx":
                        TraitementXLS(list_files[compteur].name);
                        break;

                    case ".ppt":            
                        TraitementPPT(list_files[compteur].name);
                        break;

                    case ".pptx":
                        TraitementPPT(list_files[compteur].name);
                        break;

                    case ".url":
                        TraitementURL(list_files[compteur].name);
                        break;
                }
   
            moveTimer.Interval = intervale;                       
        }    



        /// <summary>
        /// Test si la page est bien chargée
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void BrowserDocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.AbsolutePath != (sender as WebBrowser).Url.AbsolutePath)
                return;
        }



        /// <summary>
        /// Permet de detecter si il y a un mouvement de la souris
        /// et si oui, de sortir de l'application
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void ScreenSaverForm_MouseMove(object sender, MouseEventArgs e)
        {        
            if (!mouseLocation.IsEmpty)
            {
                // Termine l'écran de veille si la souris bouge
                if (Math.Abs(mouseLocation.X - e.X) > 5 || Math.Abs(mouseLocation.Y - e.Y) > 5)
                {     
                    this.Close();
                    System.Windows.Forms.Application.Exit();      
                }
            }

            // Update de la position de la souris
            mouseLocation = e.Location;
        }



        /// <summary>
        /// Permet de detecter si il y a un click de souris
        /// et si oui, de sortir de l'application
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void ScreenSaverForm_MouseClick(object sender, MouseEventArgs e)
        {
            this.Close();
            System.Windows.Forms.Application.Exit();
      
        }



        /// <summary>
        /// Permet de detecter si il y a une pression sur une touche
        /// et si oui, de sortir de l'application
        /// </summary>
        /// <param FileNoExtension="sender"></param>
        /// <param FileNoExtension="e"></param>
        private void ScreenSaverForm_KeyPress(object sender, KeyPressEventArgs e)
        {
            this.Close();
            System.Windows.Forms.Application.Exit();

        }
    }
}