using System;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Configuration;
using System.Runtime.InteropServices;

namespace druckmessungsprotokoll
{
    /** 
     * <summary>
     * Interaction logic for MainWindow.xaml
     * </summary>
     */ 
    public partial class MainWindow : System.Windows.Window
    {
        private string _userDirectory => Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Blutdruckmessungen");
        private string _appDirectory => Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BlutdruckmessungApp");
        private string _templateFilePath => Path.Combine(_appDirectory,
            ConfigurationManager.AppSettings.Get("template_file_name")!);

        private string _protocolFilePath => Path.Combine(_userDirectory,
            ConfigurationManager.AppSettings.Get("protocol_file_name")!);
        
        private string _printFilePath => Path.Combine(_userDirectory,
            ConfigurationManager.AppSettings.Get("print_file_name")!);

        private Word.Application wordApp = new Word.Application();
        private Word.Document doc = new Word.Document();
        
        private Logger logger = new Logger();

        public MainWindow()
        {
            InitializeComponent();
            PrepareDirectories();

            try
            {
                if(!File.Exists(_protocolFilePath))
                    CopyTemplateFile();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }
        /**
         * <summary>
         * Sets the working directories for application as well for the user
         * </summary>
         */
        private void PrepareDirectories()
        {
            Directory.CreateDirectory(_userDirectory);
            Directory.CreateDirectory(_appDirectory);
        }

        /**
         * <summary>
         * Copies template file into user directory
         * </summary> 
         */
        private void CopyTemplateFile()
        {
            File.Copy(_templateFilePath,_protocolFilePath);
        }

        /**
         * <summary>
         * Finds text in Word Document and replace it with the value from user input
         * </summary>
         */
        private void FindAndReplaceText(object ExistingText, object ReplaceText, int replaceMode)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object matchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDialectics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object readOnly = true;
            object visible = true;
            object replace = replaceMode;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ExistingText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                                            ref matchSoundLike, ref matchAllForms, ref forward, ref wrap, ref format,
                                            ref ReplaceText, ref replace, ref matchKashida, ref matchDialectics,
                                            ref matchAlefHamza, ref matchControl);

        }
        
        /**
         * <summary>
         * Adds input data to the Word document
         * </summary>
         */
        private void AddMeasurement()
        {
           
            object missing = Missing.Value;

            if (!File.Exists(_protocolFilePath))
            {
                MessageBox.Show("Datei nicht gefünden");
                return;
            }
                
            object read_only = true;
            object isVisible = true;
            wordApp.Visible = true;

            doc = wordApp.Documents.Open(_protocolFilePath, ref read_only, ref isVisible);
            doc.Activate();

            if (dp_dt.SelectedDate != null)
            {
                FindAndReplaceText("<date>", dp_dt.SelectedDate.Value.ToShortDateString(), 1);
                FindAndReplaceText("<time>", tb_time.Text, 1);
                FindAndReplaceText("<sys>", tb_sys.Text, 1);
                FindAndReplaceText("<dia>", tb_dia.Text, 1);
                FindAndReplaceText("<puls>", tb_puls.Text, 1);
                FindAndReplaceText("<pulsdruck>", tb_pulsdruck.Text, 1); 
                    
            }
            else
            {
                logger.Log("Input: Datumfeld ist leer");
                MessageBox.Show("Feld Datum kann nicht leer sein");
            }
            doc.Save();
            
            doc.Close();
            
            MessageBox.Show("Datei wurde erstelt");

        }
        /**
         * <summary>
         * Prints out existing file
         * </summary>
         */
        public void PrintProtocol()
        {
            if (!File.Exists(_protocolFilePath))
            {
                MessageBox.Show("Datei nicht gefunden");
                return;
            }

            object missing = Missing.Value;
            object read_only = true;
            object isVisible = true;
            wordApp.Visible = false;
            
            try
            {   
                doc = wordApp.Documents.Open(_protocolFilePath, ref read_only, ref isVisible);
                doc.Activate();

                FindAndReplaceText("<date>, <time>", "", 2);
                FindAndReplaceText("<sys>", "", 2);
                FindAndReplaceText("<dia>", "", 2);
                FindAndReplaceText("<puls>", "", 2);
                FindAndReplaceText("<pulsdruck>", "", 2);

                doc.SaveAs2(_printFilePath, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing
                                        , ref missing, ref missing);

                doc.PrintOut();
                doc.Close();
                
                Console.WriteLine("Datei bereit zu drucken");
                logger.Log("Datei bereit zu drucken");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Datei wurde nicht gedruckt");
                logger.Log(ex.Message);
               
            }      
        }

        private void AddMeassurementButtonClick(object sender, RoutedEventArgs e)
        {
            AddMeasurement();
        }

        private void PrintButtonClick(object sender, RoutedEventArgs e)
        {
            PrintProtocol();
        }
    }
}
