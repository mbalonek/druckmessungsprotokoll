using System;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;
using System.Configuration;

/*
 * application that allows user to record the results of arterial pressure measurements in a Word document. 
 */

namespace druckmessungsprotokoll
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
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

        private void PrepareDirectories()
        {
            Directory.CreateDirectory(_userDirectory);
            Directory.CreateDirectory(_appDirectory);
        }

        private void CopyTemplateFile()
        {
            File.Copy(_templateFilePath,_protocolFilePath);
        }

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
            object readOnly = false;
            object visible = true;
            object replace = replaceMode;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ExistingText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                                            ref matchSoundLike, ref matchAllForms, ref forward, ref wrap, ref format,
                                            ref ReplaceText, ref replace, ref matchKashida, ref matchDialectics,
                                            ref matchAlefHamza, ref matchControl);

        }

        private void AddMeasurement()
        {
           
            object missing = Missing.Value;


            if (!File.Exists(_protocolFilePath))
            {
                MessageBox.Show("Datei nicht gefünden");
                return;
            }
                object read_only = false;
                object isVisible = false;
                wordApp.Visible = false;

                doc = wordApp.Documents.Open(_protocolFilePath, ref read_only, ref isVisible);
                doc.Activate();
                var date = dp_dt.SelectedDate.Value.ToShortDateString();

                FindAndReplaceText("<date>", date, 1);
                FindAndReplaceText("<time>", tb_time.Text, 1);
                FindAndReplaceText("<sys>", tb_sys.Text, 1);
                FindAndReplaceText("<dia>", tb_dia.Text, 1);
                FindAndReplaceText("<puls>", tb_puls.Text, 1);
                FindAndReplaceText("<pulsdruck>", tb_pulsdruck.Text, 1);

                doc.SaveAs2(_protocolFilePath, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing
                                        , ref missing, ref missing);
                doc.Close();
                wordApp.Quit();
                MessageBox.Show("Datei wurde erstelt");

        }

        private void AddOne()
        {

            using (WordDocument document = new WordDocument())
            {
                //Opens the input Word document.
                Stream docStream = File.OpenRead(System.IO.Path.GetFullPath(@"protokoll_files\Blutdruckmessungsprotokoll.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();
                document.ReplaceFirst = true;
                //Finds all occurrences of a misspelled word and replaces with properly spelled word.
                document.Replace("<sys>", tb_sys.Text.ToString(), true, true);
                //Saves the resultant file in the given path.
                docStream = File.Create(System.IO.Path.GetFullPath(@"C:\Users\monbal\source\repos\mbalonek\druckmessungsprotokoll\druckmessungsprotokoll\Blutdruckmessungsprotokoll_temp3.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }

        private void AddMeassurementButtonClick (object sender, RoutedEventArgs e)
        {
            
            AddMeasurement();

        }

        public void PrintProtocol()
        {
            if (!File.Exists(_protocolFilePath))
            {
                MessageBox.Show("Datei nicht gefunden");
                return;
            }

            object missing = Missing.Value;
            object read_only = false;
            object isVisible = false;
            try
            {
                wordApp.Visible = false;
            }
            catch (Exception ex)
            {
                Logger.Log()
                MessageBox.Show("Drucker nicht gefünden");
            }
            

            doc = wordApp.Documents.Open(_printFilePath, ref read_only, ref isVisible);
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
            wordApp.Quit();

            MessageBox.Show("Datei wurde gedruckt");
        }
        
        private void PrintButtonClick(object sender, RoutedEventArgs e)
        {
            PrintProtocol();
        }
    }
}
