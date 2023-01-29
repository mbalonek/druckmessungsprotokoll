using System;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Configuration;
using Microsoft.Office.Interop.Word;

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
            CreateTemplateFile();

            try
            {
                if(!File.Exists(_protocolFilePath))
                    CopyTemplateFile();
            }
            catch (Exception ex)
            {
                logger.Log(ex.Message);
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
         * Creates template file
         * </summary>
         */
        private void CreateTemplateFile()
        {
            if (File.Exists(_templateFilePath))
            {
                return; 
            }
            var file = File.Create(_templateFilePath);
            file.Close();
            
            object read_only = false;
            object isVisible = false;
            wordApp.Visible = false;

            doc = wordApp.Documents.Open(_templateFilePath, read_only, isVisible);

            object missing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            
            Word.Paragraph header;

            header = doc.Content.Paragraphs.Add(ref missing);
            header.Range.Text = "Blutdruckmessungsprotokoll";
            header.Range.Font.Bold = 1;
            header.Range.Font.Size = 20;
            header.Range.Font.Name = "Bahnschrift";
            header.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            header.Range.InsertParagraphAfter();
            
            Word.Range wrdRng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            Word.Table table = doc.Tables.Add(wrdRng, 20, 5);
            table.Range.Font.Name = "Bahnschrift";
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            int r, c;
            for (r = 1; r <= 20; r++)
            {
                for (c = 1; c <= 5; c++)
                {
                    if (r == 1)
                    {
                        switch (c)
                        {
                            case 1:
                                table.Cell(r, c).Range.Text = "Datum / Uhrzeit";
                                break;
                            case 2:
                                table.Cell(r, c).Range.Text = "SYS mmHgg";
                                break;
                            case 3:
                                table.Cell(r, c).Range.Text = "DIA mmH";
                                break;
                            case 4:
                                table.Cell(r, c).Range.Text = "Puls 1/min";
                                break;
                            case 5:
                                table.Cell(r, c).Range.Text = "Pulsdruck mmHg";
                                break;                
                                        
                        }
                        table.Rows[1].Range.Font.Bold = 1;
                        table.Rows[1].Alignment = WdRowAlignment.wdAlignRowLeft;
                        table.Rows[1].Range.Font.Size = 12;
                        table.Rows[1].Range.Font.Color = WdColor.wdColorGreen;
                    }
                    else
                    {
                        switch (c)
                        {
                            case 1:
                                table.Cell(r, c).Range.Text = "<date>, <time>";
                                break;
                            case 2:
                                table.Cell(r, c).Range.Text = "<sys>";
                                break;
                            case 3:
                                table.Cell(r, c).Range.Text = "<dia>";
                                break;
                            case 4:
                                table.Cell(r, c).Range.Text = "<puls>";
                                break;
                            case 5:
                                table.Cell(r, c).Range.Text = "<pulsdruck>";
                                break;
                        }
                        table.Rows[r].Range.Font.Size = 11;
                        table.Rows[r].Alignment = WdRowAlignment.wdAlignRowLeft;
                        table.Rows[r].Range.Font.Color = WdColor.wdColorBlack;
                    }
                }

            }

            doc.Save();
            doc.Close();
        }

        /**
         * <summary>
         * Copies template file into user directory
         * </summary> 
         */
        private void CopyTemplateFile()
        {

            File.Copy(_templateFilePath, _protocolFilePath);
      
        }

        /**
         * <summary>
         * Finds text in Word Document and replace it with the values based on user input
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
         * Opens file with user's measurements and
         * calls method FindAndReplaceText() to add
         * new input data to the Word document
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

            object read_only = false;
            object isVisible = false;
            wordApp.Visible = false;
            try
            {
                doc = wordApp.Documents.Open(_protocolFilePath, ref read_only, ref isVisible);
                //doc.Activate();

                if (dp_dt.SelectedDate != null)
                {
                    FindAndReplaceText("<date>", dp_dt.SelectedDate.Value.ToShortDateString(), 1);
                    FindAndReplaceText("<time>", tb_time.Text, 1);
                    FindAndReplaceText("<sys>", tb_sys.Text, 1);
                    FindAndReplaceText("<dia>", tb_dia.Text, 1);
                    FindAndReplaceText("<puls>", tb_puls.Text, 1);
                    FindAndReplaceText("<pulsdruck>", tb_pulsdruck.Text, 1);

                    doc.Save();

                }
                else
                {
                    logger.Log("Input: Datumfeld ist leer");
                    MessageBox.Show("Feld Datum kann nicht leer sein");
                }
                
                doc.Close();

                MessageBox.Show("Datei wurde erstelt");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
         
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
            object read_only = false;
            object isVisible = false;
            wordApp.Visible = false;
            
            try
            {   
                doc = wordApp.Documents.Open(_protocolFilePath, ref read_only, ref isVisible);

                FindAndReplaceText("<date>, <time>", "", 2);
                FindAndReplaceText("<sys>", "", 2);
                FindAndReplaceText("<dia>", "", 2);
                FindAndReplaceText("<puls>", "", 2);
                FindAndReplaceText("<pulsdruck>", "", 2);

                doc.SaveAs2(_printFilePath, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing
                                        , ref missing, ref missing);
                doc.Close();
                doc = wordApp.Documents.Open(_printFilePath, ref read_only, ref isVisible);
                
                doc.PrintOut();
                
                Console.WriteLine("Datei bereit zu drucken");
                logger.Log("Datei bereit zu drucken");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Datei wurde nicht gedruckt");
                logger.Log(ex.Message);
               
            }
            finally
            {
                doc.Close();
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

        private void Window_Closed(object sender, EventArgs e)
        {
            wordApp.Quit();
        }
    }
}
