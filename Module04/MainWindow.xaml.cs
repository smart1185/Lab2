using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using Microsoft.Win32;
using GemBox;
using System.Diagnostics;
using GemBox.Document;
using Paragraph = System.Windows.Documents.Paragraph;
using Run = System.Windows.Documents.Run;

namespace Module04
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            ((Run)FirstParagraf.Inlines.FirstInline).Text = "Change Text";
            
            // Set license key to use GemBox.Document in a Free mode.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Stop reading/writing a document when the free limit is reached.
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.Stop;




            //SoftwareInteraction



            //Создание жирного теста
            Bold bold = new Bold();
            Run runBold = new Run();
            runBold.Text = "В эмиграции";
            bold.Inlines.Add(runBold);

            //Создание первой части фразы
            Run runFirst = new Run();
            runFirst.Text = "В Европу Герцен приехал, настроенный скорее радикально-республикански, чем социалистически, хотя начатая им публикация в «Отечественных записках» серии статей под заглавием «Письма с Avenue Marigny» (впоследствии в переработанном виде опубликованы в «Письмах из Франции и Италии») шокировала его друзей — либералов-западников — своим антибуржуазным пафосом. ";

            //Создание последней фразы
            Run runLast = new Run();
            runLast.Text = "Последовавшее затем Июньское восстание рабочих, его кровавое подавление и наступившая реакция потрясли Герцена, который решительно обратился к социализму.";

            //Добавлениезаголовка
            Paragraph paragraphHead = new Paragraph();
            paragraphHead.Inlines.Add(bold);

            //Добавление 2-хчастей в обзац по порядку
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(runFirst);
            paragraph.Inlines.Add(runLast);

            FlowDocument doument = new FlowDocument();
            doument.Blocks.Add(paragraphHead);
            doument.Blocks.Add(paragraph);

            //Выводдокумента
            DocViewer.Document = doument;


            using (FileStream fs = File.Open("documGertsen2.xaml", FileMode.Open))
            {

                FlowDocument document = XamlReader.Load(fs) as FlowDocument;

                if (document == null)
                {
                    MessageBox.Show("Ошибка при загрузке документа");
                }
                else
                {
                    FlowDocumentConainer.Document = document;
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*";

            if (openFile.ShowDialog() == true)
            {
                TextRange documentTextRange = new TextRange(
                    MyRichTextBox.Document.ContentStart, 
                    MyRichTextBox.Document.ContentEnd);

                using (FileStream fs = File.Open(openFile.FileName, FileMode.Open))
                {
                    documentTextRange.Load(fs, DataFormats.Rtf);
                }
            }



            /*


            var dialog = new Microsoft.Win32.OpenFileDialog()
            {
                AddExtension = true,
                Filter =
                    "All Documents (*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.htm;*.html;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.txt|" +
                    "Word Documents (*.docx)|*.docx|" +
                    "Word Macro-Enabled Documents (*.docm)|*.docm|" +
                    "Word 97-2003 Documents (*.doc)|*.doc|" +
                    "Word Templates (*.dotx)|*.dotx|" +
                    "Word Macro-Enabled Templates (*.dotm)|*.dotm|" +
                    "Word 97-2003 Templates (*.dot)|*.dot|" +
                    "Web Pages (*.htm;*.html)|*.htm;*.html|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Text Files (*.txt)|*.txt"
            };

            if (dialog.ShowDialog() == true)
                using (var stream = new MemoryStream())
                {
                    // Convert input file to RTF stream. 
                   DocumentModel.Load(dialog.FileName).Save(stream, SaveOptions.RtfDefault);

                    stream.Position = 0;

                    // Load RTF stream into RichTextBox. 
                    var textRange = new TextRange(MyRichTextBox.Document.ContentStart, MyRichTextBox.Document.ContentEnd);
                    textRange.Load(stream, DataFormats.Rtf);
                }
                */
            //OpenFileDialog openDialog = new OpenFileDialog();
            //openDialog.Filter = "RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*";

            //if (openDialog.ShowDialog() == true)
            //{
            //    TextRange documentTextRange = new TextRange(MyRichTextBox.Document.ContentStart, MyRichTextBox.Document.ContentEnd);

            //    using (FileStream fs = File.Open(openDialog.FileName, FileMode.Open))
            //    {
            //        documentTextRange.Load(fs, DataFormats.Rtf);
            //    }
            //}
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "RichText Files (*.rtf)|*.rtf|All Files (*.*)|*.*";

            if (saveFile.ShowDialog() == true)
            {
                TextRange documntTextRange = 
                    new TextRange(MyRichTextBox.Document.ContentStart, 
                    MyRichTextBox.Document.ContentEnd);
                using (FileStream fs = File.Create(saveFile.FileName))
                {
                    documntTextRange.Save(fs, DataFormats.Rtf);
                }
            }
            /* SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter =
                "All Documents (*.docx;*.docm;*.doc;*.dotx;*.dotm;*.dot;*.htm;*.html;*.rtf;*.txt)|*.docx;*.docm;*.dotx;*.dotm;*.doc;*.dot;*.htm;*.html;*.rtf;*.txt|" +
                "Word Documents (*.docx)|*.docx|" +
                "Word Macro-Enabled Documents (*.docm)|*.docm|" +
                "Word 97-2003 Documents (*.doc)|*.doc|" +
                "Word Templates (*.dotx)|*.dotx|" +
                "Word Macro-Enabled Templates (*.dotm)|*.dotm|" +
                "Word 97-2003 Templates (*.dot)|*.dot|" +
                "Web Pages (*.htm;*.html)|*.htm;*.html|" +
                "Rich Text Format (*.rtf)|*.rtf|" +
                "Text Files (*.txt)|*.txt";


            if (saveFile.ShowDialog() == true)
            {
                //Work
                var textRange = new TextRange(MyRichTextBox.Document.ContentStart, MyRichTextBox.Document.ContentEnd);
                //DocumentModel document = new DocumentModel();
                //document.Content.LoadText(textRange.);
                //document.Save(saveFile.FileName);

                using (FileStream fs = File.Create(saveFile.FileName))
                {
                    textRange.Save(fs, DataFormats.Rtf);
                }
                */
        }

    }
}

