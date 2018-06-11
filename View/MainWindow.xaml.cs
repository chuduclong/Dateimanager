using Microsoft.Win32;
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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace View
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            string path;
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                path = file.FileName;
            }
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            //TODO: Passende Lösung für das Problem finden -> Word per Befehl öffnen
            
            //Code soweit unbrauchbar.
            // Load DOCX or DOC document.
            var document = DocumentModel.Load(isDocx ? "Document.docx" : "Document.doc");

            // Iterate over all paragraphs in the document.
            foreach (Paragraph para in document.GetChildElements(true, ElementType.Paragraph))
            {
                // Iterate over all runs in the paragraph and write their text to Console.
                foreach (Run run in para.GetChildElements(true, ElementType.Run))
                    Console.Write(run.Text);
                Console.WriteLine();
            }
        }
    }
}
