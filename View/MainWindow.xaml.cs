using Microsoft.Win32;
using Model;
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
        private RunFile run1 = new RunFile();
        Datenbank db = null;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;

        public MainWindow()
        {
            db = new Datenbank();
            db.openDatabase();
            listDokumente.Items.Add(db.openDatabase());
            InitializeComponent();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void buttonHinzufuegen_Click(object sender, RoutedEventArgs e)
        {
            //TODO: FILEFILTER Bearbeiten, Documente aus Filter in DB übertragen -> Long DB!
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Geeignete Datei auswählen";
            openFileDialog1.Filter = "Word Documents (*.DOC;*.DOCX)|*.DOC;*.DOCX|";
            openFileDialog1.Filter = "Excel Workbook (*.XLS;*.XSLX)|*.XLS;*.XLSX|";
            openFileDialog1.Filter = "PowerPoint Präsentation (*.PPT;*.PPTX)|*.PPT;*.PPTX";
            openFileDialog1.ShowDialog();
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void buttonWord_Click(object sender, RoutedEventArgs e)
        {
            run1.chooseData("word");
        }

        private void buttonPower_Click(object sender, RoutedEventArgs e)
        {
            run1.chooseData("powerpoint");
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            run1.chooseData("excel");
        }
    }
}
