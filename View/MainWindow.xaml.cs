using Model;
using System.Windows;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Globalization;

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
        private string filename = null;

        public MainWindow()
        {
            InitializeComponent();
            db = new Datenbank();
            listDokumente.ItemsSource = db.openDatabase();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            //TODO: FILEFILTER Bearbeiten, Documente aus Filter in DB übertragen -> Long DB!
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Geeignete Datei auswählen";
            openFileDialog1.Filter = "Word Documents (*.DOC;*.DOCX)|*.DOC;*.DOCX";
            openFileDialog1.Filter = "Excel Workbook (*.XLS;*.XSLX)|*.XLS;*.XLSX";
            openFileDialog1.Filter = "PowerPoint Präsentation (*.PPT;*.PPTX)|*.PPT;*.PPTX";
            openFileDialog1.ShowDialog();
            filename = openFileDialog1.FileName;
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void buttonWord_Click(object sender, RoutedEventArgs e)
        {
           listDokumente.ItemsSource = db.WordAnzeigen();
        }

        private void buttonPower_Click(object sender, RoutedEventArgs e)
        {
            listDokumente.ItemsSource = db.PowerpointAnzeigen();
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            listDokumente.ItemsSource = db.ExcelAnzeigen();
        }
    }
}
