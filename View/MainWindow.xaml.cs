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
        private string dateiart = null;

        public MainWindow()
        {
            InitializeComponent();
            db = new Datenbank();
            listDokumente.ItemsSource = db.openDatabase();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Geeignete Datei auswählen";
            openFileDialog1.Filter = "Word Documents (*.DOC;*.DOCX)|*.DOC;*.DOCX|Excel Workbook (*.XLS;*.XSLX)|*.XLS;*.XLSX|PowerPoint Präsentation (*.PPT;*.PPTX)|*.PPT;*.PPTX";
            openFileDialog1.ShowDialog();
            filename = openFileDialog1.FileName;
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
           if(dateiart != null)
            {
                if(filename != null)
                {
                    switch (dateiart)
                    {
                        case "word":
                            run1.openExistingWordFile(filename);
                            break;

                        case "powerpoint":
                            run1.openExistingPowerPointPresentation(filename);
                            break;

                        case "excel":
                            run1.openExistingPowerPointPresentation(filename);
                            break;
                    }
                }
                
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
           
        }

        private void buttonWord_Click(object sender, RoutedEventArgs e)
        {
            listDokumente.ItemsSource = db.WordAnzeigen();
            dateiart = "word";
        }

        private void buttonPower_Click(object sender, RoutedEventArgs e)
        {
            listDokumente.ItemsSource = db.PowerpointAnzeigen();
            dateiart = "powerpoint";
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            listDokumente.ItemsSource = db.ExcelAnzeigen();
            dateiart = "excel";
        }
    }
}
