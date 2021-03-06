﻿using Model;
using System.Windows;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

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

        private void buttonOeffnen_Click(object sender, RoutedEventArgs e)
        {
            Projekt p = (Projekt)listDokumente.SelectedItem;
            if (p.Dateiart != null)
            {
                if (p.Name != null)
                {
                    filename = p.Name;
                    switch (p.Dateiart)
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


        private void buttonHinzufuegen_Click(object sender, RoutedEventArgs e)
        {
            openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Geeignete Datei auswählen";
            openFileDialog1.Filter = "Word Documents (*.DOC;*.DOCX)|*.DOC;*.DOCX|Excel Workbook (*.XLS;*.XSLX)|*.XLS;*.XLSX|PowerPoint Präsentation (*.PPT;*.PPTX)|*.PPT;*.PPTX";
            openFileDialog1.ShowDialog();
            filename = openFileDialog1.FileName;
            try
            {
                FileInfo fi = new FileInfo(filename);
                String[] z = filename.Split('.');
                db.AddDokus(fi.FullName, z[1], fi.CreationTime);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            db.delete(listDokumente.SelectedItem);
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
