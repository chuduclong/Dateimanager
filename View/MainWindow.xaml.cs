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
        public MainWindow()
        {
            InitializeComponent();
            db = new Datenbank();
            listDokumente.ItemsSource = db.openDatabase();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            //run1.openExistingWordFile();
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            run1.openNewWordFile();
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            run1.closeCurrentWordFile();
        }
    }
}
