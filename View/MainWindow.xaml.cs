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
        Boolean word = false;
        Boolean powerpoint = false;
        Boolean excel = false;
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

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            
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
