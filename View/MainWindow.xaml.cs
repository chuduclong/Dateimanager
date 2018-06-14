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
using Word = Microsoft.Office.Interop.Word;

namespace View
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Word.Application WordApp;
        Word.Document Doc;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonChoose_Click(object sender, RoutedEventArgs e)
        {
            Object oMissing = System.Reflection.Missing.Value;
            Doc = WordApp.Documents.Open("DATNAME", oMissing, false);
        }

        private void buttonOpen_Click(object sender, RoutedEventArgs e)
        {
            WordApp = new Word.Application();
            WordApp.Visible = true;
            Doc = WordApp.Documents.Add();
        }

        private void buttonDelete_Click(object sender, RoutedEventArgs e)
        {
            Doc.Close();
        }
    }
}
