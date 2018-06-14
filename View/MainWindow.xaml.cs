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
        private RunFile run1 = new RunFile();
        public MainWindow()
        {
            InitializeComponent();
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
