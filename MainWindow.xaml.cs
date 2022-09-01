using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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
using static ReadExcelUsingSylvan.ExcelHandler;

namespace ReadExcelUsingSylvan
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.ShowDialog();

            System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            watch.Start();

            //issue is not with Sylvan.Data.Excel. -it's the illegal characters in the data, which still works fine with Sylvan,
            //but columns that have "." in the column header won't be displayed in the DataGrid.
            //The only issue I found in Sylvan.Data.Excel is when you import data where the first row is empty. Which isn't an issue for me.

            //to see header issue, DataTable dt = SylvanReadExcelFile(fileDialog.FileName);
            DataTable dt = SylvanReadExcelFile(fileDialog.FileName);

            watch.Stop();
            System.Diagnostics.Debug.WriteLine(watch.ElapsedMilliseconds + "\tms");

            MainDataGrid.ItemsSource = dt.DefaultView;
        }
    }
}
