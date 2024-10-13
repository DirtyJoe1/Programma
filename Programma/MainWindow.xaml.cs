using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeOpenXml;
using Programma.Resources.Classes;

namespace Programma
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    
    public partial class MainWindow : Window
    {
        public string filePath = "Resources/data.xlsx";
        private bool _hasUnsavedChanges = false;
        public MainWindow()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            InitializeComponent();
            UpdateTable();
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);
            var result = MessageBox.Show("Сохранять ли изменения в файл?", "Внимание!", MessageBoxButton.YesNoCancel);
            switch (result)
            {
                case MessageBoxResult.Yes:
                    var dataTable = ExcelHelper.GetDataTableFromDataGrid<DataModel>(Table.ItemsSource);
                    ExcelHelper.SaveDataGrid(dataTable, filePath);
                    break;
                case MessageBoxResult.No:
                    break;
                case MessageBoxResult.Cancel:
                    e.Cancel = true;
                    break;
            }
        }
        public void UpdateTable()
        {
            Table.ItemsSource = ExcelHelper.ReadExcelFile(filePath);
        }
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dataTable = ExcelHelper.GetDataTableFromDataGrid<DataModel>(Table.ItemsSource);
                ExcelHelper.SaveDataGrid(dataTable, filePath);
                MessageBox.Show("Успешно сохранено");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Возникла ошибка " + ex.Message);
            }    
        }

        private void AddNewEntryButton_Click(object sender, RoutedEventArgs e)
        {
            new AddNewEntryWindow(this).Show();
        }
    }
}