using Programma.Resources.Classes;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace Programma
{
    /// <summary>
    /// Логика взаимодействия для AddNewEntryWindow.xaml
    /// </summary>
    public partial class AddNewEntryWindow : Window
    {
        public string filePath = "Resources/data.xlsx";
        public MainWindow mainWindow;
        public AddNewEntryWindow(MainWindow main)
        {
            InitializeComponent();
            mainWindow = main;
        }

        private void CreateEntryButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelHelper.AddData(filePath, FioField.Text, OtdelField.Text, DateTime.Parse(SetupField.Text), DateTime.Parse(StartField.Text), DateTime.Parse(EndField.Text), Enum.Parse<Statuses>(StatusComboBox.Text));
            mainWindow.UpdateTable();
        }
    }
}
