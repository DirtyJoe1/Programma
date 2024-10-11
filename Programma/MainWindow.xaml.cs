using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OfficeOpenXml;
using Programma.Resources;

namespace Programma
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Table.ItemsSource = ReadExcelFile("Resources/data.xlsx");
        }
        public List<DataModel> ReadExcelFile(string filePath)
        {
            List<DataModel> data = new List<DataModel>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets.First();
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    try
                    {
                        var rowData = new DataModel
                        {
                            Id = Int16.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            FIO = worksheet.Cells[row, 2].Value.ToString(),
                            Otdel = worksheet.Cells[row, 3].Value.ToString(),
                            Setup = DateTime.Parse(worksheet.Cells[row, 4].Value.ToString()),
                            Start = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString()),
                            End = DateTime.Parse(worksheet.Cells[row, 6].Value.ToString()),
                            Status = worksheet.Cells[row, 7].Value.ToString(),
                        };
                        data.Add(rowData);
                    }
                    catch (Exception) {}
                }
            }
            return data;
        }
    }
}