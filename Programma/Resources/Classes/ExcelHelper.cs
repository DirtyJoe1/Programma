using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows;

namespace Programma.Resources.Classes
{
    public static class ExcelHelper
    {

        public static DataTable GetDataTableFromDataGrid<T>(IEnumerable list)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            object[] values = new object[props.Count];
            foreach (T item in list)
            {
                for (int i = 0; i < values.Length; i++)
                    values[i] = props[i].GetValue(item) ?? DBNull.Value;
                table.Rows.Add(values);
            }
            return table;
        }
        public static List<DataModel> ReadExcelFile(string filePath)
        {
            List<DataModel> data = new();
            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets.First();
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    try
                    {
                        var rowData = new DataModel
                        {
                            Id = int.Parse(worksheet.Cells[row, 1].Value.ToString()),
                            FIO = worksheet.Cells[row, 2].Value.ToString(),
                            Otdel = worksheet.Cells[row, 3].Value.ToString(),
                            Setup = DateTime.Parse(worksheet.Cells[row, 4].Value.ToString()),
                            Start = DateTime.Parse(worksheet.Cells[row, 5].Value.ToString()),
                            End = DateTime.Parse(worksheet.Cells[row, 6].Value.ToString()),
                            Status = Enum.Parse<Statuses>(worksheet.Cells[row, 7].Value.ToString()),
                        };
                        data.Add(rowData);
                    }
                    catch (Exception) { }
                }
            }
            return data;
        }
        public static void SaveDataGrid(DataTable data, string filePath)
        {
            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = 0;
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        worksheet.Cells[rowCount + 1, j + 1].Value = data.Rows[i][j];
                    }
                    rowCount++;
                }
                package.Save();
            }
        }
        public static void AddData(string filePath, string fio, string otdel, DateTime setup, DateTime start, DateTime end, Statuses status)
        {
            if (string.IsNullOrEmpty(fio) || string.IsNullOrEmpty(otdel) /*|| (setup == null) || (start == null) || (end == null) || (status == null)*/)
            {
                MessageBox.Show("Некоторые поля пусты!");
                return;
            }
            if (!System.IO.File.Exists(filePath))
            {
                MessageBox.Show("Файл не найден!");
                return;
            }
            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var lastrow = worksheet.Dimension.Rows + 1;
                worksheet.Cells[lastrow, 1].Value = lastrow;
                worksheet.Cells[lastrow, 2].Value = fio;
                worksheet.Cells[lastrow, 3].Value = otdel;
                worksheet.Cells[lastrow, 4].Value = setup.ToString("dd/MM/yyyy");
                worksheet.Cells[lastrow, 5].Value = start.ToString("dd/MM/yyyy");
                worksheet.Cells[lastrow, 6].Value = end.ToString("dd/MM/yyyy");
                worksheet.Cells[lastrow, 7].Value = status;
                package.Save();
                MessageBox.Show("Запись создана");
            }
        }
    }
}
