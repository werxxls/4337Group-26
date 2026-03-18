using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;

namespace Group4337
{
    public partial class _4337_Titova : Window
    {
        string connStr = "Server=MariA\\MSSQLSERVER03;Database=Titova1var;Trusted_Connection=True;TrustServerCertificate=True;";

        public _4337_Titova() => InitializeComponent();

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "Excel|*.xlsx", Title = "Выберите файл" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var list = ReadExcel(dlg.FileName);
                ClearDB();
                SaveToDB(list);
                MessageBox.Show($"Импортировано: {list.Count}");
            }
            catch (Exception ex) { MessageBox.Show($"{ex.Message}"); }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var list = LoadFromDB();
                if (list.Count == 0) { MessageBox.Show("Нет данных для экспорта"); return; }
                SaveExcel(list);
                MessageBox.Show("Экспортировано");
            }
            catch (Exception ex) { MessageBox.Show($" {ex.Message}"); }
        }

        private List<Service> ReadExcel(string path)
        {
            var list = new List<Service>();
            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet(1);
                int row = 2;
                while (!ws.Cell(row, 1).IsEmpty())
                {
                    list.Add(new Service
                    {
                        ID = int.Parse(ws.Cell(row, 1).GetValue<string>().Trim()),
                        Name = ws.Cell(row, 2).GetValue<string>(),
                        Type = ws.Cell(row, 3).GetValue<string>(),
                        Code = ws.Cell(row, 4).GetValue<string>(),
                        Cost = decimal.Parse(ws.Cell(row, 5).GetValue<string>().Trim())
                    });
                    row++;
                }
            }
            return list;
        }

        private void ClearDB()
        {
            using (var c = new SqlConnection(connStr))
            {
                c.Open();
                new SqlCommand("DELETE FROM Services", c).ExecuteNonQuery();
            }
        }

        private void SaveToDB(List<Service> list)
        {
            using (var c = new SqlConnection(connStr))
            {
                c.Open();
                foreach (var s in list)
                {
                    var cmd = new SqlCommand("INSERT INTO Services VALUES(@id,@n,@t,@c,@p)", c);
                    cmd.Parameters.AddWithValue("@id", s.ID);
                    cmd.Parameters.AddWithValue("@n", s.Name ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@t", s.Type ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@c", s.Code ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@p", s.Cost);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private List<Service> LoadFromDB()
        {
            var list = new List<Service>();
            using (var c = new SqlConnection(connStr))
            {
                c.Open();
                var r = new SqlCommand("SELECT * FROM Services", c).ExecuteReader();
                while (r.Read())
                    list.Add(new Service
                    {
                        ID = (int)r["ID"],
                        Name = r["ServiceName"] as string,
                        Type = r["ServiceType"] as string,
                        Code = r["ServiceCode"] as string,
                        Cost = (decimal)r["CostPerHour"]
                    });
            }
            return list;
        }

        private void SaveExcel(List<Service> list)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                Title = "Сохранить данные",
                FileName = "Export"
            };

            if (dlg.ShowDialog() != true) return;

            using (var wb = new XLWorkbook())
            {
                var groups = list.GroupBy(x => x.Type).OrderBy(g => g.Key);

                foreach (var group in groups)
                {
                    var ws = wb.Worksheets.Add(group.Key);

                    ws.Cell(1, 1).Value = "ID";
                    ws.Cell(1, 2).Value = "Наименование услуги";
                    ws.Cell(1, 3).Value = "Стоимость, руб. за час";

                    int row = 2;
                    foreach (var service in group.OrderBy(s => s.Cost))
                    {
                        ws.Cell(row, 1).Value = service.ID;
                        ws.Cell(row, 2).Value = service.Name;
                        ws.Cell(row, 3).Value = service.Cost;
                        row++;
                    }

                    ws.Columns().AdjustToContents();
                }
                wb.SaveAs(dlg.FileName);
            }
        }
    }

    public class Service
    {
        public int ID { get; set; }
        public string? Name { get; set; }
        public string? Type { get; set; }
        public string? Code { get; set; }
        public decimal Cost { get; set; }
    }
}