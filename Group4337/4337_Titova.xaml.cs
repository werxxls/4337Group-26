using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Win32;
using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows;

namespace Group4337
{
    /// <summary>
    /// Логика взаимодействия для _4337_Titova.xaml
    /// </summary>
    public partial class _4337_Titova : Window
    {
                string connStr = "Server=MariA\\MSSQLSERVER03;Database=Titova1var;Trusted_Connection=True;TrustServerCertificate=True;";

        public _4337_Titova()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "JSON|*.json", Title = "Выберите файл" };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var list = ReadJson(dlg.FileName);
                ClearDB();
                SaveToDB(list);
                MessageBox.Show($"Импортировано: {list.Count}");
            }
            catch (Exception ex) { MessageBox.Show($"{ex.Message}"); }
        }

        private List<Service> ReadJson(string path)
        {
            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            return JsonSerializer.Deserialize<List<Service>>(json, options)
            .Select(json => new Service
            {
                ID = json.ID,
                Name = json.Name,
                Type = json.Type,
                Code = json.Code,
                Cost = json.Cost
            }).ToList();
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            InitializeComponent();
            try
            {
                var list = LoadFromDB();
                if (list.Count == 0) { MessageBox.Show("Нет данных для экспорта"); return; }
                SaveWord(list);
                MessageBox.Show("Экспортировано");
            }
            catch (Exception ex) { MessageBox.Show($" {ex.Message}"); }
        }

        private void SaveWord(List<Service> list)
        {
            var dlg = new SaveFileDialog { Filter = "Word|*.docx", Title = "Сохранить", FileName = "Export" };
            if (dlg.ShowDialog() != true) return;

            using (var doc = WordprocessingDocument.Create(dlg.FileName, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                var groups = list.GroupBy(s => s.Type).OrderBy(g => g.Key);
                bool firstPage = true;

                foreach (var group in groups)
                {
                    if (!firstPage) body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    firstPage = false;
                    var titleParagraph = new Paragraph();
                    var titleRun = new Run(new Text(group.Key ?? "Без категории"));
                    var titleProperties = new RunProperties(
                        new Bold(),
                        new FontSize { Val = "28"},
                        new RunFonts
                        {
                            Ascii = "Times New Roman",
                            HighAnsi = "Times New Roman",
                            EastAsia = "Times New Roman",
                            ComplexScript = "Times New Roman"
                    });
                    titleRun.RunProperties = titleProperties;
                    titleParagraph.AppendChild(titleRun);
                    body.AppendChild(titleParagraph);

                    var table = new Table();
                    table.AppendChild(new TableProperties(new TableWidth() { Width = "100%", Type = TableWidthUnitValues.Pct }));

                    var header = new TableRow();
                    header.AppendChild(CreateCell("ID", true));
                    header.AppendChild(CreateCell("Наименование услуги", true));
                    header.AppendChild(CreateCell("Стоимость, руб. за час", true));
                    table.AppendChild(header);

                    foreach (var s in group.OrderBy(x => x.Cost))
                    {
                        var row = new TableRow();
                        row.AppendChild(CreateCell(s.ID.ToString()));
                        row.AppendChild(CreateCell(s.Name ?? ""));
                        row.AppendChild(CreateCell(s.Cost.ToString("0.00")));
                        table.AppendChild(row);
                    }

                    body.AppendChild(table);
                }

                mainPart.Document.Save();
            }
        }

        private TableCell CreateCell(string text, bool header = false)
        {
            var cell = new TableCell();
            var paragraph = new Paragraph();
            var run = new Run(new Text(text));
            var runProperties = new RunProperties(new RunFonts
            {
                Ascii = "Times New Roman",
                HighAnsi = "Times New Roman",
                EastAsia = "Times New Roman",
                ComplexScript = "Times New Roman"
            },
            new FontSize { Val = header ? "28" : "24" });
            if (header) runProperties.AppendChild(new Bold());
            run.RunProperties = runProperties;
            paragraph.AppendChild(run);
            cell.AppendChild(paragraph);
            return cell;
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

        public class Service
        {
            public int ID { get; set; }
            public string? Name { get; set; }
            public string? Type { get; set; }
            public string? Code { get; set; }
            public decimal Cost { get; set; }
        }
    }
}