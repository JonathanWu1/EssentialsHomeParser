using System;
using Home;
using System.Text.Json;
using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using System.Text;
using ClosedXML.Excel;
using YamlDotNet.Serialization.TypeInspectors;


internal class Program
{
    private static void Main(string[] args)
    {
        string[] files = Directory.GetFiles("./userdata", "*.yml", SearchOption.AllDirectories);
        var deserializer = new DeserializerBuilder().WithNamingConvention(CamelCaseNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .WithTypeInspector(inner => inner, s => s.InsteadOf<YamlAttributesTypeInspector>())
            .WithTypeInspector(
                    inner => new YamlAttributesTypeInspector(inner),
                    s => s.Before<NamingConventionTypeInspector>()
                    )
            .Build();
        List<UserData> users = new();
        foreach (var filepath in files)
        {
            using (var reader = new StreamReader(filepath))
            {

                // Load the stream
                var yaml = new YamlStream();
                yaml.Load(reader);
                var buffer = new StringBuilder();
                using (var writer = new StringWriter(buffer))
                {
                    yaml.Save(writer);
                    var yamlText = writer.ToString();
                    var deserialized = deserializer.Deserialize<UserData>(yamlText);
                    users.Add(deserialized);

                }
            }
        }
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("UserData");
            worksheet.Cell("A1").Value = "Player";
            worksheet.Cell("B1").Value = "WorldName";
            worksheet.Cell("C1").Value = "Home";
            worksheet.Cell("D1").Value = "X";
            worksheet.Cell("E1").Value = "Y";
            worksheet.Cell("F1").Value = "Z";
            int row = 2;
            foreach (var item in users.OrderBy(o => o.LastAccountName))
            {
                foreach (var key in item.Homes.Keys)
                {
                    try
                    {
                        if (key != null)
                        {
                            var home = item.Homes[key];
                            if (home != null)
                            {
                                worksheet.Cell($"A{row}").Value = item.LastAccountName;
                                worksheet.Cell($"B{row}").Value = home.WorldName;
                                worksheet.Cell($"C{row}").Value = key;
                                worksheet.Cell($"D{row}").Value = Convert.ToInt32(home.X);
                                worksheet.Cell($"E{row}").Value = Convert.ToInt32(home.Y);
                                worksheet.Cell($"F{row}").Value = Convert.ToInt32(home.Z);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.StackTrace);
                        // Console.WriteLine($"Skipped: {JsonSerializer.Serialize(item)}");
                    }
                    row++;
                }
            }

            workbook.SaveAs("UserHomes.xlsx");
        }
    }
}
