using System.Text.Encodings.Web;
using System.Text.Json;

namespace ExcelReader
{
    public class Program
    {
        private static readonly JsonSerializerOptions options = new()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = true
        };

        public static void Main(string[] args)
        {
            string filePath = "C:\\Users\\Salvatore.Amaddio\\OneDrive - Mahindra Racing UK Limited\\Desktop\\S11_R03JED_MFE_Phase_Document_xxx_001.xlsm";

            try
            {
                List<DataModel> data = ExcelManager.Read(filePath);

                //var sheetGroup = data.GroupBy(s=>s.SheetOrigin).ToList();

                //Console.WriteLine();

                //foreach (var group in sheetGroup)
                //{
                //    Console.WriteLine($"From Sheet: {group.Key}");
                //    Console.WriteLine($"Count: {group.Count()}");

                //    foreach (DataModel? item in group)
                //    {
                //        Console.WriteLine($"\t {item}");
                //    }

                //    Console.WriteLine();
                //}

                string json = JsonSerializer.Serialize(data, options);
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = "atlasData.json";
                string jsonFilePath = Path.Combine(desktopPath, fileName);
                File.WriteAllText(jsonFilePath, json);
                Console.WriteLine($"JSON file named '{fileName}' created on your desktop");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }
    }
}