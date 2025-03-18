using ExcelDataReader;
using System.Text;

namespace ExcelReader
{
    public class ExcelManager
    {
        private const string ATLAS_NAME = "atlas name";
        private const string MIN = "min";
        private const string MAX = "max";

        public static List<DataModel> Read(string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            List<DataModel> dataModels = new();

            using FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            do
            {
                string sheetName = reader.Name;
                int atlasColumn = 0;
                int minColumn = 0;
                int maxColumn = 0;

//                Console.WriteLine($"Reading sheet: {sheetName}");

                while (reader.Read()) //read sheet's rows
                {
                    DataModel dataModel = new()
                    {
                        SheetOrigin = sheetName
                    };

                    int col = 0;

                    for (int i = 0; i < reader.FieldCount; i++) //for each cell in the current row
                    {
                        col++;

                        string? cellValue = reader.GetValue(i)?.ToString()?.Trim();

                        if (cellValue == null)
                            continue;

                        if (col == atlasColumn)
                            dataModel.AtlasName = cellValue;
                        else if (col == minColumn)
                            dataModel.Min = cellValue;
                        else if (col == maxColumn)
                            dataModel.Max = cellValue;

                        //identify in which column the values are
                        if (cellValue.ToLower().Equals(ATLAS_NAME))
                            atlasColumn = col;
                        else if (cellValue.ToLower().Equals(MIN) && minColumn == 0 && atlasColumn > 0)
                            minColumn = col;
                        else if (cellValue.ToLower().Equals(MAX) && maxColumn == 0 && atlasColumn > 0)
                            maxColumn = col;
                    }

                    if (!string.IsNullOrEmpty(dataModel.AtlasName))
                        dataModels.Add(dataModel);

                }//GO TO NEXT ROW
            }
            while (reader.NextResult()); // GO TO NEXT SHEET

            return dataModels;
        }
    }
}