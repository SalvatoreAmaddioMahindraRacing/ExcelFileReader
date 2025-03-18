using System.Text.Json.Serialization;

namespace ExcelReader
{
    public class DataModel
    {
        #region backing fields
        private string? _atlasName;
        private string? _min;
        private string? _max;
        #endregion

        [JsonIgnore]
        public string SheetOrigin { get; set; } = string.Empty;

        public string? AtlasName
        {
            get => _atlasName;
            set
            {
                if (string.IsNullOrEmpty(value))
                    _atlasName = null;
                else
                    _atlasName = value;
            }
        }

        public string? Min
        {
            get => _min;
            set
            {
                if (string.IsNullOrEmpty(value) || StringIsNotNumeric(value))
                    _min = null;
                else
                    _min = value;
            }
        }

        public string? Max
        {
            get => _max;
            set
            {
                if (string.IsNullOrEmpty(value) || StringIsNotNumeric(value))
                    _max = null;
                else
                    _max = value;
            }
        }

        private static bool StringIsNotNumeric(string? value)
        {
            return !double.TryParse(value, out double d);
        }

        public override string? ToString()
        {
            return $"Name: {AtlasName} MIN: {Min}; MAX: {Max}";
        }
    }
}