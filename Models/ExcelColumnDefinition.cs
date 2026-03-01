namespace SmartSAP.Models
{
    public class ExcelColumnDefinition
    {
        public string Header { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string SampleData { get; set; } = string.Empty;
        public int FixedWidth { get; set; }
        public bool ForceUpperCase { get; set; } = true;
        public string[]? AllowedValues { get; set; }

        public ExcelColumnDefinition(string header, string comment, string sampleData, int fixedWidth = 0, bool forceUpperCase = true, string[]? allowedValues = null)
        {
            Header = header;
            Comment = comment;
            SampleData = sampleData;
            FixedWidth = fixedWidth;
            ForceUpperCase = forceUpperCase;
            AllowedValues = allowedValues;
        }
    }
}
