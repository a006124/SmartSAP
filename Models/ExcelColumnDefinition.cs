namespace SmartSAP.Models
{
    public class ExcelColumnDefinition
    {
        public string Header { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string SampleData { get; set; } = string.Empty;
        public int FixedWidth { get; set; }

        public ExcelColumnDefinition(string header, string comment, string sampleData, int fixedWidth = 0)
        {
            Header = header;
            Comment = comment;
            SampleData = sampleData;
            FixedWidth = fixedWidth;
        }
    }
}
