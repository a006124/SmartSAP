namespace SmartSAP.Models
{
    public class ExcelColumnDefinition
    {
        public string Header { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public string SampleData { get; set; } = string.Empty;

        public ExcelColumnDefinition(string header, string comment, string sampleData)
        {
            Header = header;
            Comment = comment;
            SampleData = sampleData;
        }
    }
}
