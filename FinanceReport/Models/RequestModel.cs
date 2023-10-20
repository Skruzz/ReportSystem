namespace FinanceReport.Models
{
    public class RequestModel
    {
        public string WorksheetName { get; set; }
        public List<FieldMapping> FieldMappings { get; set; }
    }
}
