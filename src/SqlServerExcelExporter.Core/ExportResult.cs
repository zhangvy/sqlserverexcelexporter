namespace SqlServerExcelExporter.Core
{
    public sealed class ExportResult
    {
        public string FilePath { get; set; }

        public int RowCount { get; set; }

        public ExportPeriod Period { get; set; }
    }
}
