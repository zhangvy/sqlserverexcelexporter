using System;

namespace SqlServerExcelExporter.Core
{
    public sealed class ExportProgress
    {
        public string Message { get; set; }

        public int CurrentPeriod { get; set; }

        public int TotalPeriods { get; set; }

        public DateTime? PeriodStart { get; set; }

        public DateTime? PeriodEndExclusive { get; set; }
    }
}
