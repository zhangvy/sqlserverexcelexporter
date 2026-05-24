using System;

namespace SqlServerExcelExporter.Core
{
    public sealed class ExportPeriod
    {
        public ExportPeriod(string label, DateTime startInclusive, DateTime endExclusive)
        {
            Label = label;
            StartInclusive = startInclusive;
            EndExclusive = endExclusive;
        }

        public string Label { get; private set; }

        public DateTime StartInclusive { get; private set; }

        public DateTime EndExclusive { get; private set; }
    }
}
