using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

namespace SqlServerExcelExporter.Core
{
    public sealed class ExportService
    {
        private readonly SqlServerRepository _repository;
        private readonly XlsxWriter _xlsxWriter;

        public ExportService(SqlServerRepository repository, XlsxWriter xlsxWriter)
        {
            if (repository == null)
            {
                throw new ArgumentNullException("repository");
            }

            if (xlsxWriter == null)
            {
                throw new ArgumentNullException("xlsxWriter");
            }

            _repository = repository;
            _xlsxWriter = xlsxWriter;
        }

        public IList<ExportResult> Export(
            TableName tableName,
            string dateColumn,
            DateColumnType dateColumnType,
            DateTime startDate,
            DateTime endDate,
            ExportGroupMode groupMode,
            string outputDirectory,
            IProgress<ExportProgress> progress)
        {
            if (string.IsNullOrWhiteSpace(dateColumn))
            {
                throw new ArgumentException("请选择日期列。", "dateColumn");
            }

            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new ArgumentException("请选择输出目录。", "outputDirectory");
            }

            Directory.CreateDirectory(outputDirectory);

            var periods = PeriodBuilder.Build(startDate, endDate, groupMode);
            var results = new List<ExportResult>();

            for (var index = 0; index < periods.Count; index++)
            {
                var period = periods[index];
                Report(progress, "正在导出 " + period.Label + "...", index + 1, periods.Count, period);

                SqlConnection connection;
                using (var reader = _repository.ExecuteExportReader(
                    tableName,
                    dateColumn,
                    dateColumnType,
                    period.StartInclusive,
                    period.EndExclusive,
                    out connection))
                {
                    var filePath = Path.Combine(outputDirectory, FileNameHelper.BuildExportFileName(tableName, period.Label));
                    var rowCount = _xlsxWriter.Write(filePath, reader);

                    results.Add(new ExportResult
                    {
                        FilePath = filePath,
                        RowCount = rowCount,
                        Period = period
                    });
                }
            }

            Report(progress, "导出完成。", periods.Count, periods.Count, null);
            return results;
        }

        private static void Report(IProgress<ExportProgress> progress, string message, int current, int total, ExportPeriod period)
        {
            if (progress == null)
            {
                return;
            }

            progress.Report(new ExportProgress
            {
                Message = message,
                CurrentPeriod = current,
                TotalPeriods = total,
                PeriodStart = period == null ? (DateTime?)null : period.StartInclusive,
                PeriodEndExclusive = period == null ? (DateTime?)null : period.EndExclusive
            });
        }
    }
}
