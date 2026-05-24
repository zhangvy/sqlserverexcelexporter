using System;
using System.Linq;
using SqlServerExcelExporter.Core;

namespace SqlServerExcelExporter.Tests
{
    internal static class Program
    {
        private static int _passed;
        private static int _failed;

        private static int Main()
        {
            Run("TableName defaults schema to dbo", TableNameDefaultsSchema);
            Run("TableName accepts schema.table", TableNameAcceptsSchemaTable);
            Run("TableName rejects invalid identifier", TableNameRejectsInvalidIdentifier);
            Run("None mode uses inclusive UI end date", NoneModeUsesInclusiveEndDate);
            Run("Monthly mode splits natural months and clips edges", MonthlyModeSplitsNaturalMonthsAndClipsEdges);
            Run("Weekly mode starts on Monday and clips edges", WeeklyModeStartsOnMondayAndClipsEdges);
            Run("Date column type recognizes SQL Server date types", DateColumnTypeRecognizesDateTypes);
            Run("File name helper builds xlsx name", FileNameHelperBuildsXlsxName);

            Console.WriteLine();
            Console.WriteLine("Passed: " + _passed);
            Console.WriteLine("Failed: " + _failed);
            return _failed == 0 ? 0 : 1;
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("[PASS] " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("[FAIL] " + name);
                Console.WriteLine("       " + ex.Message);
            }
        }

        private static void TableNameDefaultsSchema()
        {
            var tableName = TableName.Parse("Orders");
            Equal("dbo", tableName.Schema);
            Equal("Orders", tableName.Name);
            Equal("[dbo].[Orders]", tableName.SqlName);
        }

        private static void TableNameAcceptsSchemaTable()
        {
            var tableName = TableName.Parse("sales.Orders");
            Equal("sales", tableName.Schema);
            Equal("Orders", tableName.Name);
            Equal("sales.Orders", tableName.DisplayName);
        }

        private static void TableNameRejectsInvalidIdentifier()
        {
            Throws<ArgumentException>(() => TableName.Parse("dbo.Orders;DROP TABLE Users"));
            Throws<ArgumentException>(() => TableName.Parse("dbo.2026Orders"));
            Throws<ArgumentException>(() => TableName.Parse("a.b.c"));
        }

        private static void NoneModeUsesInclusiveEndDate()
        {
            var periods = PeriodBuilder.Build(new DateTime(2026, 1, 1), new DateTime(2026, 1, 31), ExportGroupMode.None);
            Equal(1, periods.Count);
            Equal(new DateTime(2026, 1, 1), periods[0].StartInclusive);
            Equal(new DateTime(2026, 2, 1), periods[0].EndExclusive);
        }

        private static void MonthlyModeSplitsNaturalMonthsAndClipsEdges()
        {
            var periods = PeriodBuilder.Build(new DateTime(2026, 1, 15), new DateTime(2026, 3, 10), ExportGroupMode.Month);
            Equal(3, periods.Count);

            Equal("2026-01", periods[0].Label);
            Equal(new DateTime(2026, 1, 15), periods[0].StartInclusive);
            Equal(new DateTime(2026, 2, 1), periods[0].EndExclusive);

            Equal("2026-02", periods[1].Label);
            Equal(new DateTime(2026, 2, 1), periods[1].StartInclusive);
            Equal(new DateTime(2026, 3, 1), periods[1].EndExclusive);

            Equal("2026-03", periods[2].Label);
            Equal(new DateTime(2026, 3, 1), periods[2].StartInclusive);
            Equal(new DateTime(2026, 3, 11), periods[2].EndExclusive);
        }

        private static void WeeklyModeStartsOnMondayAndClipsEdges()
        {
            var periods = PeriodBuilder.Build(new DateTime(2026, 1, 7), new DateTime(2026, 1, 20), ExportGroupMode.Week);
            Equal(3, periods.Count);

            Equal("2026-W02", periods[0].Label);
            Equal(new DateTime(2026, 1, 7), periods[0].StartInclusive);
            Equal(new DateTime(2026, 1, 12), periods[0].EndExclusive);

            Equal("2026-W03", periods[1].Label);
            Equal(new DateTime(2026, 1, 12), periods[1].StartInclusive);
            Equal(new DateTime(2026, 1, 19), periods[1].EndExclusive);

            Equal("2026-W04", periods[2].Label);
            Equal(new DateTime(2026, 1, 19), periods[2].StartInclusive);
            Equal(new DateTime(2026, 1, 21), periods[2].EndExclusive);
        }

        private static void DateColumnTypeRecognizesDateTypes()
        {
            Equal(DateColumnType.Date, ColumnInfo.GetDateColumnType("date"));
            Equal(DateColumnType.DateTime, ColumnInfo.GetDateColumnType("datetime"));
            Equal(DateColumnType.DateTime, ColumnInfo.GetDateColumnType("datetime2"));
            Equal(DateColumnType.DateTime, ColumnInfo.GetDateColumnType("smalldatetime"));
            Equal(DateColumnType.DateTime, ColumnInfo.GetDateColumnType("datetimeoffset"));
            Equal(DateColumnType.Unknown, ColumnInfo.GetDateColumnType("varchar"));
        }

        private static void FileNameHelperBuildsXlsxName()
        {
            var fileName = FileNameHelper.BuildExportFileName(TableName.Parse("dbo.Orders"), "2026-01");
            Equal("dbo_Orders_2026-01.xlsx", fileName);
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!object.Equals(expected, actual))
            {
                throw new InvalidOperationException("Expected <" + expected + "> but got <" + actual + ">.");
            }
        }

        private static void Throws<T>(Action action) where T : Exception
        {
            try
            {
                action();
            }
            catch (T)
            {
                return;
            }

            throw new InvalidOperationException("Expected exception " + typeof(T).Name + ".");
        }
    }
}
