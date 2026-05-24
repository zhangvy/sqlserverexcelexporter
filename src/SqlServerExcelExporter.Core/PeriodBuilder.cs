using System;
using System.Collections.Generic;

namespace SqlServerExcelExporter.Core
{
    public static class PeriodBuilder
    {
        public static IList<ExportPeriod> Build(DateTime startDate, DateTime endDate, ExportGroupMode mode)
        {
            var startInclusive = startDate.Date;
            var endExclusive = endDate.Date.AddDays(1);

            if (endExclusive <= startInclusive)
            {
                throw new ArgumentException("结束日期必须大于或等于起始日期。");
            }

            switch (mode)
            {
                case ExportGroupMode.None:
                    return new List<ExportPeriod>
                    {
                        new ExportPeriod(FormatSingleLabel(startInclusive, endExclusive), startInclusive, endExclusive)
                    };
                case ExportGroupMode.Week:
                    return BuildWeekly(startInclusive, endExclusive);
                case ExportGroupMode.Month:
                    return BuildMonthly(startInclusive, endExclusive);
                default:
                    throw new ArgumentOutOfRangeException("mode");
            }
        }

        private static IList<ExportPeriod> BuildMonthly(DateTime startInclusive, DateTime endExclusive)
        {
            var result = new List<ExportPeriod>();
            var cursor = new DateTime(startInclusive.Year, startInclusive.Month, 1);

            while (cursor < endExclusive)
            {
                var naturalStart = cursor;
                var naturalEnd = cursor.AddMonths(1);
                var periodStart = Max(startInclusive, naturalStart);
                var periodEnd = Min(endExclusive, naturalEnd);

                if (periodStart < periodEnd)
                {
                    result.Add(new ExportPeriod(naturalStart.ToString("yyyy-MM"), periodStart, periodEnd));
                }

                cursor = naturalEnd;
            }

            return result;
        }

        private static IList<ExportPeriod> BuildWeekly(DateTime startInclusive, DateTime endExclusive)
        {
            var result = new List<ExportPeriod>();
            var cursor = StartOfWeekMonday(startInclusive);

            while (cursor < endExclusive)
            {
                var naturalStart = cursor;
                var naturalEnd = cursor.AddDays(7);
                var periodStart = Max(startInclusive, naturalStart);
                var periodEnd = Min(endExclusive, naturalEnd);

                if (periodStart < periodEnd)
                {
                    result.Add(new ExportPeriod(FormatWeekLabel(naturalStart), periodStart, periodEnd));
                }

                cursor = naturalEnd;
            }

            return result;
        }

        private static DateTime StartOfWeekMonday(DateTime date)
        {
            var diff = ((int)date.DayOfWeek + 6) % 7;
            return date.Date.AddDays(-diff);
        }

        private static string FormatSingleLabel(DateTime startInclusive, DateTime endExclusive)
        {
            return startInclusive.ToString("yyyyMMdd") + "-" + endExclusive.AddDays(-1).ToString("yyyyMMdd");
        }

        private static string FormatWeekLabel(DateTime naturalWeekStart)
        {
            var thursday = naturalWeekStart.AddDays(3);
            var firstThursday = new DateTime(thursday.Year, 1, 4);
            firstThursday = StartOfWeekMonday(firstThursday).AddDays(3);
            var week = 1 + (int)((thursday - firstThursday).TotalDays / 7);

            if (week < 1)
            {
                return FormatWeekLabel(StartOfWeekMonday(new DateTime(thursday.Year - 1, 12, 31)));
            }

            var weeksInYear = WeeksInIsoYear(thursday.Year);
            if (week > weeksInYear)
            {
                return (thursday.Year + 1).ToString("0000") + "-W01";
            }

            return thursday.Year.ToString("0000") + "-W" + week.ToString("00");
        }

        private static int WeeksInIsoYear(int year)
        {
            var dec28 = new DateTime(year, 12, 28);
            var firstThursday = StartOfWeekMonday(new DateTime(year, 1, 4)).AddDays(3);
            var lastThursday = StartOfWeekMonday(dec28).AddDays(3);
            return 1 + (int)((lastThursday - firstThursday).TotalDays / 7);
        }

        private static DateTime Min(DateTime left, DateTime right)
        {
            return left <= right ? left : right;
        }

        private static DateTime Max(DateTime left, DateTime right)
        {
            return left >= right ? left : right;
        }
    }
}
