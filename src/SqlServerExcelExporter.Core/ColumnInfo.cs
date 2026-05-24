namespace SqlServerExcelExporter.Core
{
    public sealed class ColumnInfo
    {
        public string Name { get; set; }

        public string DataType { get; set; }

        public bool IsNullable { get; set; }

        public int? MaxLength { get; set; }

        public int? NumericPrecision { get; set; }

        public int? NumericScale { get; set; }

        public int Ordinal { get; set; }

        public bool IsDateColumn
        {
            get { return GetDateColumnType(DataType) != DateColumnType.Unknown; }
        }

        public DateColumnType DateColumnType
        {
            get { return GetDateColumnType(DataType); }
        }

        public static DateColumnType GetDateColumnType(string dataType)
        {
            if (string.IsNullOrWhiteSpace(dataType))
            {
                return DateColumnType.Unknown;
            }

            switch (dataType.Trim().ToLowerInvariant())
            {
                case "date":
                    return DateColumnType.Date;
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                case "datetimeoffset":
                    return DateColumnType.DateTime;
                default:
                    return DateColumnType.Unknown;
            }
        }
    }
}
