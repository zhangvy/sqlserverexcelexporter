using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace SqlServerExcelExporter.Core
{
    public sealed class SqlServerRepository
    {
        private readonly string _connectionString;

        public SqlServerRepository(string connectionString)
        {
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException("请在 App.config 的 connectionStrings 中配置 DefaultConnection。", "connectionString");
            }

            _connectionString = connectionString;
        }

        public IList<ColumnInfo> GetColumns(TableName tableName)
        {
            var columns = new List<ColumnInfo>();

            using (var connection = new SqlConnection(_connectionString))
            using (var command = connection.CreateCommand())
            {
                command.CommandText = @"
SELECT
    c.COLUMN_NAME,
    c.DATA_TYPE,
    c.IS_NULLABLE,
    c.CHARACTER_MAXIMUM_LENGTH,
    c.NUMERIC_PRECISION,
    c.NUMERIC_SCALE,
    c.ORDINAL_POSITION
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_SCHEMA = @SchemaName
  AND c.TABLE_NAME = @TableName
ORDER BY c.ORDINAL_POSITION;";
                command.Parameters.Add("@SchemaName", SqlDbType.NVarChar, 128).Value = tableName.Schema;
                command.Parameters.Add("@TableName", SqlDbType.NVarChar, 128).Value = tableName.Name;

                connection.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(new ColumnInfo
                        {
                            Name = reader.GetString(0),
                            DataType = reader.GetString(1),
                            IsNullable = string.Equals(reader.GetString(2), "YES", StringComparison.OrdinalIgnoreCase),
                            MaxLength = reader.IsDBNull(3) ? (int?)null : Convert.ToInt32(reader.GetValue(3)),
                            NumericPrecision = reader.IsDBNull(4) ? (int?)null : Convert.ToInt32(reader.GetValue(4)),
                            NumericScale = reader.IsDBNull(5) ? (int?)null : Convert.ToInt32(reader.GetValue(5)),
                            Ordinal = Convert.ToInt32(reader.GetValue(6))
                        });
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new InvalidOperationException("找不到表 " + tableName.DisplayName + "，请确认表名和 schema。");
            }

            return columns;
        }

        public SqlDataReader ExecuteExportReader(
            TableName tableName,
            string dateColumn,
            DateColumnType dateColumnType,
            DateTime startInclusive,
            DateTime endExclusive,
            out SqlConnection connection)
        {
            connection = new SqlConnection(_connectionString);
            var command = connection.CreateCommand();
            command.CommandTimeout = 0;
            command.CommandText =
                "SELECT * FROM " + tableName.SqlName +
                " WHERE " + EscapeIdentifier(dateColumn) + " >= @StartDate" +
                " AND " + EscapeIdentifier(dateColumn) + " < @EndDate" +
                " ORDER BY " + EscapeIdentifier(dateColumn) + ";";

            AddDateParameter(command, "@StartDate", dateColumnType, startInclusive);
            AddDateParameter(command, "@EndDate", dateColumnType, endExclusive);

            try
            {
                connection.Open();
                return command.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch
            {
                command.Dispose();
                connection.Dispose();
                throw;
            }
        }

        private static void AddDateParameter(SqlCommand command, string name, DateColumnType dateColumnType, DateTime value)
        {
            var dbType = dateColumnType == DateColumnType.Date ? SqlDbType.Date : SqlDbType.DateTime2;
            command.Parameters.Add(name, dbType).Value = value;
        }

        private static string EscapeIdentifier(string identifier)
        {
            return "[" + identifier.Replace("]", "]]") + "]";
        }
    }
}
