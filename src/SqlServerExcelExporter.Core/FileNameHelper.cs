using System.IO;
using System.Text;

namespace SqlServerExcelExporter.Core
{
    public static class FileNameHelper
    {
        public static string BuildExportFileName(TableName tableName, string label)
        {
            return Clean(tableName.Schema + "_" + tableName.Name + "_" + label) + ".xlsx";
        }

        public static string Clean(string value)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var builder = new StringBuilder(value.Length);

            foreach (var ch in value)
            {
                builder.Append(System.Array.IndexOf(invalid, ch) >= 0 ? '_' : ch);
            }

            return builder.ToString();
        }
    }
}
