using System;
using System.Text.RegularExpressions;

namespace SqlServerExcelExporter.Core
{
    public sealed class TableName
    {
        private static readonly Regex IdentifierRegex = new Regex(@"^[A-Za-z_][A-Za-z0-9_@$#]*$", RegexOptions.Compiled);

        private TableName(string schema, string name)
        {
            Schema = schema;
            Name = name;
        }

        public string Schema { get; private set; }

        public string Name { get; private set; }

        public string DisplayName
        {
            get { return Schema + "." + Name; }
        }

        public string SqlName
        {
            get { return EscapeIdentifier(Schema) + "." + EscapeIdentifier(Name); }
        }

        public static TableName Parse(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                throw new ArgumentException("请输入表名。", "input");
            }

            var trimmed = input.Trim();
            var parts = trimmed.Split('.');

            if (parts.Length == 1)
            {
                return Create("dbo", parts[0]);
            }

            if (parts.Length == 2)
            {
                return Create(parts[0], parts[1]);
            }

            throw new ArgumentException("表名格式必须是 TableName 或 schema.TableName。", "input");
        }

        public override string ToString()
        {
            return DisplayName;
        }

        private static TableName Create(string schema, string name)
        {
            schema = (schema ?? string.Empty).Trim();
            name = (name ?? string.Empty).Trim();

            ValidateIdentifier(schema, "schema");
            ValidateIdentifier(name, "table");

            return new TableName(schema, name);
        }

        private static void ValidateIdentifier(string value, string partName)
        {
            if (!IdentifierRegex.IsMatch(value))
            {
                throw new ArgumentException(partName + " 名称只能包含字母、数字、下划线、@、$、#，且不能以数字开头。");
            }
        }

        private static string EscapeIdentifier(string identifier)
        {
            return "[" + identifier.Replace("]", "]]") + "]";
        }
    }
}
