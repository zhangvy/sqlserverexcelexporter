using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace SqlServerExcelExporter.Core
{
    public sealed class XlsxWriter
    {
        private const int MaxExcelRows = 1048576;

        public int Write(string filePath, IDataReader reader)
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var archive = ZipFile.Open(filePath, ZipArchiveMode.Create))
            {
                WriteTextEntry(archive, "[Content_Types].xml", ContentTypesXml());
                WriteTextEntry(archive, "_rels/.rels", RootRelationshipsXml());
                WriteTextEntry(archive, "xl/workbook.xml", WorkbookXml());
                WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml());
                WriteTextEntry(archive, "xl/styles.xml", StylesXml());

                var sheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
                using (var stream = sheetEntry.Open())
                using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    CloseOutput = false,
                    Indent = false
                }))
                {
                    return WriteWorksheet(writer, reader);
                }
            }
        }

        private static int WriteWorksheet(XmlWriter writer, IDataReader reader)
        {
            writer.WriteStartDocument(true);
            writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteStartElement("sheetData");

            WriteHeaderRow(writer, reader);

            var rowIndex = 2;
            var dataRows = 0;
            while (reader.Read())
            {
                if (rowIndex > MaxExcelRows)
                {
                    throw new InvalidOperationException("单个 Excel 工作表最多支持 1,048,576 行，请缩小日期范围后重试。");
                }

                WriteDataRow(writer, reader, rowIndex);
                rowIndex++;
                dataRows++;
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndDocument();
            return dataRows;
        }

        private static void WriteHeaderRow(XmlWriter writer, IDataReader reader)
        {
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", "1");

            for (var column = 0; column < reader.FieldCount; column++)
            {
                WriteInlineStringCell(writer, 1, column + 1, reader.GetName(column));
            }

            writer.WriteEndElement();
        }

        private static void WriteDataRow(XmlWriter writer, IDataReader reader, int rowIndex)
        {
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", rowIndex.ToString());

            for (var column = 0; column < reader.FieldCount; column++)
            {
                if (reader.IsDBNull(column))
                {
                    continue;
                }

                var value = reader.GetValue(column);
                WriteCell(writer, rowIndex, column + 1, value);
            }

            writer.WriteEndElement();
        }

        private static void WriteCell(XmlWriter writer, int rowIndex, int columnIndex, object value)
        {
            var type = value.GetType();

            if (type == typeof(byte) || type == typeof(short) || type == typeof(int) ||
                type == typeof(long) || type == typeof(float) || type == typeof(double) ||
                type == typeof(decimal))
            {
                WriteNumberCell(writer, rowIndex, columnIndex, Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));
                return;
            }

            if (type == typeof(bool))
            {
                WriteBooleanCell(writer, rowIndex, columnIndex, (bool)value);
                return;
            }

            if (type == typeof(DateTime))
            {
                WriteInlineStringCell(writer, rowIndex, columnIndex, ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss"));
                return;
            }

            if (type == typeof(DateTimeOffset))
            {
                WriteInlineStringCell(writer, rowIndex, columnIndex, ((DateTimeOffset)value).ToString("yyyy-MM-dd HH:mm:ss zzz"));
                return;
            }

            if (type == typeof(byte[]))
            {
                WriteInlineStringCell(writer, rowIndex, columnIndex, Convert.ToBase64String((byte[])value));
                return;
            }

            WriteInlineStringCell(writer, rowIndex, columnIndex, Convert.ToString(value));
        }

        private static void WriteInlineStringCell(XmlWriter writer, int rowIndex, int columnIndex, string value)
        {
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is");
            writer.WriteStartElement("t");
            writer.WriteAttributeString("xml", "space", null, "preserve");
            writer.WriteString(value ?? string.Empty);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteNumberCell(XmlWriter writer, int rowIndex, int columnIndex, string value)
        {
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
            writer.WriteStartElement("v");
            writer.WriteString(value);
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static void WriteBooleanCell(XmlWriter writer, int rowIndex, int columnIndex, bool value)
        {
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", GetCellReference(rowIndex, columnIndex));
            writer.WriteAttributeString("t", "b");
            writer.WriteStartElement("v");
            writer.WriteString(value ? "1" : "0");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static string GetCellReference(int rowIndex, int columnIndex)
        {
            return GetColumnName(columnIndex) + rowIndex.ToString();
        }

        private static string GetColumnName(int columnIndex)
        {
            var name = new StringBuilder();
            while (columnIndex > 0)
            {
                var remainder = (columnIndex - 1) % 26;
                name.Insert(0, (char)('A' + remainder));
                columnIndex = (columnIndex - 1) / 26;
            }

            return name.ToString();
        }

        private static void WriteTextEntry(ZipArchive archive, string name, string content)
        {
            var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(content);
            }
        }

        private static string ContentTypesXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
  <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
  <Default Extension=""xml"" ContentType=""application/xml""/>
  <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
  <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
  <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>";
        }

        private static string RootRelationshipsXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>";
        }

        private static string WorkbookXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""Data"" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>";
        }

        private static string WorkbookRelationshipsXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
  <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
  <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/>
</Relationships>";
        }

        private static string StylesXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <fonts count=""1""><font><sz val=""11""/><name val=""Calibri""/></font></fonts>
  <fills count=""1""><fill><patternFill patternType=""none""/></fill></fills>
  <borders count=""1""><border/></borders>
  <cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>
  <cellXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/></cellXfs>
</styleSheet>";
        }
    }
}
