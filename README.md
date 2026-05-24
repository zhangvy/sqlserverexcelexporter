# SQL Server 导出 Excel 工具

这是一个 `.NET Framework 4.6.2` WinForms 桌面程序，用于按日期范围把 SQL Server 表导出为 `.xlsx`。

## 使用方式

1. 修改 `src/SqlServerExcelExporter.App/App.config`：

   ```xml
   <connectionStrings>
     <add name="DefaultConnection"
          connectionString="Server=YOUR_SERVER;Database=YOUR_DATABASE;User Id=YOUR_USER;Password=YOUR_PASSWORD;TrustServerCertificate=True;"
          providerName="System.Data.SqlClient" />
   </connectionStrings>
   ```

2. 启动程序后输入表名，例如 `dbo.Orders`。
3. 点击“查询表结构”，确认字段和日期列。
4. 选择日期列、起始日期、结束日期、导出方式和输出目录。
5. 点击“开始导出”。

## 日期规则

- 界面上的结束日期是包含的。
- SQL 查询内部使用左闭右开：`>= 起始日期` 且 `< 结束日期 + 1天`。
- 不分组：整个日期范围导出为一个 Excel。
- 按周：周一到周日为一组，首尾周按用户日期裁剪。
- 按月：自然月为一组，首尾月按用户日期裁剪。

## 验证

当前环境中 `dotnet build` 因无法访问本机 Microsoft SDK 缓存目录而失败。已用 .NET Framework 编译器直接验证：

```powershell
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

验证内容包括核心库编译、WinForms 程序编译，以及测试程序 8 项规则测试。
