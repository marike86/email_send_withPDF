using System.Text;
using ClosedXML.Excel;

namespace PdfMailer.Core;

public static class ResultExporter
{
    private static string Safe(string value) => value.Length > 0 && "=+-@\t\r".Contains(value[0]) ? "'" + value : value;

    public static void Export(IReadOnlyList<MailJob> jobs, string path)
    {
        var rows = new List<string[]> { new[] { "文件名", "学号", "姓名", "邮箱", "状态", "说明" } };
        rows.AddRange(jobs.Select(job => new[] { job.Filename, job.StudentId, job.Name, job.Email, job.StatusLabel, job.Error }.Select(Safe).ToArray()));
        if (Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.AddWorksheet("检查与发送结果");
            for (var row = 0; row < rows.Count; row++)
                for (var column = 0; column < rows[row].Length; column++)
                    sheet.Cell(row + 1, column + 1).Value = rows[row][column];
            sheet.Row(1).Style.Font.SetBold().Font.SetFontColor(XLColor.White);
            sheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0067C0");
            sheet.SheetView.FreezeRows(1);
            sheet.RangeUsed()!.SetAutoFilter();
            var widths = new[] { 36, 20, 18, 36, 14, 65 };
            for (var column = 1; column <= widths.Length; column++) sheet.Column(column).Width = widths[column - 1];
            workbook.SaveAs(path);
        }
        else
        {
            using var output = new StreamWriter(path, false, new UTF8Encoding(true));
            foreach (var row in rows) output.WriteLine(string.Join(",", row.Select(value => "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"")));
        }
    }
}
