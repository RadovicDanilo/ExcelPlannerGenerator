using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

class ExcelPlanner
{
    static string fileName = "Planner.xlsx";

    static int startHour = 8;
    static int endHour = 23;
    static int numberOfNotes = 1;

    static int firstColumnWidth = 10;
    static int columnWidth = 22;
    static int dateRowHeight = 25;
    static int dayRowHeight = 25;
    static int notesRowHeight = 50;
    static int timeRowHeight = 20;

    static ExcelBorderStyle normalBorder = ExcelBorderStyle.Thin;
    static ExcelBorderStyle thickBorder = ExcelBorderStyle.Medium;

    static void CreateExcelPlanner(DateTime startDate, int numDays)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Planner");
            worksheet.Cells.Style.Fill.PatternType = ExcelFillStyle.None; // Remove all default styles

            // === Define Styles ===
            CreateHeaderStyle(package, "HeaderStyle", Color.FromArgb(255, 200, 10));
            CreateHeaderStyle(package, "WorkdayStyle", Color.FromArgb(200, 150, 10), true, 14);
            CreateTimeColumnStyle(package, "TimeColumnStyle", Color.LightBlue);
            CreateHeaderStyle(package, "NoteStyle", Color.FromArgb(255, 250, 150));

            // === User Defined Colors ===
            CreateCellStyle(package, "-NormalCellStyle", Color.White);
            CreateCellStyle(package, "-RedCellStyle", Color.FromArgb(255, 204, 203));
            CreateCellStyle(package, "-BlueCellStyle", Color.FromArgb(173, 216, 230));
            CreateCellStyle(package, "-GreenCellStyle", Color.FromArgb(194, 255, 194));
            CreateCellStyle(package, "-YellowCellStyle", Color.FromArgb(255, 250, 205));
            CreateCellStyle(package, "-OrangeCellStyle", Color.FromArgb(255, 218, 185));
            CreateCellStyle(package, "-PurpleCellStyle", Color.FromArgb(230, 190, 255));
            CreateCellStyle(package, "-GrayCellStyle", Color.FromArgb(224, 224, 224));
            CreateCellStyle(package, "-CyanCellStyle", Color.FromArgb(175, 238, 238));
            CreateCellStyle(package, "-TealCellStyle", Color.FromArgb(180, 255, 255));
            CreateCellStyle(package, "-PinkCellStyle", Color.FromArgb(255, 182, 193));
            CreateCellStyle(package, "-BrownCellStyle", Color.FromArgb(210, 180, 140));

            // === Set Column Widths ===
            worksheet.Column(1).Width = firstColumnWidth;
            for (int col = 2; col <= numDays + 1; col++)
                worksheet.Column(col).Width = columnWidth;

            // === Header Row Setup ===
            worksheet.Cells[1, 1].Value = "Date";
            worksheet.Cells[2, 1].Value = "Day";
            worksheet.Row(1).Height = dateRowHeight;
            worksheet.Row(2).Height = dayRowHeight;
            worksheet.Cells[1, 1, 2, 1].StyleName = "HeaderStyle";

            // === Notes Section ===
            int row = 3;
            while (row < numberOfNotes + 3)
            {
                worksheet.Cells[row, 1].Value = "Note";
                worksheet.Row(row).Height = notesRowHeight;
                worksheet.Cells[row, 1].StyleName = "NoteStyle";

                for (int col = 2; col <= numDays + 1; col++)
                {
                    worksheet.Cells[row, col].StyleName = "NoteStyle";
                }
                row++;
            }

            // === Populate Date & Day Headers ===
            for (int i = 0; i < numDays; i++)
            {
                int col = i + 2;
                DateTime currentDate = startDate.AddDays(i);
                int dayNum = (int)currentDate.DayOfWeek == 0 ? 7 : (int)currentDate.DayOfWeek;

                worksheet.Cells[1, col].Value = currentDate;

                worksheet.Cells[2, col].Value = dayNum;

                if (dayNum <= 5) 
                {
                    worksheet.Cells[1, col].StyleName = "WorkdayStyle";
                    worksheet.Cells[2, col].StyleName = "WorkdayStyle";
                }
                else
                {
                    worksheet.Cells[1, col].StyleName = "HeaderStyle";
                    worksheet.Cells[2, col].StyleName = "HeaderStyle";
                }

                worksheet.Cells[1, col].Style.Numberformat.Format = "dd-mmm-yyyy";
            }

            // === Time Column and Cells ===
            for (int hour = startHour; hour <= endHour; hour++)
            {
                foreach (int minute in new int[] { 0, 30 })
                {
                    worksheet.Row(row).Height = timeRowHeight;
                    worksheet.Cells[row, 1].Value = $"{hour:D2}:{minute:D2}";
                    worksheet.Cells[row, 1].StyleName = "TimeColumnStyle";

                    for (int col = 2; col <= numDays + 1; col++)
                    {
                        worksheet.Cells[row, col].StyleName = "-NormalCellStyle";
                    }

                    row++;
                }
            }

            // === Freeze Header and Time Column ===
            worksheet.View.FreezePanes(3 + numberOfNotes, 2);

            // Mark past dates in red
            var pastDateRule = worksheet.ConditionalFormatting.AddLessThan(worksheet.Cells[1, 2, 1, numDays + 1]);
            pastDateRule.Formula = "TODAY()";
            pastDateRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            pastDateRule.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 204, 203));
            pastDateRule.Style.Font.Color.SetColor(Color.Black);

            // Mark today's date in blue
            var todayDateRule = worksheet.ConditionalFormatting.AddEqual(worksheet.Cells[1, 2, 1, numDays + 1]);
            todayDateRule.Formula = "TODAY()";
            todayDateRule.Style.Fill.PatternType = ExcelFillStyle.Solid;
            todayDateRule.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(173, 216, 230));
            todayDateRule.Style.Font.Color.SetColor(Color.Black);

            string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
            File.WriteAllBytes(filePath, package.GetAsByteArray());
            Console.WriteLine($"Excel Planner Created: {filePath}");
        }
    }

    static void CreateCellStyle(ExcelPackage package, string styleName, Color bgColor)
    {
        var style = package.Workbook.Styles.CreateNamedStyle(styleName);
        style.Style.Fill.PatternType = ExcelFillStyle.Solid;
        style.Style.Fill.BackgroundColor.SetColor(bgColor);
        style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        style.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.Style.Border.Top.Style = normalBorder;
        style.Style.Border.Bottom.Style = normalBorder;
        style.Style.Border.Left.Style = normalBorder;
        style.Style.Border.Right.Style = normalBorder;
    }

    static void CreateHeaderStyle(ExcelPackage package, string styleName, Color bgColor, bool bold = true, int fontSize = 11)
    {
        var style = package.Workbook.Styles.CreateNamedStyle(styleName);
        style.Style.Font.Bold = bold;
        style.Style.Font.Size = fontSize;
        style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        style.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.Style.Fill.PatternType = ExcelFillStyle.Solid;
        style.Style.Fill.BackgroundColor.SetColor(bgColor);
        style.Style.Border.Top.Style = thickBorder;
        style.Style.Border.Bottom.Style = thickBorder;
        style.Style.Border.Left.Style = thickBorder;
        style.Style.Border.Right.Style = thickBorder;
    }

    static void CreateTimeColumnStyle(ExcelPackage package, string styleName, Color bgColor)
    {
        var style = package.Workbook.Styles.CreateNamedStyle(styleName);
        style.Style.Font.Bold = true;
        style.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        style.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        style.Style.Fill.PatternType = ExcelFillStyle.Solid;
        style.Style.Fill.BackgroundColor.SetColor(bgColor);
        style.Style.Border.Top.Style = thickBorder;
        style.Style.Border.Bottom.Style = thickBorder;
        style.Style.Border.Left.Style = thickBorder;
        style.Style.Border.Right.Style = thickBorder;
    }

    static void Main(string[] args)
    {
        if (args.Length < 2 || !DateTime.TryParse(args[0], out DateTime startDate) || !int.TryParse(args[1], out int numDays) || numDays <= 0)
        {
            Console.WriteLine("Usage: ExcelPlanner <YYYY-MM-DD> <num_days>");
            return;
        }

        CreateExcelPlanner(startDate, numDays);
    }
}
