using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Task
{
    static class ExcelStyles
    {
        public static void set_worksheet(ExcelPackage excel,string Worksheet)
        {
            excel.Workbook.Worksheets.Add(Worksheet);
        }
        public static ExcelWorksheet get_worksheet(ExcelPackage excel, string Worksheet)
        {
            return excel.Workbook.Worksheets[Worksheet];
        }

        public static void set_style(ExcelWorksheet sheet,List<string[]>header)
        {
            string headerRange = "E7:" + Char.ConvertFromUtf32(header[0].Length + 68) + "7";

            var range = sheet.Cells[headerRange];
         

            range.Style.Font.Bold = true;
            range.Style.Font.Size = 11;
            range.Style.Font.Color.SetColor(Color.White);
          
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.Black);
            var borderstyle = range.Style.Border;
            borderstyle.Left.Style = borderstyle.Right.Style = borderstyle.Left.Style = ExcelBorderStyle.Thin;

            sheet.Cells[headerRange].LoadFromArrays(header);
            range.AutoFilter = true;
            range.AutoFitColumns();



        }
        public static void set_color(ExcelRange range,Color color)
        {
            range.Style.Font.Color.SetColor(color);
        }
         public static void format_money(ExcelRange range,string format)
        {
            range.Style.Numberformat.Format = format;
        }
         public static void load_data(ExcelRange range, List<object[]> Data)
        {
          range.LoadFromArrays(Data);
        }
         public static void set_borders(ExcelRange range)
        {
            
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

    }
}
