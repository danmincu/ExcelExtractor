using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ExcelExtractor2
{
    class Program
    {
        static void Main(string[] args)
        {
            var parameters = @"""c:\temp\pJETS VS PILOTS SKED 06 April 2017 New.xlsx"",""T1 Jets only"",2,1,17,1,1,4,2016 - 10 - 17,30,""c:\temp\text.xlsx""";
            Processor.Run(new string[] { parameters });
        }

        static void Main1(string[] args)
        {
            SLDocument sl = new SLDocument();

            for (int i = 1; i < 20; ++i)
            {
                for (int j = 1; j < 15; ++j)
                {
                    sl.SetCellValue(i, j, string.Format("R{0}C{1}", i, j));
                }
            }

            // copy row 3 to row 21
            sl.CopyRow(3, 21);

            // cuts rows 15 through 18 and paste it to row 23
            // So row 15 is at row 23, row 16 is at row 24 and so on.
            // Default behaviour is to copy-and-paste (false).
            sl.CopyRow(15, 18, 23, true);

            // cuts column 10 and paste it to column 5
            sl.CopyColumn(10, 5, true);

            // copy rows 10 through 12 to column 11
            // So column 10 is at column 11, column 11 is at column 12 and so on.
            // Note that column 10 is currently blank because of the cut-and-paste
            // operation above. So column 11 is also blank.
            sl.CopyColumn(10, 12, 11);

            // You can also copy row and column styles.

            SLStyle rowstyle = sl.CreateStyle();
            rowstyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);

            SLStyle colstyle = sl.CreateStyle();
            colstyle.Fill.SetPattern(PatternValues.Solid, SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Accent6Color);

            // set rows 3 through 10 with the given style
            sl.SetRowStyle(3, 10, rowstyle);
            // set columns 5 through 8 with the given style
            sl.SetColumnStyle(5, 8, colstyle);

            // copy the style from row 5 to rows 12 through 15
            sl.CopyRowStyle(5, 12, 15);
            // copy the style from column 7 to column 2
            sl.CopyColumnStyle(7, 2);

            sl.SaveAs("CopyRowColumn.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
