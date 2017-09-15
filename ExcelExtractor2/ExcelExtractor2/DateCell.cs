using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ExcelExtractor2
{    
    public class DateCell : OutputCell
    {
        public DateCell()
        {
            this.CellStyle = new SLStyle();
            this.CellStyle.SetWrapText(true);
            this.CellStyle.Fill.SetPattern(PatternValues.LightTrellis, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Light2Color);
            this.CellStyle.Border.SetBottomBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetTopBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetLeftBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetRightBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.SetVerticalAlignment(VerticalAlignmentValues.Center);
            this.CellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            this.CellStyle.Font.Bold = true;
        }
    }
}
