using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace ExcelExtractor2
{
    public class TitleCell : OutputCell
    {
        public TitleCell()
        {
            this.CellStyle = new SLStyle();
            //this.CellStyle.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 205, 207, 227), SLThemeColorIndexValues.Light2Color);
            this.CellStyle.Fill.SetGradient(SLGradientShadingStyleValues.FromCenter, System.Drawing.Color.White, System.Drawing.Color.LightSkyBlue);
            this.CellStyle.Border.SetBottomBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetTopBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetLeftBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.Border.SetRightBorder(BorderStyleValues.Medium, SLThemeColorIndexValues.Dark1Color);
            this.CellStyle.SetVerticalAlignment(VerticalAlignmentValues.Center);
            this.CellStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            this.CellStyle.Font.Bold = true;
            this.CellStyle.SetWrapText(true);
        }
    }
}
