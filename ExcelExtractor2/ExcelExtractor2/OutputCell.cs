using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.SqlServer.Server;
using SpreadsheetLight;

namespace ExcelExtractor2
{
    public class OutputCell
    {
        public int? Row { set; get; }
        public int? Column { set; get; }
        public string Text { set; get; }
        public bool IsStrikeout { set; get; }
        protected SLStyle CellStyle { set; get; }

        public Comment Comment { set; private get; }

        public void SetOutputCell(int row, int column, SLDocument document)
        {
            if (this.Comment != null)
            {
                //// linear gradients
                //var comm = document.CreateComment();
                //// 40% transparency on the first gradient point
                ////comm.GradientFromTransparency = 90;
                //// 80% transparency on the last gradient point
                ////comm.GradientToTransparency = 100;
                //// 45 degrees, so gradient is from top-left to bottom-right
                //comm.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Ocean, 45);
                //comm.SetText(this.TextComment.Replace("\n", " "));
                //document.InsertComment(row, column, comm);


                SLFont font = document.CreateFont();
                font.SetFont("Tahoma", 9);
                font.Bold = true;
                SLRstType rst = document.CreateRstType();

                var lines = this.Comment.InnerText.Split('\n');

                foreach (var line in lines.Where(l => !string.IsNullOrEmpty(l)))
                {
                    rst.AppendText($"{line}\n");
                }

                //rst.AppendText("Karen:\n", font);
                //rst.AppendText("Dan:\n", font);
                //rst.AppendText("1:\n", font);
                //rst.AppendText("2:\n", font);
                //rst.AppendText("Karen:\n", font);
                //rst.AppendText("We have a troll! That is so immature...");
                var comm = document.CreateComment();
                comm.SetText(rst);
                comm.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Ocean, 45);
                comm.AutoSize = true;

                document.InsertComment(row, column, comm);

            }
            if (this.CellStyle == null)
            {
                this.CellStyle = new SLStyle();
                ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF0070C0" };
                this.CellStyle.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 241, 253, 255), SLThemeColorIndexValues.Light2Color);                
                this.CellStyle.Border.SetBottomBorder(BorderStyleValues.Hair, SLThemeColorIndexValues.Dark1Color);
                this.CellStyle.Border.SetTopBorder(BorderStyleValues.Hair, SLThemeColorIndexValues.Dark1Color);
                this.CellStyle.Border.SetLeftBorder(BorderStyleValues.Hair, SLThemeColorIndexValues.Dark1Color);
                this.CellStyle.Border.SetRightBorder(BorderStyleValues.Hair, SLThemeColorIndexValues.Dark1Color);
                this.CellStyle.SetWrapText(true);
            }
            if (this.IsStrikeout)
                this.CellStyle.Font.Strike = true;
            document.SetCellValue(row, column, this.Text);
            document.SetCellStyle(row, column, this.CellStyle);
        }

        public override string ToString()
        {
            return this.Text ?? base.ToString();
        }
    }
}
