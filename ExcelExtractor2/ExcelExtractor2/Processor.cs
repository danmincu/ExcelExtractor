using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Configuration;
using SpreadsheetLight;

namespace ExcelExtractor2
{
 

    class Processor
    {
        const string separator = "|";
        private static string[] dateTypes;
        public static int Run(string[] args)
        {
            dateTypes = ConfigurationManager.AppSettings["dateTypes"].Split(',');

            var path = Path.GetTempPath();
            Console.WriteLine(path);
            try
            {
                if (args.Length == 0)
                {
                    Console.WriteLine("*******************************************************************************");
                    Console.WriteLine("************************ EXCEL DATA EXTRACTOR *********************************");
                    Console.WriteLine("*******************************************************************************");
                    Console.WriteLine("Usage example: ExcelExtractor.exe \"c:\\temp\\JETS VS PILOTS SKED 06 April 2017 New.xlsx\",\"T1 Jets only\",2,1,17,1,1,4,2017-03-01,30,\"c:\\temp\\text.xlsx\"");
                    Console.WriteLine("*******************************************************************************");
                    Console.WriteLine("Parameters: excel file, sheet name, airplane names start row, airplane names start column, stop row, stop column, calendar start row, calendar start column, min date, days to extract, output xlsx file");
                    return 2;
                }

                // Example of arguments "c:\\temp\\source.xlsx","T1 Jets only",2,1,17,1,1,4,2017-03-01,30,"c:\\temp\\text.txt";
                // args[0] = "\"c:\\temp\\JETS VS PILOTS SKED 06 April 2017 New.xlsx\",\"T1 Jets only\",2,1,17,1,1,4,2017-03-01,30,\"c:\\temp\\text.txt\""; 

                var arguments = args[0].Split(',').Select(s => s.Trim('\"')).ToArray();
                var fileSource = arguments[0]; //ConfigurationManager.AppSettings["excel file"];
                var sheetName = arguments[1];
                var airplaneNames = new Tuple<Position, Position>(new Position { ColumnIndex = int.Parse(arguments[2]), RowIndex = int.Parse(arguments[3]) },
                    new Position { ColumnIndex = int.Parse(arguments[4]), RowIndex = int.Parse(arguments[5]) });
                var calendarStartPosition = new Position { ColumnIndex = int.Parse(arguments[6]), RowIndex = int.Parse(arguments[7]) };
                var minimDate = DateTime.Parse(arguments[8]);
                var maximDate = DateTime.Parse(arguments[8]) + TimeSpan.FromDays(int.Parse(arguments[9]));
                var destinationFile = path + Path.GetFileName(arguments[10]);

                ExtractAirplaneCalendar(sheetName, fileSource, destinationFile, airplaneNames, calendarStartPosition, minimDate, maximDate);
                return 0;
            }
            catch (Exception ex)
            {
                System.IO.File.WriteAllText(path + "Exception.txt", ex.ToString());
                return 1;
            }
        }

        public static Sheet GetSheetFromWorkSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            string relationshipId = workbookPart.GetIdOfPart(worksheetPart);
            IEnumerable<Sheet> sheets = workbookPart.Workbook.Sheets.Elements<Sheet>();
            return sheets.FirstOrDefault(s => s.Id.HasValue && s.Id.Value == relationshipId);
        }

        static void ExtractAirplaneCalendar(string sheetName, string filePath, string saveFilePath, Tuple<Position, Position> airplaneNames, Position calendarStartPosition, DateTime minimDate, DateTime maximDate)
        {

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                int currentRow = 1, currentColumn = 1;
                SLDocument xlsxDocument = new SLDocument();
                xlsxDocument.SetColumnWidth(1, 15);
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (theSheet == null)
                    return;

                Dictionary<string, string> lstComments = new Dictionary<string, string>();
                foreach (WorksheetPart sheet in wbPart.WorksheetParts)
                {
                    var s = GetSheetFromWorkSheet(wbPart, sheet);

                    if (s.Name == sheetName)
                    {
                        foreach (WorksheetCommentsPart commentsPart in sheet.GetPartsOfType<WorksheetCommentsPart>())
                        {
                            foreach (Comment comment in commentsPart.Comments.CommentList)
                            {
                                lstComments.Add(comment.Reference, comment.InnerText);
                            }
                        }
                    }
                }

                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                var theRows = wsPart.Worksheet.Descendants<Row>();

                RegexOptions options = RegexOptions.None;
                Regex regex = new Regex("[ ]{2,}", options);


                var titleRow = theRows.Skip(airplaneNames.Item1.RowIndex - 1).FirstOrDefault();
                var titleNames = titleRow.Descendants<Cell>().Skip(airplaneNames.Item1.ColumnIndex - 1).Take(airplaneNames.Item2.ColumnIndex - 1)
                    .Select(t => (regex.Replace(GetCellValue(t, wbPart, out var b).ToString().Trim().Replace("\n", " "), " ")).Replace("|", ""));
                var titleCellValues = titleRow.Descendants<Cell>().Skip(airplaneNames.Item1.ColumnIndex - 1).Take(airplaneNames.Item2.ColumnIndex - 1).
                    Select(t => Regex.Replace(t.CellReference, @"[^A-Z]+", String.Empty)).ToList();

                //clean the airplane names
                titleNames = titleNames.Select(t => Regex.Replace(t, @"[^0-9a-zA-Z\s\-]+", String.Empty)).Select(t => regex.Replace(t, " "));

                titleNames = (new string[] { "Date" }).Union(titleNames);
                titleNames.ToList().ForEach(title =>
                {
                    (new TitleCell() { Text = title }).SetOutputCell(currentRow, currentColumn++, xlsxDocument);
                });
                currentRow++;
                currentColumn = 1;

                var meaninfulRows = theRows.Skip(calendarStartPosition.RowIndex - 1)
                    .Where(r =>
                    {
                        var dtobj = GetCellValue(r.Descendants<Cell>().ElementAt(calendarStartPosition.ColumnIndex - 1), wbPart, out var b);
                        if (!(dtobj is DateTime))
                            return false;
                        var dt = (DateTime)dtobj;
                        return dt >= minimDate && dt < maximDate;
                    }).ToList();
                string line = null;
                foreach (var row in meaninfulRows)
                {

                    var cells = row.Descendants<Cell>().ToList();
                    var values = row.Descendants<Cell>()
                        .Skip(airplaneNames.Item1.ColumnIndex - 1)
                        .Take(airplaneNames.Item2.ColumnIndex - 1)
                        .Select(t =>
                        {
                            var ocell = new OutputCell()
                            {
                                Text = GetCellValue(t, wbPart, out var isStrikeout).ToString().Trim().Replace("\n", " "),
                                IsStrikeout = isStrikeout
                            };
                            if (lstComments.ContainsKey(t.CellReference.ToString()))
                            {
                                ocell.TextComment = lstComments[t.CellReference].Replace("\n", " ");
                            }
                            return System.Tuple.Create(Regex.Replace(t.CellReference, @"[^A-Z]+", String.Empty), ocell/*v*/);
                        }).ToDictionary(t => t.Item1, t => t.Item2);

                    List<OutputCell> vals = new List<OutputCell>();
                    foreach (var letter in titleCellValues)
                    {
                        if (values.ContainsKey(letter))
                            vals.Add(values[letter]);
                        else
                            vals.Add(new OutputCell() { Text = null });
                    }

                    line = string.Join(separator, vals);
                    var DateTime = (DateTime)GetCellValue(cells.ElementAt(calendarStartPosition.ColumnIndex - 1), wbPart, out var b);
                    var cellValues = (new OutputCell[] { new DateCell() { Text = DateTime.ToString("yyyy-MM-dd") } }).Union(vals);

                    cellValues.ToList().ForEach(a => a.SetOutputCell(currentRow, currentColumn++, xlsxDocument));              
                    currentRow++;
                    currentColumn = 1;
                }

                xlsxDocument.SaveAs(saveFilePath);
            }
        }

        static object GetCellValue(Cell theCell, WorkbookPart wbPart, out bool isStrike)
        {
            isStrike = false;
            var attrib = theCell.GetAttributes();
            Object value = theCell.InnerText;

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (theCell.DataType != null)
            {
                switch (theCell.DataType.Value)
                {
                    case CellValues.Date:

                        TimeSpan datefromexcel = new TimeSpan(int.Parse(value.ToString()), 0, 0, 0);
                        value = (new DateTime(1899, 12, 30).Add(datefromexcel)).ToString();
                        break;
                    case CellValues.SharedString:

                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable =
                            wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (stringTable != null)
                        {
                            var v = stringTable.SharedStringTable
                                .ElementAt(int.Parse(value.ToString()));

                            //detect whther the text was part of striked from string item
                            foreach (Strike strike in stringTable.SharedStringTable
                               .ElementAt(int.Parse(value.ToString())).Descendants<Strike>())
                            {
                                if (strike.Val == null || strike.Val != false)
                                {
                                    isStrike = true;
                                }
                            }


                            //detect whther the text was striked from cell style
                            var cellFormat = (CellFormat)wbPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(int.Parse(theCell.StyleIndex));
                            var font = wbPart.WorkbookStylesPart.Stylesheet.Fonts.ElementAt(int.Parse(cellFormat.FontId));
                            foreach (Strike strike in font.Descendants<Strike>())
                            {
                                if (strike.Val == null || strike.Val != false)
                                {
                                    isStrike = true;
                                }
                            }


                            value = v.InnerText;



                        }
                        break;

                    case CellValues.Boolean:
                        switch (value.ToString())
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }
            else //null cell data type
            {
                int dateInteger;
                if (!string.IsNullOrEmpty(value.ToString()) && int.TryParse(value.ToString(), out dateInteger) && attrib.Count >= 1 && dateTypes.Contains(attrib[1].Value)) //attrib.Count >= 1 && (attrib[1].Value == "611" || attrib[1].Value == "108"))                            {
                {
                    value = (new DateTime(1899, 12, 30).Add(new TimeSpan(dateInteger, 0, 0, 0)));//.ToString("yyyy-MM-dd");
                }
            }
            return value;

        }
    }
}
