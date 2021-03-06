﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelExtractor2
{
    static class Processor
    {
        const string Separator = "|";
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

                //"c:\TACS\JETS VS PILOTS SKED 10-33-39.xlsx",T1%20Jets%20only;2;1;17;1;1;4|T2%20Jets%20only;2;1;5;1;1;2,2017-10-01,30,"30 DAYS JETS VS PILOTS SCHEDULE.xlsx"

                var arguments = args[0].Split(',').Select(s => s.Trim('\"')).ToArray();

                var fileSource = arguments[0];

                var sheetInstructionsArgs = arguments[1].Split('|');

                var minimDate = DateTime.Parse(arguments[2]);
                var maximDate = DateTime.Parse(arguments[2]) + TimeSpan.FromDays(int.Parse(arguments[3]));
                var destinationFile = path + Path.GetFileName(arguments[4]);

                var sheetInstructionEnum = sheetInstructionsArgs.Select(si =>
                {
                    var sheetArgs = si.Split(';');
                    var sheetName = sheetArgs[0].Replace("%20", " ");
                    var airplaneNames = new Tuple<Position, Position>(new Position { ColumnIndex = int.Parse(sheetArgs[1]), RowIndex = int.Parse(sheetArgs[2]) },
                    new Position { ColumnIndex = int.Parse(sheetArgs[3]), RowIndex = int.Parse(sheetArgs[4]) });
                    var calendarStartPosition = new Position { ColumnIndex = int.Parse(sheetArgs[5]), RowIndex = int.Parse(sheetArgs[6]) };
                    return new SheetInstructions(sheetName, airplaneNames, calendarStartPosition, minimDate, maximDate);
                });

                ExtractAirplaneCalendar(fileSource, destinationFile, sheetInstructionEnum);

                return 0;
            }
            catch (Exception ex)
            {
                System.IO.File.WriteAllText(path + "Exception.txt", ex.ToString());
                throw;
            }
        }

        private static Sheet GetSheetFromWorkSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
            var sheets = workbookPart.Workbook.Sheets.Elements<Sheet>();
            return sheets.FirstOrDefault(s => s.Id.HasValue && s.Id.Value == relationshipId);
        }

        private static void ExtractAirplaneCalendar(string filePath, string saveFilePath, IEnumerable<SheetInstructions> sheetInstructions)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                var xlsxDocument = new SLDocument();

                var workSheetNumber = 0;
                foreach (var sheetInstruction in sheetInstructions)
                {
                    if (workSheetNumber++ == 0)
                    {
                        xlsxDocument.RenameWorksheet(SLDocument.DefaultFirstSheetName, sheetInstruction.SheetName);
                    }
                    else
                    {
                        xlsxDocument.AddWorksheet(sheetInstruction.SheetName);
                    }

                    int currentRow = 1, currentColumn = 1;
                    xlsxDocument.SetColumnWidth(1, 15);
                    xlsxDocument.SetRowHeight(1, 50);
                    // Retrieve a reference to the workbook part.
                    var wbPart = document.WorkbookPart;

                    var theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.ToString().Equals(sheetInstruction.SheetName, StringComparison.OrdinalIgnoreCase));

                    if (theSheet == null)
                        return;

                    var lstRawComments = new Dictionary<string, Comment>();
                    foreach (var sheet in wbPart.WorksheetParts)
                    {
                        var s = GetSheetFromWorkSheet(wbPart, sheet);

                        if (s.Name.HasValue && s.Name.Value.Equals(sheetInstruction.SheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            foreach (var commentsPart in sheet.GetPartsOfType<WorksheetCommentsPart>())
                            {
                                foreach (Comment comment in commentsPart.Comments.CommentList)
                                {
                                    lstRawComments.Add(comment.Reference, comment);
                                }
                            }
                        }
                    }

                    var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    var theRows = wsPart.Worksheet.Descendants<Row>();

                    var options = RegexOptions.None;
                    var regex = new Regex("[ ]{2,}", options);


                    var titleRow = theRows.Skip(sheetInstruction.AirplaneNames.Item1.RowIndex - 1).FirstOrDefault();

                    var airplaneNamesColumnIndex = sheetInstruction.AirplaneNames.Item1.ColumnIndex - 1;
                    var airplaneCount = sheetInstruction.AirplaneNames.Item2.ColumnIndex -
                                        sheetInstruction.AirplaneNames.Item1.ColumnIndex + 1;

                    var titleNames = titleRow.Descendants<Cell>()
                        .Skip(airplaneNamesColumnIndex)
                        .Take(airplaneCount)
                        .Select(t => (regex.Replace(GetCellValue(t, wbPart, out var b).ToString().Trim().Replace("\n", " "), " ")).Replace("|", ""));

                    var titleCellValues = titleRow
                        .Descendants<Cell>()
                        .Skip(airplaneNamesColumnIndex)
                        .Take(airplaneCount)
                        .Select(t => Regex.Replace(t.CellReference, @"[^A-Z]+", String.Empty)).ToList();

                    //clean the airplane names
                    titleNames = titleNames.Select(t => Regex.Replace(t, @"[^0-9a-zA-Z\s\-]+", String.Empty)).Select(t => regex.Replace(t, " "));

                    titleNames = (new string[] { "Date" }).Union(titleNames);
                    titleNames.ToList().ForEach(title =>
                    {
                        xlsxDocument.SetColumnWidth(currentColumn, 20);
                        (new TitleCell() { Text = title }).SetOutputCell(currentRow, currentColumn++, xlsxDocument);
                    });
                    currentRow++;
                    currentColumn = 1;

                    var meaninfulRows = theRows.Skip(sheetInstruction.CalendarStartPosition.RowIndex - 1)
                        .Where(r =>
                        {
                            if (!TryGetDateFromCell(r.Descendants<Cell>().ElementAt(sheetInstruction.CalendarStartPosition.ColumnIndex - 1), wbPart, out var dateTime))
                            {
                                return false;
                            }
                            return dateTime >= sheetInstruction.MinimDate && dateTime < sheetInstruction.MaximDate;
                        }).ToList();

                    string line = null;
                    foreach (var row in meaninfulRows)
                    {

                        var cells = row.Descendants<Cell>().ToList();
                        var values = row.Descendants<Cell>()

                            // this list in not linear - it skips columns with no values therefore it needs to be used by reference
                            //.Skip(airplaneNamesColumnIndex)
                            //.Take(airplaneCount)
                            .Select(t =>
                            {
                                var ocell = new OutputCell()
                                {
                                    Text = GetCellValue(t, wbPart, out var isStrikeout).ToString().Trim().Replace("\n", " "),
                                    IsStrikeout = isStrikeout
                                };
                                if (lstRawComments.ContainsKey(t.CellReference.ToString()))
                                {
                                    // ocell.TextComment = lstComments[t.CellReference];
                                    ocell.Comment = lstRawComments[t.CellReference];
                                }
                                return System.Tuple.Create(Regex.Replace(t.CellReference, @"[^A-Z]+", String.Empty), ocell/*v*/);
                            }).ToDictionary(t => t.Item1, t => t.Item2);

                        var vals = new List<OutputCell>();
                        foreach (var letter in titleCellValues)
                        {
                            if (values.ContainsKey(letter))
                                vals.Add(values[letter]);
                            else
                                vals.Add(new OutputCell() { Text = null });
                        }

                        TryGetDateFromCell(cells.ElementAt(sheetInstruction.CalendarStartPosition.ColumnIndex - 1), wbPart, out var dateTime);
                        var cellValues = (new OutputCell[] { new DateCell() { Text = dateTime.ToString("yyyy-MM-dd") } }).Union(vals);

                        cellValues.ToList().ForEach(a => a.SetOutputCell(currentRow, currentColumn++, xlsxDocument));
                        currentRow++;
                        currentColumn = 1;
                    }
                }
                xlsxDocument.SelectWorksheet(sheetInstructions.First().SheetName);
                xlsxDocument.SaveAs(saveFilePath);
            }
        }

        private static bool TryGetDateFromCell(Cell cell, WorkbookPart wbPart, out DateTime dateTime)
        {
            dateTime = DateTime.MinValue;

            var dtobj = GetCellValue(cell, wbPart, out var b);
            if (!(dtobj is DateTime))
            {
                if (!Int32.TryParse(dtobj.ToString(), out var potentialDateTimeValue))
                    return false;

                // I determined these values using the Excel DateValue function - whatever that is; +1 mean plus one day therefore the next code
                const int firstJanuary1990 = 32874;
                const int firstJanuary2100 = 73051;

                if (potentialDateTimeValue > firstJanuary2100 || potentialDateTimeValue < firstJanuary1990)
                    return false;

                dateTime = (new DateTime(1990, 1, 1)).AddDays(potentialDateTimeValue - firstJanuary1990);
                return true;
            }

            dateTime = (DateTime)dtobj;
            return true;
        }

        private static object GetCellValue(Cell theCell, WorkbookPart wbPart, out bool isStrike)
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

                var cellFormats = wbPart.WorkbookStylesPart.Stylesheet.CellFormats;
                var numberingFormats = wbPart.WorkbookStylesPart.Stylesheet.NumberingFormats;


                bool isDate = false;
                var styleIndex = (int)theCell.StyleIndex.Value;
                var cellFormatt = (CellFormat)cellFormats.ElementAt(styleIndex);

                if (cellFormatt.NumberFormatId != null)
                {
                    var numberFormatId = cellFormatt.NumberFormatId.Value;
                    var numberingFormat = numberingFormats.Cast<NumberingFormat>()
                        .SingleOrDefault(f => f.NumberFormatId.Value == numberFormatId);

                    // Here's yer string! Example: $#,##0.00_);[Red]($#,##0.00)
                    if (numberingFormat != null && numberingFormat.FormatCode.Value.Contains("mmm"))
                    {
                        string formatString = numberingFormat.FormatCode.Value;
                        isDate = true;
                    }
                }


                int dateInteger;
                if (!string.IsNullOrEmpty(value.ToString()) && int.TryParse(value.ToString(), out dateInteger) && attrib.Count >= 1 && (dateTypes.Contains(attrib[1].Value) || isDate)) //attrib.Count >= 1 && (attrib[1].Value == "611" || attrib[1].Value == "108"))                            {
                {
                    value = (new DateTime(1899, 12, 30).Add(new TimeSpan(dateInteger, 0, 0, 0)));//.ToString("yyyy-MM-dd");
                }
            }
            return value;

        }
    }
}
