using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExtractor
{
    public static class OpenXMLUtilities
    {

        public static void CopySheetToNewFile(string filename, string sheetName, string tempFileName)
        {
            File.Copy(filename, tempFileName, true);
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(tempFileName, true))
            {
                WorkbookPart wbPart = mySpreadsheet.WorkbookPart;
                var theSheet = wbPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault((s) => s.Name.InnerText.ToUpper() == sheetName.ToUpper());
                if (theSheet == null)
                    return;
                
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                //var theRows = wsPart.Worksheet.Descendants<Row>().Skip(30).ToList();
                //while (theRows.Count() > 0)
                //{
                //    theRows.First().Remove();
                //}
                


                


                IEnumerable<Sheet> sheets = mySpreadsheet.WorkbookPart.Workbook.Descendants<Sheet>()
                    .Where((s) => s.Name.InnerText.ToUpper() != sheetName.ToUpper());
                while (sheets.Count() > 0)
                {
                    sheets.First().Remove();
                }
                mySpreadsheet.WorkbookPart.Workbook.Save();
            }
        }



        static int tableId = 0;
        static public void CopySheet(string filename, string sheetName, string clonedSheetName)
        {
            //Open workbook
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(filename, true))
            {
                WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;
                //Get the source sheet to be copied
                WorksheetPart sourceSheetPart = GetWorkSheetPart(workbookPart, sheetName);


                //Take advantage of AddPart for deep cloning

                SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), mySpreadsheet.DocumentType);

                WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();

                WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);

                //Add cloned sheet and all associated parts to workbook

                WorksheetPart clonedSheet = workbookPart.AddPart<WorksheetPart>(tempWorksheetPart);



                //Table definition parts are somewhat special and need unique ids...so let's make an id based on count

                int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();

                tableId = numTableDefParts;

                //Clean up table definition parts (tables need unique ids)

                if (numTableDefParts != 0)

                    FixupTableParts(clonedSheet, numTableDefParts);

                //There should only be one sheet that has focus

                CleanView(clonedSheet);

                
                //Add new sheet to main workbook part

                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

                Sheet copiedSheet = new Sheet();

                copiedSheet.Name = clonedSheetName;

                copiedSheet.Id = workbookPart.GetIdOfPart(clonedSheet);

                copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;

                sheets.Append(copiedSheet);

                //Save Changes

                workbookPart.Workbook.Save();



            }
        }

        static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)

        {

            //Get the relationship id of the sheetname

            string relId = workbookPart.Workbook.Descendants<Sheet>()

            .Where(s => s.Name.Value.Equals(sheetName))

            .First()

            .Id;

            return (WorksheetPart)workbookPart.GetPartById(relId);

        }


        static void CleanView(WorksheetPart worksheetPart)

        {

            //There can only be one sheet that has focus

            SheetViews views = worksheetPart.Worksheet.GetFirstChild<SheetViews>();

            if (views != null)

            {

                views.Remove();

                worksheetPart.Worksheet.Save();

            }

        }

        static void FixupTableParts(WorksheetPart worksheetPart, int numTableDefParts)
        {
            //Every table needs a unique id and name
            foreach (TableDefinitionPart tableDefPart in worksheetPart.TableDefinitionParts)
            {
                tableId++;
                tableDefPart.Table.Id = (uint)tableId;
                tableDefPart.Table.DisplayName = "CopiedTable" + tableId;
                tableDefPart.Table.Name = "CopiedTable" + tableId;
                tableDefPart.Table.Save();
            }
        }

    }
}
