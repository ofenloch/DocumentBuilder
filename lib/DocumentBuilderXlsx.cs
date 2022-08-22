using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace dblib
{
    public class DocumentBuilderXlsx : DocumentBuilder
    {
        protected static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public DocumentBuilderXlsx(string templateFileName, string xmlDataFileName, string outputFileName) :
        base(templateFileName, xmlDataFileName, outputFileName)
        {
            _documentType = XLSX;

        }

        public override int
        ProcessTemplate(
            string templateFielName,
            string xmlDataFileName,
            string outputFileName
        )
        {
            int error = 0;
            Console.WriteLine("DocumentBuilderExcel: doing the Excel stuff ...");

            // copy the template to the output file and work with copy (never use the template)
            File.Delete(outputFileName);
            File.Copy(templateFielName, outputFileName);

            XmlDataFileParser xmlParser = new XmlDataFileParser(xmlDataFileName);
            DataFieldList dataFields = xmlParser.GetDataFields();

            // Open a SpreadsheetDocument for read-write access based on a filepath.
            SpreadsheetDocument document =
                SpreadsheetDocument.Open(outputFileName, true);
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData =
                worksheetPart.Worksheet.Elements<SheetData>().First();
            string cellContents = null;
            int iRow = 0;

            // We search for something like "${replace_me_by_data}"
            // where "replace_me_by_data" is the data field's name.
            const string pattern = @"\$\{.*\}"; 

            // Define a regular expression for repeated words.
            Regex regex = new Regex(pattern, RegexOptions.Compiled);
            foreach (Row row in sheetData.Elements<Row>())
            {
                iRow++;
                Logger.Debug("row " + iRow + " : \n");
                int iCell = 0;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    iCell++;

                    // Type cellType = cell.GetType();
                    // cellContents = cell.CellValue.Text;
                    cellContents = GetCellValue(cell, workbookPart, worksheetPart);
                    if (cellContents == null || cellContents.Length < 1)
                    {
                        continue;
                    }
                    //Console.Write(cellContents + ", ");
                    // Find field experessions like "f(project_number)"
                    MatchCollection matches = regex.Matches(cellContents);
                    foreach (Match match in matches)
                    {
                        string fieldName = match.Value;
                        Logger.Debug(
                            "found " + fieldName + " in row " + iRow + ", cell " + iCell);
                        fieldName = fieldName.Remove(0, 2);
                        fieldName = fieldName.Remove(fieldName.Length - 1, 1);
                        Logger.Debug("  -> fieldName \"" + fieldName + "\"\n");

                        DataField df = dataFields.GetById(fieldName);
                        if (df != null)
                        {
                            Logger.Debug("  -> found field \"" + df.GetXmlId() + "\" in xml file\n");
                            DataTypes type = df.GetDataType();
                            if (type == DataTypes.NUMBER)
                            {
                                cell.CellValue = new CellValue(df.GetDoubleValue());
                                cell.DataType =
                                    new EnumValue<CellValues>(CellValues.Number);
                            }
                            else if (type == DataTypes.BOOLEAN)
                            {
                                cell.CellValue =
                                    new CellValue(df.GetBooleanValue());
                                cell.DataType =
                                    new EnumValue<CellValues>(CellValues.Boolean);
                            }
                            else
                            {
                                cell.CellValue = new CellValue(df.GetValue());
                                cell.DataType =
                                    new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                        else
                        {
                            // The field was not found in the xml file.
                            // So we make the cell an empty one.
                            cell.CellValue = new CellValue("");
                            cell.DataType =
                                    new EnumValue<CellValues>(CellValues.String);
                        }
                    } // foreach (Match match in matches)
                } // foreach (Cell cell in row.Elements<Cell>())
                Console.WriteLine();
            } // foreach (Row row in sheetData.Elements<Row>())
            workbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            workbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
            worksheetPart.Worksheet.Save();
            document.Save();
            document.Close();
            Xlsx2Csv(outputFileName, Path.GetDirectoryName(outputFileName));
            return error;
        } // ProcessTemplate(string templateFielName, string xmlDataFileName, string outputFileName)

        static public int Xlsx2Csv(string xlsxFileName, string targetDirectory, char delimiter = ',')
        {
            int error = 0;
            xlsxFileName = Path.GetFullPath(xlsxFileName);
            string xlsxFileBaseName = Path.GetFileNameWithoutExtension(xlsxFileName);
            targetDirectory = Path.GetFullPath(targetDirectory);
            // Open the document in read-only mode:
            using (SpreadsheetDocument spreadsheetDocument =
                SpreadsheetDocument.Open(xlsxFileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                int iSheet = 0;
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    SheetData sheetData =
                        worksheetPart.Worksheet.Elements<SheetData>().First();
                    string cellContents = null;
                    int iRow = 0;
                    string outFileName = targetDirectory + '/' + xlsxFileBaseName + "_Sheet_" + iSheet + ".csv";
                    StreamWriter outFile = new StreamWriter(outFileName);
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        int iCell = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            cellContents = GetCellValue(cell, workbookPart, worksheetPart);
                            if (cellContents == null || cellContents.Length < 1)
                            {
                                cellContents = "";
                            }
                            outFile.Write(cellContents + delimiter);
                            iCell++;
                        } // foreach (Cell cell in row.Elements<Cell>())
                        outFile.Write('\n');
                        iRow++;
                    } // foreach (Row row in sheetData.Elements<Row>())
                    outFile.Close();
                    iSheet++;
                } // foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)

            } // using (SpreadsheetDocument spreadsheetDocument ...

            return error;
        } // static public int Xlsx2Csv(string xlsxFileName, string targetDirectory, char delimiter = ',')

        // Retrieve the value of a cell as string
        public static string
        GetCellValue(
            Cell theCell,
            WorkbookPart workbookPart,
            WorksheetPart worksheetPart
        )
        {
            // If the cell doesn't exist return an empty string
            if (theCell == null)
            {
                return null;
            }

            // This is our return value:
            string value = theCell.InnerText;

            // If the cell doesn't exist, return an empty string.
            if (value.Length < 1)
            {
                return null;
            }

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
                    case CellValues.SharedString:
                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable =
                            workbookPart
                                .GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                        // If the shared string table is missing, something
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in
                        // the table.
                        if (stringTable != null)
                        {
                            value =
                                stringTable
                                    .SharedStringTable
                                    .ElementAt(int.Parse(value))
                                    .InnerText;
                        }
                        break;
                    case CellValues.Boolean:
                        switch (value)
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
            return value;
        }

        // Retrieve the value of a cell, given a file name, sheet name,
        // and address name.
        public static string
        GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (
                SpreadsheetDocument document =
                    SpreadsheetDocument.Open(fileName, false)
            )
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet =
                    wbPart
                        .Workbook
                        .Descendants<Sheet>()
                        .Where(s => s.Name == sheetName)
                        .FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell
                // whose address matches the address you supplied.
                Cell theCell =
                    wsPart
                        .Worksheet
                        .Descendants<Cell>()
                        .Where(c => c.CellReference == addressName)
                        .FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell.InnerText.Length > 0)
                {
                    value = theCell.InnerText;

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
                            case CellValues.SharedString:
                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart
                                        .GetPartsOfType<SharedStringTablePart>()
                                        .FirstOrDefault();

                                // If the shared string table is missing, something
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable
                                            .SharedStringTable
                                            .ElementAt(int.Parse(value))
                                            .InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (value)
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
                }
            }
            return value;
        } // public static string GetCellValue(string fileName, string sheetName, string addressName)


    } // class DocumentBuilderXlsx
} // namespace dblib