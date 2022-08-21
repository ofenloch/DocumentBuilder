namespace dblib
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Text;

    class PackageTools
    {
        // To create a new package as a Word document.
        public static string CreateNewWordDocument(string document)
        {
            // this is from https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-package
            document = Path.GetFullPath(document);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(document, WordprocessingDocumentType.Document))
            {
                // Set the content of the document so that Word can open it.
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();

                SetMainDocumentContent(mainPart);
                return document;
            }
        }

        // Set the content of MainDocumentPart.
        public static void SetMainDocumentContent(MainDocumentPart part)
        {
            // this is from https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-package
            const string docXml =
             @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
                <w:document xmlns:w=""https://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:body>
                        <w:p>
                            <w:r>
                                <w:t>Hello world!</w:t>
                            </w:r>
                        </w:p>
                    </w:body>
                </w:document>";

            using (Stream stream = part.GetStream())
            {
                byte[] buf = (new UTF8Encoding()).GetBytes(docXml);
                stream.Write(buf, 0, buf.Length);
            }
        }

        public static string CreateNewSpreadsheetWorkbook(string filepath)
        {
            // this is from https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name

            filepath = Path.GetFullPath(filepath);

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();
            return filepath;
        }
    } // class PackageTools

} // namespace dblib