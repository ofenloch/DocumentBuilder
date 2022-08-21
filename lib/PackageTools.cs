namespace dblib
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.IO.Packaging;
    using System.Text;

    class PackageTools
    {
        // To create a new package as a Word document.
        public static string CreateNewWordDocument(string document)
        {
            // this is from https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-package
            document = Path.GetFullPath(document);
            Utilities.CreateDirectory(Path.GetDirectoryName(document));
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
            Utilities.CreateDirectory(Path.GetDirectoryName(filepath));
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

        static public void UnpackPackage(string filePath, string targetDirectory = "")
        {
            // open the package for reading
            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                if (targetDirectory.Length < 1)
                {
                    targetDirectory = filePath + "-unpacked";
                }
                // unpack the package to target directory
                UnpackPackage(package, targetDirectory);
                // close the package
                package.Close();
            } // using ...
        } // static public void unpackPackage(string filePath, string targetDirectory = "")

        // unpack the given package to the filesystem in the given directory and format the XML parts nicely
        // the directory and the required subdirectories will be created
        // the given packagg is not modified (you may pass a read-only file)
        static public void UnpackPackage(Package package, string targetDirectory)
        {
            // create the target directory
            Utilities.CreateDirectory(targetDirectory);
            // get all package parts contained in the package
            PackagePartCollection packageParts = package.GetParts();
            // loop over the package's parts and process each part
            foreach (PackagePart packagePart in packageParts)
            {
                Uri uri = packagePart.Uri;
                Console.WriteLine("Package part: {0}", uri);
                // construct a file name:
                string fileName = targetDirectory + uri;
                string dirName = Path.GetDirectoryName(fileName);
                dirName = Utilities.CreateDirectory(dirName);
                Console.WriteLine("  file {0}", fileName);
                if (packagePart.ContentType.EndsWith("xml"))
                {
                    // open the XML from the Page Contents part
                    System.Xml.Linq.XDocument packagePartXML = GetXMLFromPart(packagePart);
                    // and save it to the file
                    // (the result is fine for me, but you might wanna use an XMLWriter for better/nicer formatting)
                    packagePartXML.Save(fileName);
                }
                else
                {
                    // just save the non XML as it is
                    FileStream newFileStrem = new FileStream(fileName, FileMode.Create);
                    packagePart.GetStream().CopyTo(newFileStrem);
                }
            }
        } // static public void unpackPackage(Package package, string targetDirectory)

        static private System.Xml.Linq.XDocument GetXMLFromPart(PackagePart packagePart)
        {
            System.Xml.Linq.XDocument partXml = null;
            // read the XML document from the package part's stream
            Stream partStream = packagePart.GetStream();
            partXml = System.Xml.Linq.XDocument.Load(partStream);
            // Important: Close the stream or we will get an exception when writing the xml back to the package part.
            partStream.Close();
            return partXml;
        } // static private XDocument GetXMLFromPart(PackagePart packagePart)

    } // class PackageTools

} // namespace dblib