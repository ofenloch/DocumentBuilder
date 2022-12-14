namespace dblib
{

    public abstract class DocumentBuilder
    {

        protected static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        protected string _outputDirectory = "";
        protected string _outputFileName = "";
        protected string _outputFileBaseName = "";
        public class DocumentType
        {
            // the file extension (max 4 characters), e.g. xlsx, docs, pptx, ...
            string _extension = "";
            string _name = "";
            ushort _index = 0;
            public DocumentType(string extension, string name, ushort index)
            {
                if (extension.Length > 4)
                {
                    throw new ArgumentException("extensions may only be 4 characters");
                }
                _extension = extension.ToLower();
                _name = name;
                _index = index;
            }
            public string GetExtension()
            {
                return _extension;
            }
            public string GetName()
            {
                return _name;
            }
            public ushort GetIndex()
            {
                return _index;
            }
            public bool SameAs(string extension)
            {
                return _extension == extension.ToLower();
            }
            public bool SameAs(DocumentType documentType)
            {
                return this._extension == documentType.GetExtension();
            }
        } // public class DocumentType

        public static DocumentType NONE = new("----", "NONE", 0);
        public static DocumentType DOCX = new("docx", "Word Document", 0);
        public static DocumentType XLSX = new("xlsx", "Excel Document", 0);

        protected DocumentType _documentType = NONE;
        public DocumentBuilder(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            _outputFileName = Path.GetFullPath(outputFileName);
            _outputFileBaseName = Path.GetFileNameWithoutExtension(_outputFileName);
            _outputDirectory = Path.GetDirectoryName(_outputFileName);
            Utilities.CreateDirectory(_outputDirectory);
            Logger.Info("DocumentBuilder: templateFileName is {0}", templateFileName);
            Logger.Info("DocumentBuilder: xmlDataFileName is {0}", xmlDataFileName);
            Logger.Info("DocumentBuilder: outputFileName is {0}", outputFileName);
        }

        public static string CreateNewWordDocument(string fileName)
        {
            return PackageTools.CreateNewWordDocument(fileName);
        }

        public static string CreateNewExcelDocument(string fileName)
        {
            return PackageTools.CreateNewSpreadsheetWorkbook(fileName);
        }

        public static void UnpackPackage(string fileName, string targetDirectory = "")
        {
            PackageTools.UnpackPackage(fileName, targetDirectory);
        }

        public virtual int ProcessTemplate(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            // 'DocumentBuilder.ProcessTemplate(string, string, string)' must declare a body because it is not marked abstract, extern, or partial
            Logger.Error("DocumentBuilder: DocumentBuilder.ProcessTemplate should never ever be called!");
            throw new InvalidOperationException("calling DocumentBuilder.ProcessTemplate(...) is not allowed");
        }
    } // public class DocumentBuilder

} // namespace dblib
