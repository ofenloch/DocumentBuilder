namespace dblib
{
    class DocumentBuilderXlsx : DocumentBuilder
    {
        public DocumentBuilderXlsx(string templateFileName, string xmlDataFileName, string outputFileName) :
        base(templateFileName, xmlDataFileName, outputFileName)
        { 
            _documentType = XLSX;

        }
    } // class DocumentBuilderXlsx
} // namespace dblib