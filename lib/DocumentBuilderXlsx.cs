namespace dblib
{
    class DocumentBuilderXlsx : DocumentBuilder
    {
        public DocumentBuilderXlsx(string templateFileName, string xmlDataFileName, string outputFileName) :
        base(templateFileName, xmlDataFileName, outputFileName)
        { 
            _documentType = XLSX;

        }

        override public int ProcessTemplate(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            
        }

    } // class DocumentBuilderXlsx
} // namespace dblib