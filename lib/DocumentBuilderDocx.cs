namespace dblib
{
    class DocumentBuilderDocx : DocumentBuilder
    {
        public DocumentBuilderDocx(string templateFileName, string xmlDataFileName, string outputFileName) :
            base(templateFileName, xmlDataFileName, outputFileName)
        { 
            _documentType = DOCX;

        }
    } // class DocumentBuilderDocx
} // namespace dblib