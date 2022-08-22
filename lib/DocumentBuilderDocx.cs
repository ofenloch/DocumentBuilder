namespace dblib
{
    class DocumentBuilderDocx : DocumentBuilder
    {
        public DocumentBuilderDocx(string templateFileName, string xmlDataFileName, string outputFileName) :
            base(templateFileName, xmlDataFileName, outputFileName)
        { 
            _documentType = DOCX;

        }

        override public int ProcessTemplate(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            
        }

    } // class DocumentBuilderDocx
} // namespace dblib