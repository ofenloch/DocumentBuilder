namespace dblib
{
    class DocumentBuilderDocx : DocumentBuilder
    {
        public DocumentBuilderDocx(string templateFileName, string xmlDataFileName, string outputFileName) :
            base(templateFileName, xmlDataFileName, outputFileName)
        {
            _documentType = DOCX;

        }

        public override int ProcessTemplate(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            int error = 0;



            return error;
        }

    } // class DocumentBuilderDocx
} // namespace dblib