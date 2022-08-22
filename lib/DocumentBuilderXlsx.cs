namespace dblib
{
    public class DocumentBuilderXlsx : DocumentBuilder
    {
        public DocumentBuilderXlsx(string templateFileName, string xmlDataFileName, string outputFileName) :
        base(templateFileName, xmlDataFileName, outputFileName)
        {
            _documentType = XLSX;

        }

        public override int ProcessTemplate(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            int error = 0;



            return error;
        }

    } // class DocumentBuilderXlsx
} // namespace dblib