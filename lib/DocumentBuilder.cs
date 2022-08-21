namespace dblib
{
    public class DocumentBuilder
    {

        private string _outputDirectory = "";
        private string _outputFileName = "";
        private string _outputFileBaseName = "";
        public DocumentBuilder(string templateFileName, string xmlDataFileName, string outputFileName)
        {
            _outputFileName = Path.GetFullPath(outputFileName);
            _outputFileBaseName = Path.GetFileNameWithoutExtension(_outputFileName);
            _outputDirectory = Path.GetDirectoryName(_outputFileName);
            Utilities.CreateDirectory(_outputDirectory);
        }

    } // public class DocumentBuilder

} // namespace dblib
