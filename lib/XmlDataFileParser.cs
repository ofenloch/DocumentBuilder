namespace dblib
{
    using System;
    using System.Xml;
    using System.Xml.Schema;

    class XmlDataFileParser
    {
        private XmlDataFileParser()
        {
        }

        private DataFieldList _dataFields;

        private XmlDocument _xmlDocument;

        private string _fileName = "bookdtd.xml";

        private XmlReader _xmlReader;

        public XmlDataFileParser(string fileName)
        {
            ValidationEventHandler eventHandler =
                new ValidationEventHandler(XmlDataFileParser.ValidationCallback);

            try
            {
                _dataFields = new DataFieldList();
                _xmlDocument = new XmlDocument();
                _fileName = fileName;

                // Create the validating reader and specify DTD validation.
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Parse;
                settings.ValidationType = ValidationType.DTD;
                settings.ValidationEventHandler += eventHandler;
                _xmlReader = XmlReader.Create(_fileName, settings);
                _xmlDocument.Load(_xmlReader);

                //Console.WriteLine(_xmlDocument.OuterXml);
                XmlNodeList xmlFields = _xmlDocument.GetElementsByTagName("Field");

                Console.WriteLine(" found {0} <Field> elements", xmlFields.Count);
                foreach (XmlNode xmlField in xmlFields)
                {
                    DataField newDataField = new DataField(xmlField);
                    _dataFields.Add(newDataField);
                }
            }
            finally
            {
                if (_xmlReader != null)
                {
                    _xmlReader.Close();
                }
            }
        } // public XmlDocument(string fileName)

        //************************************************************************************
        //
        //  Event handler that is raised when XML doesn't validate against the schema.
        //
        //************************************************************************************
        private static void ValidationCallback(
            object sender,
            System.Xml.Schema.ValidationEventArgs e
        )
        {
            if (e.Severity == XmlSeverityType.Warning)
            {
                Console
                    .WriteLine("The following validation warning occurred: " +
                    e.Message);
            }
            else if (e.Severity == XmlSeverityType.Error)
            {
                Console
                    .WriteLine("The following critical validation errors occurred: " +
                    e.Message);
            }
        }

        public DataFieldList GetDataFields()
        {
            return _dataFields;
        }
    } // class XmlDataFileParser

} // namespace dblib