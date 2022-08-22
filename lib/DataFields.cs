namespace dblib
{
    using System.Globalization;
    using System.Xml;

    enum DataTypes : ushort
    {
        UNKNOWN = 0,
        TEXT = 1,
        NUMBER = 2,
        BOOLEAN = 3
    }

    class DataType
    {
        private DataTypes _dataType;

        public DataType(string TypeName)
        {
            string typename = TypeName.ToLower();
            if (typename == "text")
            {
                _dataType = DataTypes.TEXT;
            }
            else if (typename == "number")
            {
                _dataType = DataTypes.NUMBER;
            }
            else if (typename == "boolean")
            {
                _dataType = DataTypes.BOOLEAN;
            }
            else
            {
                _dataType = DataTypes.UNKNOWN;
            }
        }

        public DataTypes Get()
        {
            return _dataType;
        }
    }

    class DataField
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private string _xmlId;

        private string _xmlType;

        private string _tebenName;

        private string _valueAsString;

        private DataType _dataType;

        private DataField()
        {
        }

        public DataField(
            string xmlId,
            string xmlType,
            string tebenName,
            string valueAsString
        )
        {
            _xmlId = xmlId;
            _xmlType = xmlType;
            _tebenName = tebenName;
            _valueAsString = valueAsString;
            _dataType = new DataType(_xmlType);
        }

        public DataField(XmlNode xmlNode)
        {
            if (xmlNode.Name != "Field")
            {
                throw new ArgumentException("DataField accepts only <Field> nodes.");
            }

            // get the attributes
            XmlElement xmlElement = (XmlElement)xmlNode;
            XmlAttribute attrib = xmlElement.GetAttributeNode("id");
            _xmlId = attrib.InnerXml;
            attrib = xmlElement.GetAttributeNode("type");
            _xmlType = attrib.InnerXml;
            // attribute "tbn-name" is optional and we ignore all errors here
            try
            {
                attrib = xmlElement.GetAttributeNode("tbn-name");
                _tebenName = attrib.InnerXml;
            }
            catch (Exception ex)
            {
                _tebenName = "";
            }
            // get the value as string
            _valueAsString = xmlElement["Value"].InnerText;

            _dataType = new DataType(_xmlType);
        }

        public string GetXmlId()
        {
            return _xmlId;
        }

        public string GetTebenName()
        {
            return _tebenName;
        }

        public string GetValue()
        {
            return _valueAsString;
        }

        public DataTypes GetDataType()
        {
            return _dataType.Get();
        }

        public double GetDoubleValue()
        {
            // Gets a NumberFormatInfo associated with the en-US culture.
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            double parsed = Double.Parse(_valueAsString, nfi);
            Logger.Trace("string \"" + _valueAsString + "\" as double: " + parsed);
            return parsed;
        }

        public bool GetBooleanValue()
        {
            return System.Convert.ToBoolean(_valueAsString);
        }
    } // class DataField

    class DataFieldList
    {
        private List<DataField> _list;

        public DataFieldList()
        {
            _list = new List<DataField>();
        }

        public void Add(DataField dataField)
        {
            _list.Add(dataField);
        }

        public DataField GetById(string id)
        {
            return _list.Find(f => f.GetXmlId() == id);
        }

        public DataField GetByTebenName(string tebenname)
        {
            return _list.Find(f => f.GetTebenName() == tebenname);
        }
    }

} // namespace dblib