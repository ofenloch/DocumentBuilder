// This is a .NET 5 (and earlier) console app template
// See https://aka.ms/new-console-template for more information

using System;

namespace DocumentBuilder
{
    internal class Program
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            Console.WriteLine("DocumentBuilder.NET");

            int nArgs = args.Length;
            if (nArgs != 3)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("      DocumentBuilder <template file> <data file> <output file>");
                Console.WriteLine("examples");
                Console.WriteLine("      DocumentBuilder template.xlsx data.xml generated-document.xlsx");
                Console.WriteLine("      DocumentBuilder template.docx data.xml generated-document.docx");
                return;
            }
            string templateFileName = Path.GetFullPath(args[0]);
            string templateDirectory = Path.GetDirectoryName(templateFileName);
            string dataFileName = Path.GetFullPath(args[1]);
            string dataDirectory = Path.GetDirectoryName(dataFileName);
            string outputFileName = Path.GetFullPath(args[1]);
            string outputDirectory = Path.GetDirectoryName(outputFileName);

            outputDirectory = dblib.Utilities.CreateDirectory(outputDirectory);
            Logger.Info("output directory is {0}", outputDirectory);
            Logger.Info("data directory is {0}", dataDirectory);

            string templateWord = Path.Combine(dataDirectory, "template.docx");
            Logger.Info("Word template is {0}", templateWord);
            string templateExcel = Path.Combine(dataDirectory, "template.xlsx");
            Logger.Info("Excel template is {0}", templateExcel);

            string xmlDataFileName = Path.Combine(dataDirectory, "template.xml");


            string testDocx = Path.Combine(outputDirectory, "test.docx");
            string testXlsx = Path.Combine(outputDirectory, "test.xlsx");

            dblib.DocumentBuilder.CreateNewWordDocument(testDocx);
            dblib.DocumentBuilder.CreateNewExcelDocument(testXlsx);

            dblib.DocumentBuilder.UnpackPackage(testDocx, testDocx + "_unpacked");
            dblib.DocumentBuilder.UnpackPackage(testXlsx, testXlsx + "_unpacked");

            dblib.DocumentBuilderXlsx dbExcel = new dblib.DocumentBuilderXlsx(templateExcel, xmlDataFileName, Path.Combine(outputDirectory, "example.xlsx"));
            dblib.DocumentBuilderDocx dbWord = new dblib.DocumentBuilderDocx(templateWord, xmlDataFileName, Path.Combine(outputDirectory, "example.docx"));

        }
    } // internal class Program
} // namespace DocumentBuilder