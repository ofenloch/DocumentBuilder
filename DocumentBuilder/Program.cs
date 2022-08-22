// This is a .NET 5 (and earlier) console app template
// See https://aka.ms/new-console-template for more information

using System;

namespace DocumentBuilder
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            string outDir = dblib.Utilities.CreateDirectory("./output/");
            Console.WriteLine("output directory is {0}", outDir);

            string dataDir = dblib.Utilities.CreateDirectory("./data/");
            Console.WriteLine("data directory is {0}", dataDir);

            string templateWord = Path.Combine(dataDir, "template.docx");
            Console.WriteLine("Word template is {0}", templateWord);
            string templateExcel = Path.Combine(dataDir, "template.xlsx");
            Console.WriteLine("Excel template is {0}", templateExcel);

            string xmlDataFileName = Path.Combine(dataDir, "template.xml");


            string testDocx = Path.Combine(outDir, "test.docx");
            string testXlsx = Path.Combine(outDir, "test.xlsx");

            dblib.DocumentBuilder.CreateNewWordDocument(testDocx);
            dblib.DocumentBuilder.CreateNewExcelDocument(testXlsx);

            dblib.DocumentBuilder.UnpackPackage(testDocx, testDocx + "_unpacked");
            dblib.DocumentBuilder.UnpackPackage(testXlsx, testXlsx + "_unpacked");

            dblib.DocumentBuilderXlsx dbExcel = new dblib.DocumentBuilderXlsx(templateExcel, xmlDataFileName, Path.Combine(outDir, "example.xlsx"));
            dblib.DocumentBuilderDocx dbWord = new dblib.DocumentBuilderDocx(templateWord, xmlDataFileName, Path.Combine(outDir, "example.docx"));

        }
    } // internal class Program
} // namespace DocumentBuilder