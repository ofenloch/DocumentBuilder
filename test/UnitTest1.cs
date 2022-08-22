using System.IO;
using Xunit;

namespace unit_tests;

public class UnitTest1
{
    private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
    static string _outDir = Path.GetFullPath("./output");

    [Fact]
    public void TestMethod()
    {
        string currentWorkingDirectory = Directory.GetCurrentDirectory();
        Logger.Debug("UnitTest1.TestMethod: running in working directory {0}", currentWorkingDirectory);
        Assert.False(true, $"Expected path {currentWorkingDirectory}");
    }

    [Fact]
    public void test_DocumentBuilder_Basic()
    {
        string testDocx = Path.Combine(_outDir, "test.docx");
        Logger.Debug("test_DocumentBuilder_Basic: testDocx is {0}", testDocx);
        string testXlsx = Path.Combine(_outDir, "test.xlsx");
        Logger.Debug("test_DocumentBuilder_Basic: testXlsx is {0}", testXlsx);

        dblib.DocumentBuilder.CreateNewWordDocument(testDocx);
        Assert.True(File.Exists(testDocx), $"file {testDocx} should exist");

        dblib.DocumentBuilder.CreateNewExcelDocument(testXlsx);
        Assert.True(File.Exists(testXlsx), $"file {testXlsx} should exist");
    }
}