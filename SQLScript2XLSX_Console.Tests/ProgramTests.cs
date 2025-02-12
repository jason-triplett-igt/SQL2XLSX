using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace SQLScript2XLSX_Console.Tests
{
    [TestClass]
    public class ProgramTests
    {
        [TestMethod]
        public void Main_WithInvalidArguments_ShouldPrintUsage()
        {
            // Arrange
            var args = new string[0];
            var output = new StringWriter();
            Console.SetOut(output);

            // Act
            Program.Main(args);

            // Assert
            var expectedOutput = "Usage: SQLScript2XLSX_Console <connectionString> <sqlFilePath> <outputExcelPath>\r\n";
            Assert.AreEqual(expectedOutput, output.ToString());
        }

        [TestMethod]
        public void Main_WithValidArguments_ShouldExportDataToExcel()
        {
            // Arrange
            var sqlconnectionstringbuilder = new SqlConnectionStringBuilder();
            sqlconnectionstringbuilder.DataSource = "sql01";
            sqlconnectionstringbuilder.InitialCatalog = "master";
            sqlconnectionstringbuilder.TrustServerCertificate = true;
            sqlconnectionstringbuilder.IntegratedSecurity = true;
            var connectionString = sqlconnectionstringbuilder.ConnectionString;
            var sqlFilePath = "test.sql";
            var outputExcelPath = "output.xlsx";
            var args = new[] { connectionString, sqlFilePath, outputExcelPath };

            var sqlQuery = "SELECT 1 AS Column1; SELECT 2 AS Column2;";
            File.WriteAllText(sqlFilePath, sqlQuery);

            var mockConnection = new Mock<IDbConnection>();
            var mockCommand = new Mock<IDbCommand>();
            var mockReader = new Mock<IDataReader>();

            mockReader.SetupSequence(r => r.Read())
                .Returns(true)
                .Returns(false)
                .Returns(true)
                .Returns(false);

            mockReader.SetupSequence(r => r.NextResult())
                .Returns(true)
                .Returns(false);

            mockCommand.Setup(c => c.ExecuteReader()).Returns(mockReader.Object);
            mockConnection.Setup(c => c.CreateCommand()).Returns(mockCommand.Object);
            mockConnection.Setup(c => c.Open());

            // Act
            Program.Main(args);

            // Assert
            Assert.IsTrue(File.Exists(outputExcelPath));
            using (var workbook = new XLWorkbook(outputExcelPath))
            {
                Assert.AreEqual(2, workbook.Worksheets.Count);
                Assert.AreEqual("Results1", workbook.Worksheets.Worksheet(1).Name);
                Assert.AreEqual("Results2", workbook.Worksheets.Worksheet(2).Name);
            }

            // Cleanup
            File.Delete(sqlFilePath);
            File.Delete(outputExcelPath);
        }
    }
}
