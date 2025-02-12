using System;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 3)
        {
            Console.WriteLine("Usage: SQLScript2XLSX_Console <connectionString> <sqlFilePath> <outputExcelPath>");
            return;
        }

        string connectionString = args[0];
        string sqlFilePath = args[1];
        string outputExcelPath = args[2];

        try
        {
            string sqlQuery = File.ReadAllText(sqlFilePath);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(sqlQuery, connection);
                connection.Open();

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        int sheetIndex = 1;

                        do
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);
                            workbook.Worksheets.Add(dataTable, "Results" + sheetIndex);
                            sheetIndex++;
                        } while (!reader.IsClosed);

                        workbook.SaveAs(outputExcelPath);
                    }
                }
            }

            Console.WriteLine("Data exported successfully to " + outputExcelPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
