
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using SQLScript2XLSX.ViewModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SQLScript2XLSX.Models
{
    public static class ExportDataToXLSXfromSQL
    {
        public static async Task<Unit> ExportAsync(MainWindowViewModel vm, CancellationToken cancellationToken = new CancellationToken())
        {
            await Task.Run(async () =>
            {
                var constrbuilder = new SqlConnectionStringBuilder();
                var conn = new SqlConnection();
                constrbuilder.ConnectionString = $"Data Source={vm.Datasource};Initial Catalog={vm.InitialCatalog};Integrated Security={vm.IntegratedSecurity};";
                if (!vm.IntegratedSecurity)
                {
                    constrbuilder.UserID = vm.Username;
                    constrbuilder.Password = vm.Password;
                }
                constrbuilder.TrustServerCertificate = true;
                constrbuilder.ApplicationName = "ExportSQLScriptToXLSX";
                conn.ConnectionString = constrbuilder.ConnectionString;
                await conn.OpenAsync(cancellationToken);
                using var cmd = new SqlCommand(vm.Script, conn);
                var reader = await cmd.ExecuteReaderAsync(cancellationToken);
                //get datatype of each column
                var outputdatatable = new DataTable();
                outputdatatable.TableName = "QueryResults";
                while (!cancellationToken.IsCancellationRequested && await reader.ReadAsync(cancellationToken))
                {
                    if (outputdatatable is not null)
                    {
                        if (outputdatatable.Columns.Count == 0)
                        {
                            foreach (var datacol in await reader.GetColumnSchemaAsync(cancellationToken))
                            {
                                outputdatatable.Columns.Add(datacol.ColumnName, datacol.DataType ?? typeof(string));
                            }
                        }

                        var newdatarow = outputdatatable.NewRow();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {

                            newdatarow[i] = reader.GetValue(i);
                        }
                        outputdatatable.Rows.Add(newdatarow);

                    }


                }
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(outputdatatable);
                ws.Columns(1, ws.ColumnCount()).AdjustToContents();
                var wsquery = wb.Worksheets.Add("Source Query");
                wsquery.Cell(1, 1).Value = vm.Script;
                wb.SaveAs(vm.OutputPath);
                reader.Close();
                conn.Close();
            });
            return Unit.Default;
        }
    }
}
