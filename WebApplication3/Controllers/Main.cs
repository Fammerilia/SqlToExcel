using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;

namespace ExcelApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public ExcelController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpPost]
        [Route("execute")]
        public async Task<IActionResult> ExecuteSqlQuery([FromBody] SqlQueryModel queryModel)
        {
            try
            {
                var connectionString = _configuration.GetConnectionString("Default");

                var dataTable = new DataTable();

                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var command = new SqlCommand(queryModel.Query, connection))
                    {
                        using (var adapter = new SqlDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    int headerIndex = 1;
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        worksheet.Cells[1, headerIndex].Value = column.ColumnName;

                        headerIndex++;
                    }

                    int rowCount = 2; 

                    foreach (DataRow row in dataTable.Rows)
                    {
                        int colCount = 1;

                        foreach (var item in row.ItemArray)
                        {
                            if (item == DBNull.Value) 
                            {
                                worksheet.Cells[rowCount, colCount].Value = "null";
                            }
                            else
                            {
                                worksheet.Cells[rowCount, colCount].Value = item;
                            }
                            colCount++;
                        }

                        rowCount++;
                    }
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();


                    var fileName = $"result_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                    var filePath = Path.Combine(Path.GetTempPath(), fileName);

                    package.SaveAs(new FileInfo(filePath));

                    var stream = new FileStream(filePath, FileMode.Open);
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Ошибка при выполнении SQL-запроса: {ex.Message}");
            }
        }
    }
}
