using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Text.Json;

namespace YourERP.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CostManagementController : ControllerBase
    {
        private readonly IConfiguration _config;

        public CostManagementController(IConfiguration config)
        {
            _config = config;
        }

        // GET api/CostManagement/GetAll
        [HttpGet("GetAll")]
        public async Task<IActionResult> GetAll()
        {
            var result = new
            {
                priceList  = new List<Dictionary<string, object>>(),
                rawCost    = new List<Dictionary<string, object>>(),
                bom        = new List<Dictionary<string, object>>()
            };

            var priceList = new List<Dictionary<string, object>>();
            var rawCost   = new List<Dictionary<string, object>>();
            var bom       = new List<Dictionary<string, object>>();

            try
            {
                var connStr = _config.GetConnectionString("ERPManagment");

                using var conn = new SqlConnection(connStr);
                using var cmd  = new SqlCommand("[INV].[APIGetCostManagementData]", conn)
                {
                    CommandType    = CommandType.StoredProcedure,
                    CommandTimeout = 60
                };

                cmd.Parameters.AddWithValue("@Operation",         "GetAll");
                cmd.Parameters.AddWithValue("@User",              "API");
                cmd.Parameters.AddWithValue("@PlatForm",          "Web");
                cmd.Parameters.AddWithValue("@AppVersionWeb",     "1.0");
                cmd.Parameters.AddWithValue("@AppVersionAndroid", DBNull.Value);
                cmd.Parameters.AddWithValue("@AppVersionIos",     DBNull.Value);
                cmd.Parameters.AddWithValue("@AppVersionDesktop", DBNull.Value);
                cmd.Parameters.AddWithValue("@FireBaseToken",     DBNull.Value);
                cmd.Parameters.AddWithValue("@SqlStatement",      DBNull.Value);

                var stateParam = cmd.Parameters.Add("@State",   SqlDbType.Int);
                stateParam.Direction = ParameterDirection.Output;

                var msgParam = cmd.Parameters.Add("@Message", SqlDbType.NVarChar, 500);
                msgParam.Direction = ParameterDirection.Output;

                await conn.OpenAsync();

                using var reader = await cmd.ExecuteReaderAsync();

                // Result Set 1 — Price List
                while (await reader.ReadAsync())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null! : reader.GetValue(i);
                    priceList.Add(row);
                }

                // Result Set 2 — Raw Cost
                await reader.NextResultAsync();
                while (await reader.ReadAsync())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null! : reader.GetValue(i);
                    rawCost.Add(row);
                }

                // Result Set 3 — BOM
                await reader.NextResultAsync();
                while (await reader.ReadAsync())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null! : reader.GetValue(i);
                    bom.Add(row);
                }

                return Ok(new { priceList, rawCost, bom });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }
    }
}
