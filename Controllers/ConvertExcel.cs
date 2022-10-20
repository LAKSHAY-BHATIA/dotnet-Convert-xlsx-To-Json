using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
namespace XLS_To_JSON_dotnet.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConvertExcel : ControllerBase
    {
        [HttpPost]
        [Route("ToJson")]
        public IActionResult ConvertToJson(IFormFile file)
        {
            try
            {
                string JSONString = String.Empty;
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using(var reader = ExcelReaderFactory.CreateReader(file.OpenReadStream()))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration() { 
                    ConfigureDataTable = (data)=> new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true,
                    }
                    });
                    JSONString = JsonConvert.SerializeObject(result.Tables);
                    return Ok(JSONString);
                }
            }
            catch(Exception Ex)
            {
                return BadRequest(new { success = false, Message = Ex.Message });
            }
        }
    }
}
