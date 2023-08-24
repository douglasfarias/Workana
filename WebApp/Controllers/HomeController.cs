using System.Data;
using System.Diagnostics;

using Microsoft.AspNetCore.Mvc;

using Newtonsoft.Json;

using OfficeOpenXml;

using WebApp.Models;

namespace WebApp.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        ViewBag.Indice = true;
        return View();
    }

    [Route("XlsToJson")]
    public IActionResult XlsToJson()
    {
        ViewBag.Action = "XlsToJson";
        return View("Index");
    }

    [HttpPost]
    [Route("XlsToJson")]
    public IActionResult XlsToJson([FromForm] IFormFile file)
    {
        if(Path.GetExtension(file.FileName).ToLower() != ".xlsx")
        {
            ViewBag.Erro = "Somente arquivos .xlsx são permitidos.";
            return View("Index");
        }
        
        try
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var stream = file.OpenReadStream();
            using var package = new ExcelPackage(stream);
            var data = package.Workbook.Worksheets[0].Cells["B1:CF2"].ToDataTable();
            var items = data.Rows.OfType<DataRow>().Select(row => data.Columns.OfType<DataColumn>().ToDictionary(col => col.ColumnName, c => row[c]?.ToString()));
            ViewBag.Resultado = JsonConvert.SerializeObject(items, Formatting.Indented);
            return View("Index");
        }
        catch
        {

            ViewBag.Erro = "Houve um erro inesperado, contate o administrador do sistema";
            return View("Index");
        }
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
