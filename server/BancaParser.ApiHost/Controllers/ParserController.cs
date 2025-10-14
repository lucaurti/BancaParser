using BancaParser.Core;
using BancaParser.Host.Helpers;
using Microsoft.AspNetCore.Mvc;

namespace BancaParser.Host.Controllers
{
  [Route("api/[controller]")]
  [ApiController]
  [Microsoft.AspNetCore.Authorization.Authorize]
  public class ParserController : ControllerBase
  {
    private readonly Config _configuration;
    public ParserController()
    {
      _configuration = new Config();
    }

    [HttpPost]
    public IActionResult GeneraResoconto() 
    {
      try
      {
        var pdfMovimentiExtractor = new PdfMovimentiExtractor();
        string basePath = Path.Combine(Utility.PATH_SHARED_FOLDER, "dati", "banca-parser");
        var files = Directory.GetFiles(Path.Combine(basePath, _configuration.RelativeFolderInputFiles), "*.pdf").ToList();
        files.AddRange(Directory.GetFiles(Path.Combine(basePath, _configuration.RelativeFolderInputFiles), "*.xlsx"));
        files.AddRange(Directory.GetFiles(Path.Combine(basePath, _configuration.RelativeFolderInputFiles), "*.csv"));

        List<Operazione> operazioni = pdfMovimentiExtractor.RecuperaOperazioni(files);
        List<Operazione> operazioniDefinitive = pdfMovimentiExtractor.LavoraOperazioni(_configuration.Descrizioni, operazioni);
        string pathCsv = $"{basePath}{_configuration.RelativeFolderOutputFiles}";
        pdfMovimentiExtractor.ExportToCsv(pathCsv, operazioniDefinitive);
        return Ok();
      }
      catch (Exception)
      {
        return BadRequest();
      }
    }
  }
}
