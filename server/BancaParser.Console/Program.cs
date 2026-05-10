using BancaParser.Core;
using Microsoft.Extensions.Configuration;

class Config
{
  public string folderFiles { get; set; } = "";
  public string OutputOperazioni { get; set; } = "";
  public Dictionary<string, string> Descrizioni { get; set; } = new Dictionary<string, string>();
}

class Program
{
  static void Main()
  {
    // Costruisco la configuration leggendo appsettings.json
    var configuration = new ConfigurationBuilder()
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .Build();
    var appConfig = configuration.Get<Config>();
    // Binding diretto su un Dictionary
    var mySettings = configuration.GetSection("Descrizione").Get<Dictionary<string, string>>();
    var pdfMovimentiExtractor = new PdfMovimentiExtractor();
    var files = Directory.GetFiles(appConfig.folderFiles, "*.pdf").ToList();
    files.AddRange(Directory.GetFiles(appConfig.folderFiles, "*.xlsx"));
    files.AddRange(Directory.GetFiles(appConfig.folderFiles, "*.csv"));

    RecuperaOperazioniResult recuperaOperazioniResult = pdfMovimentiExtractor.RecuperaOperazioni(files);
    List<Operazione> operazioniDefinitive = pdfMovimentiExtractor.LavoraOperazioni(appConfig.Descrizioni, recuperaOperazioniResult.Operazioni);
    pdfMovimentiExtractor.ExportToCsv(appConfig.OutputOperazioni, operazioniDefinitive);
    Console.WriteLine($"File CSV creato - {appConfig.OutputOperazioni}");
    if (recuperaOperazioniResult.Errori.Any())
    {
      Console.WriteLine("Alcuni file non sono stati elaborati:");
      foreach (var errore in recuperaOperazioniResult.Errori)
      {
        Console.WriteLine($"- {errore.File}: {errore.Errore}");
      }
    }
    Console.ReadLine();
  }
}
