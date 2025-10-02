using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

class Config
{
  public string PdfFolder { get; set; } = "estratti_pdf";
  public string OutputOperazioni { get; set; } = "operazioni.csv";
  public string OutputRiepilogo { get; set; } = "riepilogo_mensile.csv";
  public bool Verbose { get; set; } = true;
}

class Operazione
{
  public DateTime Data { get; set; }
  public string Tipo { get; set; } = "";
  public string Descrizione { get; set; } = "";
  public decimal Entrata { get; set; }
  public decimal Uscita { get; set; }
  public decimal Saldo { get; set; }
  public string Mese => Data.ToString("yyyy-MM");
}

class Program
{
  static void Main()
  {
    var config = CaricaConfigurazione("appsettings.json");
    var operazioni = new List<Operazione>();

    foreach (var file in Directory.GetFiles(config.PdfFolder, "*.pdf"))
    {
      if (config.Verbose) Console.WriteLine($"Parsing {file} ...");
      operazioni.AddRange(EstrattiDaPdf(file, config));
    }

    if (operazioni.Count == 0)
    {
      Console.WriteLine("⚠️ Nessuna operazione trovata.");
      return;
    }

    // CSV operazioni
    using (var sw = new StreamWriter(config.OutputOperazioni))
    {
      sw.WriteLine("Data,Tipo,Descrizione,InEntrata,InUscita,Saldo");
      foreach (var op in operazioni)
      {
        sw.WriteLine($"{op.Data:yyyy-MM-dd},\"{op.Tipo}\",\"{op.Descrizione}\",{op.Entrata.ToString(CultureInfo.InvariantCulture)},{op.Uscita.ToString(CultureInfo.InvariantCulture)},{op.Saldo.ToString(CultureInfo.InvariantCulture)}");
      }
    }

    // CSV riepilogo mensile
    var riepilogo = operazioni
        .GroupBy(o => o.Mese)
        .Select(g => new
        {
          Mese = g.Key,
          TotaleEntrate = g.Sum(x => x.Entrata),
          TotaleUscite = g.Sum(x => x.Uscita)
        })
        .OrderBy(x => x.Mese)
        .ToList();

    using (var sw = new StreamWriter(config.OutputRiepilogo))
    {
      sw.WriteLine("Mese,TotaleEntrate,TotaleUscite");
      foreach (var r in riepilogo)
      {
        sw.WriteLine($"{r.Mese},{r.TotaleEntrate.ToString(CultureInfo.InvariantCulture)},{r.TotaleUscite.ToString(CultureInfo.InvariantCulture)}");
      }
    }

    Console.WriteLine($"✅ File CSV creati:\n - {config.OutputOperazioni}\n - {config.OutputRiepilogo}");
  }

  static Config CaricaConfigurazione(string path)
  {
    if (!File.Exists(path)) return new Config();
    var json = File.ReadAllText(path);
    return JsonSerializer.Deserialize<Config>(json) ?? new Config();
  }

  static List<Operazione> EstrattiDaPdf(string pdfPath, Config config)
  {
    var operazioni = new List<Operazione>();

    using (var reader = new PdfReader(pdfPath))
    using (var pdfDoc = new PdfDocument(reader))
    {
      for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
      {
        var text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i));
        var righe = text.Split('\n')
                        .Select(r => r.Trim())
                        .Where(r => r.Length > 0)
                        .ToList();

        foreach (var riga in righe)
        {
          if (Regex.IsMatch(riga, @"\d{1,2}\s\w{3}\s\d{4}", RegexOptions.IgnoreCase))
          {
            var match = Regex.Match(riga, @"^(\d{1,2}\s\w{3}\s\d{4})\s+([A-Za-zÀ-ÿ ]+)\s+(.*?)\s+([\d\.,]*)\s*€?\s*([\d\.,]*)\s*€?\s*([\d\.,]*)\s*€?$");

            if (match.Success)
            {
              var data = DateTime.ParseExact(match.Groups[1].Value, "d MMM yyyy", new CultureInfo("it-IT"));
              var tipo = match.Groups[2].Value.Trim();
              var descrizione = match.Groups[3].Value.Trim();

              decimal entrata = ParseDecimal(match.Groups[4].Value);
              decimal uscita = ParseDecimal(match.Groups[5].Value);
              decimal saldo = ParseDecimal(match.Groups[6].Value);

              operazioni.Add(new Operazione
              {
                Data = data,
                Tipo = tipo,
                Descrizione = descrizione,
                Entrata = entrata,
                Uscita = uscita,
                Saldo = saldo
              });

              if (config.Verbose)
                Console.WriteLine($"Trovata op: {data:yyyy-MM-dd} {tipo} {descrizione} +{entrata} -{uscita} saldo={saldo}");
            }
          }
        }
      }
    }

    return operazioni;
  }

  static decimal ParseDecimal(string s)
  {
    if (string.IsNullOrWhiteSpace(s)) return 0;
    s = s.Replace(".", "").Replace(",", ".");
    return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;
  }
}
