using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;

class Config
{
  public string PdfFolder { get; set; } = "";
  public string OutputOperazioni { get; set; } = "";
}

public class Operazione
{
  public DateTime Data { get; set; }
  public string Tipo { get; set; } = "";
  public string Descrizione { get; set; } = "";
  public decimal Importo { get; set; }
}

class Program
{
  static void Main()
  {
    var config = CaricaConfigurazione("appsettings.json");
    var operazioni = new List<Operazione>();
    var pdfMovimentiExtractor = new PdfMovimentiExtractor();
    foreach (var file in Directory.GetFiles(config.PdfFolder, "*.pdf"))
    {
      FileInfo fileInfo = new FileInfo(file);
      switch (fileInfo.Name.Replace(".pdf", "").ToLower())
      {
        case "tr":
        case "trade_repubblic":
        case "traderepubblic":
          operazioni.AddRange(pdfMovimentiExtractor.EstraiMovimentiFromTr(fileInfo.FullName));
          break;
        case "ing":
          operazioni.AddRange(pdfMovimentiExtractor.EstraiMovimentiFromIng(fileInfo.FullName));
          break;
      }
    }

    // CSV operazioni
    using (var sw = new StreamWriter(config.OutputOperazioni))
    {
      sw.WriteLine("Data;Tipo;Descrizione;Importo");
      foreach (var op in operazioni)
      {
        sw.WriteLine($"{op.Data:yyyy-MM-dd};\"{op.Tipo}\";\"{op.Descrizione}\";{op.Importo.ToString(CultureInfo.CurrentCulture)}");
      }
    }

    Console.WriteLine($"File CSV creato - {config.OutputOperazioni}");
    Console.ReadLine();
  }

  static Config CaricaConfigurazione(string path)
  {
    if (!File.Exists(path)) return new Config();
    var json = File.ReadAllText(path);
    return JsonSerializer.Deserialize<Config>(json) ?? new Config();
  }
}

public class PdfMovimentiExtractor
{
  public List<Operazione> EstraiMovimentiFromIng(string pdfPath)
  {
    // 1) estraggo testo
    string fullText;
    using (var reader = new PdfReader(pdfPath))
    using (var pdf = new PdfDocument(reader))
    {
      var sb = new System.Text.StringBuilder();
      for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
      {
        sb.Append(PdfTextExtractor.GetTextFromPage(pdf.GetPage(i)));
      }
      fullText = sb.ToString();
    }
    int start = fullText.IndexOf("(EUR)(EUR)") + 10;
    string movimentiString = fullText.Substring(start, fullText.Length - start);
    movimentiString = movimentiString.Replace("TipoTipo", "");
    movimentiString = movimentiString.Replace("ImporImportoto", "");
    movimentiString = movimentiString.Replace("DatData Ca Contontabileabile DescrizioneDescrizione", "");
    movimentiString = movimentiString.Replace("TTrransazioneansazione (EUR)(EUR)", "");
    int startIndex = 0;
    List<string> listaMovimentiString = new List<string>();
    while (true)
    {
      int indexSaldoDisponibile = movimentiString.IndexOf("Saldo Disponibile: ", startIndex);
      if (indexSaldoDisponibile == -1)
      {
        break;
      }
      int indexEurDopoSaldoDisponibile = movimentiString.IndexOf("EUR", indexSaldoDisponibile) + 3;
      string movimento = movimentiString.Substring(startIndex, indexEurDopoSaldoDisponibile - startIndex);
      listaMovimentiString.Add(movimento);
      startIndex = indexEurDopoSaldoDisponibile;
    }
    var results = new List<Operazione>();
    for (int i = 0; i < listaMovimentiString.Count; i++)
    {
      List<string> movList = listaMovimentiString[i].Split("\n", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
      string dataGiornoMeseImportoTipo = movList[1];
      List<string> dataGiornoMeseImportoTipoList = dataGiornoMeseImportoTipo.Split(" ", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
      movList.RemoveAt(1);
      string annoTipoTransazione = movList[2];
      List<string> annoTipoTransazioneList = annoTipoTransazione.Split(" ", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
      movList.RemoveAt(2);
      string descrizione = "";
      foreach (var item in movList)
      {
        descrizione += item;
      }
      int giorno = int.Parse(dataGiornoMeseImportoTipoList[0]);
      int mese = Array.IndexOf(Months, dataGiornoMeseImportoTipoList[1]) + 1;
      int anno = int.Parse(annoTipoTransazioneList[0]);
      decimal importo = decimal.Parse(dataGiornoMeseImportoTipoList[3]);
      string tipoTransazione = dataGiornoMeseImportoTipoList[2];
      for (int j = 1; j < annoTipoTransazioneList.Count; j++)
      {
        tipoTransazione += $" {annoTipoTransazioneList[j]}";
      }
      descrizione = descrizione.Replace("Saldo Disponibile", " Saldo Disponibile");
      results.Add(new Operazione
      {
        Data = new DateTime(anno, mese, giorno),
        Descrizione = descrizione,
        Tipo = tipoTransazione,
        Importo = importo
      });
    }

    return results;
  }

  public List<Operazione> EstraiMovimentiFromTr(string pdfPath)
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
              string rigaSenzaData = riga.Replace(match.Groups[1].Value, "").Trim();
              var a = rigaSenzaData.Replace(" €", "€").Split("€", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
              rigaSenzaData = a[0];
              var match1 = Regex.Match(rigaSenzaData, @"\d{1,3}(?:\.\d{3})*,\d{2}");
              string valoreStr = "";
              string tipoDescrizione = "";
              if (match1.Success)
              {
                valoreStr = match1.Value;
                tipoDescrizione = rigaSenzaData.Replace(match1.Value, "").Trim();
              }

              string tipo = "";
              string descrizione = "";
              decimal entrata = 0;
              decimal uscita = 0;
              if (tipoDescrizione.StartsWith("Bonifico"))
              {
                tipo = "Bonifico";
                descrizione = tipoDescrizione.Replace(tipo, "").Trim();
                if (tipoDescrizione.ToLower().Contains("incoming"))
                {
                  entrata = ParseDecimal(valoreStr);
                }
                else
                {
                  uscita = ParseDecimal(valoreStr);
                }

              }
              else if (tipoDescrizione.StartsWith("Transazione con carta"))
              {
                tipo = "Transazione con carta";
                descrizione = tipoDescrizione.Replace(tipo, "").Trim();
                uscita = ParseDecimal(valoreStr);
              }
              else if (tipoDescrizione.StartsWith("Trasferimento"))
              {
                tipo = "Trasferimento";
                descrizione = tipoDescrizione.Replace(tipo, "").Trim();
                entrata = ParseDecimal(valoreStr);
              }

              if (entrata == 0 && uscita > 0)
              {
                operazioni.Add(new Operazione
                {
                  Data = data,
                  Tipo = tipo,
                  Descrizione = descrizione,
                  Importo = uscita,
                });
              }
            }
          }
        }
      }
    }

    return operazioni;
  }

  private decimal ParseDecimal(string s)
  {
    if (string.IsNullOrWhiteSpace(s)) return 0;
    s = s.Replace(".", "").Replace(",", ".");
    return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;
  }

  // mesi italiani
  private readonly string[] Months = new[]
  {
        "gennaio","febbraio","marzo","aprile","maggio","giugno",
        "luglio","agosto","settembre","ottobre","novembre","dicembre"
    };
}