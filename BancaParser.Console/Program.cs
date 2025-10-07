using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

class Config
{
  public string folderFiles { get; set; } = "";
  public string OutputOperazioni { get; set; } = "";
  public Dictionary<string, string> Descrizioni { get; set; } = new Dictionary<string, string>();
}

public class Operazione
{
  public DateTime Data { get; set; }
  public string Tipo { get; set; } = "";
  public string Descrizione { get; set; } = "";
  public decimal Importo { get; set; }
  public decimal ImportoRossella { get; set; }
  public decimal ImportoLuca { get; set; }
  public bool IsContabilizzato { get; set; }

  public override string ToString()
  {
    string s = "";
    s += $"{this.Data:yyyy-MM-dd};";
    s += $"\"{this.Tipo}\";";
    s += $"\"{this.Descrizione}\";";
    s += $"{this.ImportoRossella.ToString(CultureInfo.CurrentCulture)};";
    s += $"{this.ImportoLuca.ToString(CultureInfo.CurrentCulture)};";
    s += $"{this.Importo.ToString(CultureInfo.CurrentCulture)};";
    s += $"{this.IsContabilizzato:\"S\":\"N\"}";
    return s;
  }
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
    var operazioni = new List<Operazione>();
    var pdfMovimentiExtractor = new PdfMovimentiExtractor();
    var files = Directory.GetFiles(appConfig.folderFiles, "*.pdf").ToList();
    files.AddRange(Directory.GetFiles(appConfig.folderFiles, "*.xlsx"));

    foreach (var file in files)
    {
      FileInfo fileInfo = new FileInfo(file);
      switch (fileInfo.Name.Replace(fileInfo.Extension, "").ToLower())
      {
        case string s when s.StartsWith("tr"):
        case string s1 when s1.StartsWith("trade_repubblic"):
        case string s2 when s2.StartsWith("traderepubblic"):
          operazioni.AddRange(pdfMovimentiExtractor.EstraiMovimentiFromTr(fileInfo.FullName));
          break;
        case string s when s.StartsWith("ing"):
        case string s1 when s1.StartsWith("ing direct"):
        case string s2 when s2.StartsWith("ingdirect"):
        case string s3 when s3.StartsWith("ing_direct"):
          operazioni.AddRange(pdfMovimentiExtractor.EstraiMovimentiFromIng(fileInfo.FullName));
          break;
        case string s when s.StartsWith("bpm"):
          operazioni.AddRange(pdfMovimentiExtractor.EstraiMovimentiFromBpm(fileInfo.FullName));
          break;
      }
    }
    var operazioniDefinitive = new List<Operazione>();
    foreach (var op in operazioni)
    {
      if (op.Importo > 0)
      {
        continue;
      }
      Operazione newOperazione = new Operazione();
      newOperazione.Data = op.Data;
      newOperazione.Tipo = op.Tipo;
      newOperazione.Importo = Math.Abs(op.Importo);
      newOperazione.ImportoRossella = newOperazione.Importo / 2 * -1;
      newOperazione.ImportoLuca = newOperazione.Importo / 2;
      newOperazione.Descrizione = op.Descrizione;
      foreach (var item in appConfig.Descrizioni)
      {
        if (op.Descrizione.ToLower().Contains(item.Key))
        {
          newOperazione.Descrizione = item.Value;
          break;
        }
      }
      operazioniDefinitive.Add(newOperazione);
    }

    // CSV operazioni
    using (var sw = new StreamWriter(appConfig.OutputOperazioni))
    {
      sw.WriteLine("Data;Tipo;Descrizione;Rossella;Luca;Importo;Contabilizzato");
      foreach (var op in operazioniDefinitive)
      {
        sw.WriteLine($"{op.ToString()}");
      }
    }

    Console.WriteLine($"File CSV creato - {appConfig.OutputOperazioni}");
    Console.ReadLine();
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
        Importo = importo,
        IsContabilizzato = false
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
                  IsContabilizzato = false
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
    s = s.Replace(".", "").Replace(",", ".").Replace("\"","");
    decimal dd = decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;
    return dd;
  }

  public List<Operazione> EstraiMovimentiFromBpm(string fullName)
  {
    using (var workbook = new XLWorkbook(fullName))
    {
      var worksheet = workbook.Worksheet(1);
      List<List<string>> list = new List<List<string>>();
      foreach (var row in worksheet.RowsUsed())
      {
        var values = row.Cells().Select(c =>
        {
          string v = c.GetValue<string>();
          if (v.Contains(",") || v.Contains("\""))
            v = $"\"{v.Replace("\"", "\"\"")}\"";
          return v;
        }).ToList();
        list.Add(values);
      }
      var results = new List<Operazione>();
      for (int i = 1; i < list.Count; i++)
      {
        results.Add(new Operazione
        {
          Data = Convert.ToDateTime(list[i][0]),
          Descrizione = list[i][4],
          Tipo = "",
          Importo = ParseDecimal(list[i][2]),
          IsContabilizzato = true
        });

      }
      return results;
    }
  }

  // mesi italiani
  private readonly string[] Months = new[]
  {
        "gennaio","febbraio","marzo","aprile","maggio","giugno",
        "luglio","agosto","settembre","ottobre","novembre","dicembre"
    };
}