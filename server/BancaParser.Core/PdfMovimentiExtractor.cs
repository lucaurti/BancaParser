using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Globalization;
using System.Text.RegularExpressions;

namespace BancaParser.Core
{
  public class PdfMovimentiExtractor
  {
    private const string SPLITWISE = "Splitwise";
    private const string HYPE = "Hype";
    private const string BPM = "BPM";
    private const string TRADEREPUBBLIC = "Trade Republic";
    private const string ING = "ING";

    public List<Operazione> RecuperaOperazioni(List<string> files)
    {
      List<Operazione> operazioni = new List<Operazione>();
      foreach (var file in files)
      {
        FileInfo fileInfo = new FileInfo(file);
        switch (fileInfo.Name.Replace(fileInfo.Extension, "").ToLower())
        {
          case string s when s.StartsWith("tr"):
          case string s1 when s1.StartsWith("trade_repubblic"):
          case string s2 when s2.StartsWith("traderepubblic"):
            operazioni.AddRange(EstraiMovimentiFromTr(fileInfo.FullName));
            break;
          case string s when s.StartsWith("ing"):
          case string s1 when s1.StartsWith("ing direct"):
          case string s2 when s2.StartsWith("ingdirect"):
          case string s3 when s3.StartsWith("ing_direct"):
            operazioni.AddRange(EstraiMovimentiFromIng(fileInfo.FullName));
            break;
          case string s when s.StartsWith("bpm"):
            operazioni.AddRange(EstraiMovimentiFromBpm(fileInfo.FullName));
            break;
          case string s when s.StartsWith("hype"):
            operazioni.AddRange(EstraiMovimentiFromHype(fileInfo.FullName));
            break;
          case string s when s.StartsWith("splitwise"):
            operazioni.AddRange(EstraiMovimentiFromSplitwise(fileInfo.FullName));
            break;
        }
      }
      return operazioni;
    }

    public void ExportToCsv(string outputOperazioni, List<Operazione> operazioniDefinitive)
    {
      // CSV operazioni
      using (var sw = new StreamWriter(outputOperazioni))
      {
        sw.WriteLine("Data;Tipo;Descrizione;Rossella;Luca;Importo;Contabilizzato;Banca");
        foreach (var op in operazioniDefinitive)
        {
          sw.WriteLine($"{op.ToString()}");
        }
      }
    }

    public List<Operazione> LavoraOperazioni(Dictionary<string, string> descrizioniInConfig, List<Operazione> operazioni)
    {
      var operazioniDefinitive = new List<Operazione>();
      foreach (var op in operazioni)
      {
        Operazione newOperazione = new Operazione();
        newOperazione.Data = op.Data;
        newOperazione.Tipo = op.Tipo;
        newOperazione.Importo = Math.Abs(op.Importo);
        newOperazione.Banca = op.Banca;
        if (newOperazione.Banca == SPLITWISE)
        {
          newOperazione.ImportoRossella = op.ImportoRossella;
          newOperazione.ImportoLuca = op.ImportoLuca;
        }
        else 
        {
          newOperazione.ImportoRossella = newOperazione.Importo / 2 * -1;
          newOperazione.ImportoLuca = newOperazione.Importo / 2;
        }
        
        if (op.IsContabilizzato)
        {
          newOperazione.ImportoRossella = 0;
          newOperazione.ImportoLuca = 0;
        }
        newOperazione.Descrizione = CapitalizeWords(op.Descrizione);
        foreach (var item in descrizioniInConfig)
        {
          if (op.Descrizione.ToLower().Contains(item.Key.ToLower()))
          {
            newOperazione.Descrizione = CapitalizeWords(item.Value);
            break;
          }
        }
        operazioniDefinitive.Add(newOperazione);
      }

      return operazioniDefinitive;
    }

    public static string CapitalizeWords(string input)
    {
      if (string.IsNullOrWhiteSpace(input))
        return input;

      var words = input
          .ToLowerInvariant()
          .Split(' ', StringSplitOptions.RemoveEmptyEntries);

      for (int i = 0; i < words.Length; i++)
      {
        words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
      }

      return string.Join(' ', words);
    }

    private List<Operazione> EstraiMovimentiFromIng(string pdfPath)
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
        decimal importo = ParseDecimal(dataGiornoMeseImportoTipoList[3]);
        string tipoTransazione = dataGiornoMeseImportoTipoList[2];
        for (int j = 1; j < annoTipoTransazioneList.Count; j++)
        {
          tipoTransazione += $" {annoTipoTransazioneList[j]}";
        }
        descrizione = descrizione.Replace("Saldo Disponibile", " Saldo Disponibile");
        if (importo < 0)
        {
          results.Add(new Operazione
          {
            Data = new DateTime(anno, mese, giorno),
            Descrizione = descrizione,
            Tipo = tipoTransazione,
            Importo = importo,
            IsContabilizzato = false,
            Banca = ING
          });
        }
      }

      return results;
    }

    private List<Operazione> EstraiMovimentiFromTr(string pdfPath)
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
          int rowheader = righe.IndexOf("DATA TIPO DESCRIZIONE IN ENTRATA IN USCITA SALDO");
          string firstRowDate = righe[rowheader + 1];
          string secondRowDate = righe[rowheader + 2];

          if (Regex.IsMatch(firstRowDate, @"\d{1,2}\s\w{3}\s\d{4}", RegexOptions.IgnoreCase) && Regex.IsMatch(secondRowDate, @"\d{1,2}\s\w{3}\s\d{4}", RegexOptions.IgnoreCase))
          {
            //ci troviamo nel caso in cui un singolo resoconto è presente nella stessa riga
            foreach (var riga in righe)
            {
              if (Regex.IsMatch(riga, @"^\d{1,2}\s\w{3}\s\d{4}", RegexOptions.IgnoreCase))
              {
                Operazione operazione = getOperazioneTr(riga);
                if (operazione == null)
                {
                }
                else
                {
                  operazioni.Add(operazione);
                }
              }
            }
          }
          else
          {
            //ci troviamo nel caso in cui un singolo resoconto è splittato su più righe
            List<string> lista = new List<string>();
            List<string> listaTipologie = new List<string>() { "Pagamento degli", "Transazione con", "Premio", "Commercio", "Bonifico" };
            for (int j = rowheader + 1; j < righe.Count; j = j + 3)
            {
              string str = "";
              var match = Regex.Match(righe[j], @"^\d{2}\s\w{3}");
              string s = righe[j];
              if (match.Success)
              {
                string giornoMese = match.Groups[0].Value.Trim();
                if (giornoMese == righe[j])
                {
                  // ci troviamo nel caso in cui nella riga c'è solo la data
                  string anno = righe[j + 2].Trim();
                  string tipoParte = righe[j + 1].Split(" ").FirstOrDefault();
                  string descrizione = righe[j + 1].Replace(tipoParte, "");
                  str += $"{giornoMese.Trim()} {anno.Trim()} {tipoParte.Trim()} {descrizione.Trim()}";
                }
                else if (!listaTipologie.Any(t => righe[j].Contains(t, StringComparison.OrdinalIgnoreCase)))
                {
                  string descrizioneParziale1 = righe[j].Replace(giornoMese, "").Trim();
                  string descrizioneParziale2 = "";

                  Regex r = new Regex(@"\d");
                  Match m = r.Match(righe[j + 1]);
                  string tipoParte = "";
                  if (m.Success)
                  {
                    tipoParte = righe[j + 1].Substring(0, m.Index).Trim();
                    descrizioneParziale2 = righe[j + 1].Substring(m.Index).Trim();
                  }
                  string anno = righe[j + 2].Substring(0, 4);
                  string descrizioneParziale3 = righe[j + 2].Substring(4).Trim();
                  str += $"{giornoMese.Trim()} {anno.Trim()} {tipoParte.Trim()} {descrizioneParziale1} {descrizioneParziale3} {descrizioneParziale2}";
                }
                else
                {
                  // ci troviamo nel caso in cui nella data c'è una parte di tipologia
                  string tipoParte1 = righe[j].Replace(giornoMese, "");
                  List<string> annoTipoParte2 = righe[j + 2].Split(" ").ToList();
                  string anno = annoTipoParte2.FirstOrDefault();
                  string tipoParte2 = "";
                  if (annoTipoParte2.Count > 1)
                  {
                    tipoParte2 = annoTipoParte2.LastOrDefault();
                  }
                  str += $"{giornoMese.Trim()} {anno.Trim()} {tipoParte1.Trim()} {tipoParte2.Trim()} {righe[j + 1].Trim()}";
                }
                lista.Add(str);
              }
            }

            for (int k = 0; k < lista.Count; k++)
            {
              string riga = lista[k].ToString();
              Operazione operazione = getOperazioneTr(riga);
              if (operazione == null)
              {
              }
              else
              {
                operazioni.Add(operazione);
              }
            }
          }
        }
      }

      return operazioni;
    }

    private Operazione getOperazioneTr(string riga)
    {
      var match = Regex.Match(riga, @"^(\d{1,2}\s\w{3}\s\d{4})\s+([A-Za-zÀ-ÿ '*&.]+)\s+(.*?)\s+([\d\.,]*)\s*€?\s*([\d\.,]*)\s*€?\s*([\d\.,]*)\s*€?$");
      Operazione op = null;
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
          tipoDescrizione = rigaSenzaData.Replace(match1.Value, "").Replace("null","").Trim();
        }

        string tipo = "";
        string descrizione = "";
        decimal entrata = 0;
        decimal uscita = 0;
        if (tipoDescrizione.Contains("Bonifico"))
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
        else if (tipoDescrizione.Contains("Transazione con carta"))
        {
          tipo = "Transazione con carta";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          uscita = ParseDecimal(valoreStr);
        }
        else if (tipoDescrizione.Contains("Commercio"))
        {
          tipo = "Commercio azioni";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          uscita = ParseDecimal(valoreStr);
        }
        else if (tipoDescrizione.Contains("Addebito diretto"))
        {
          tipo = "Addebito diretto";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          uscita = ParseDecimal(valoreStr);
        }
        else if (tipoDescrizione.Contains("Trasferimento"))
        {
          tipo = "Trasferimento";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          entrata = ParseDecimal(valoreStr);
        }
        else if (tipoDescrizione.Contains("Pagamento degli interessi"))
        {
          tipo = "Pagamento degli interessi";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          entrata = ParseDecimal(valoreStr);
        }
        else if (tipoDescrizione.Contains("Premio"))
        {
          tipo = "Premio";
          descrizione = tipoDescrizione.Replace(tipo, "").Trim();
          entrata = ParseDecimal(valoreStr);
        }

        if (entrata == 0 && uscita > 0)
        {
          op = new Operazione
          {
            Data = data,
            Tipo = tipo,
            Descrizione = descrizione,
            Importo = uscita,
            IsContabilizzato = false,
            Banca = TRADEREPUBBLIC
          };
        }
      }
      return op;
    }

    private decimal ParseDecimal(string s)
    {
      if (string.IsNullOrWhiteSpace(s)) return 0;
      s = s.Replace(".", "").Replace(",", ".").Replace("\"", "");
      decimal dd = decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;
      return dd;
    }

    private List<Operazione> EstraiMovimentiFromBpm(string fullName)
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
          decimal importoTemp = ParseDecimal(list[i][2]);
          if (importoTemp < 0)
          {
            results.Add(new Operazione
            {
              Data = Convert.ToDateTime(list[i][0]),
              Descrizione = list[i][4],
              Tipo = "",
              Importo = importoTemp,
              IsContabilizzato = true,
              Banca = BPM
            });
          }
        }
        return results;
      }
    }    

    private List<Operazione> EstraiMovimentiFromHype(string fullName)
    {
      var results = new List<Operazione>();
      List<string> rows = File.ReadAllLines(fullName).ToList();
      for (int i = 1; i < rows.Count; i++)
      {
        List<string> columns = rows[i].Split(",").ToList();

        results.Add(new Operazione
        {
          Data = DateTime.ParseExact(columns[0], "dd/MM/yyyy", CultureInfo.InvariantCulture),
          Descrizione = $"{columns[4]} {columns[5]}",
          Tipo = columns[3],
          Importo = ParseDecimal(columns[6].Replace(".", ",")),
          IsContabilizzato = false,
          Banca = HYPE
        });
      }
      return results;
    }

    private List<Operazione> EstraiMovimentiFromSplitwise(string fullName)
    {
      var results = new List<Operazione>();
      List<string> rows = File.ReadAllLines(fullName).ToList();
      for (int i = 1; i < rows.Count; i++)
      {
        if (string.IsNullOrWhiteSpace(rows[i]))
        {
          continue;
        }
        List<string> columns = rows[i].Split(",").ToList();
        results.Add(new Operazione
        {
          Data = Convert.ToDateTime(columns[0]),
          Descrizione = columns[1],
          Tipo = "",
          Importo= ParseDecimal(columns[3].Replace(".", ",")),
          ImportoRossella = ParseDecimal(columns[6].Replace(".", ",")),
          ImportoLuca = ParseDecimal(columns[5].Replace(".", ",")),
          IsContabilizzato = false,
          Banca = SPLITWISE
        });
      }
      return results;
    }

    // mesi italiani
    private readonly string[] Months = new[]
    {
        "gennaio","febbraio","marzo","aprile","maggio","giugno",
        "luglio","agosto","settembre","ottobre","novembre","dicembre"
    };
  }
}
