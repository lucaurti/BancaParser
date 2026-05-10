using ClosedXML.Excel;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace BancaParser.Core
{
  public class RecuperaOperazioniResult
  {
    public List<Operazione> Operazioni { get; set; } = new List<Operazione>();
    public List<FileElaborazioneErrore> Errori { get; set; } = new List<FileElaborazioneErrore>();
  }

  public class FileElaborazioneErrore
  {
    public string File { get; set; } = "";
    public string Errore { get; set; } = "";
  }

  public class PdfMovimentiExtractor
  {
    private const string SPLITWISE = "Splitwise";
    private const string HYPE = "Hype";
    private const string BPM = "BPM";
    private const string SATISPAY = "Satispay";
    private const string TRADEREPUBBLIC = "Trade Republic";
    private const string ING = "ING";

    public RecuperaOperazioniResult RecuperaOperazioni(List<string> files)
    {
      RecuperaOperazioniResult result = new RecuperaOperazioniResult();
      foreach (var file in files)
      {
        FileInfo fileInfo = new FileInfo(file);
        try
        {
          switch (fileInfo.Name.Replace(fileInfo.Extension, "").ToLower())
          {
            case string s when s.StartsWith("tr"):
            case string s1 when s1.StartsWith("trade_repubblic"):
            case string s2 when s2.StartsWith("traderepubblic"):
              if (fileInfo.Extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
              {
                result.Operazioni.AddRange(EstraiMovimentiFromTrCsv(fileInfo.FullName));
              }
              else
              {
                result.Operazioni.AddRange(EstraiMovimentiFromTr(fileInfo.FullName));
              }
              break;
            case string s when s.StartsWith("ing"):
            case string s1 when s1.StartsWith("ing direct"):
            case string s2 when s2.StartsWith("ingdirect"):
            case string s3 when s3.StartsWith("ing_direct"):
              //operazioni.AddRange(EstraiMovimentiFromIng(fileInfo.FullName));
              result.Operazioni.AddRange(EstraiGrigliaPdfConRigheMultiple(fileInfo.FullName));
              break;
            case string s when s.StartsWith("bpm"):
              result.Operazioni.AddRange(EstraiMovimentiFromBpm(fileInfo.FullName));
              break;
            case string s when s.StartsWith("hype"):
              result.Operazioni.AddRange(EstraiMovimentiFromHype(fileInfo.FullName));
              break;
            case string s when s.StartsWith("splitwise"):
              result.Operazioni.AddRange(EstraiMovimentiFromSplitwise(fileInfo.FullName));
              break;
            case string s when s.StartsWith("satispay"):
              result.Operazioni.AddRange(EstraiMovimentiFromSatispay(fileInfo.FullName));
              break;
          }
        }
        catch (Exception ex)
        {
          result.Errori.Add(new FileElaborazioneErrore
          {
            File = fileInfo.Name,
            Errore = ex.Message
          });
        }
      }
      return result;
    }

    private IEnumerable<Operazione> EstraiMovimentiFromSatispay(string fullName)
    {
      using (var workbook = new XLWorkbook(fullName))
      {
        var worksheet = workbook.Worksheet(1);
        var results = new List<Operazione>();
        foreach (var row in worksheet.RowsUsed().Skip(1))
        {
          decimal importoTemp = ParseDecimal(row.Cell(4).GetValue<string>());
          if (importoTemp < 0)
          {
            results.Add(new Operazione
            {
              Data = ParseSatispayDate(row.Cell(1)),
              Descrizione = row.Cell(2).GetValue<string>(),
              Tipo = "",
              Importo = importoTemp,
              IsContabilizzato = true,
              Banca = SATISPAY
            });
          }
        }
        return results;
      }
    }

    private DateTime ParseSatispayDate(IXLCell cell)
    {
      if (cell.TryGetValue<DateTime>(out var date))
      {
        return date;
      }

      return ParseSatispayDate(cell.GetValue<string>());
    }

    private DateTime ParseSatispayDate(string value)
    {
      value = value.Trim().Trim('"');

      string[] formatiData =
      {
        "dd/MM/yyyy HH:mm:ss",
        "d/M/yyyy H:mm:ss",
        "MM/dd/yyyy HH:mm:ss",
        "M/d/yyyy H:mm:ss",
        "dd/MM/yyyy",
        "d/M/yyyy",
        "MM/dd/yyyy",
        "M/d/yyyy",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-dd"
      };

      if (DateTime.TryParseExact(value, formatiData, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
      {
        return parsedDate;
      }

      if (DateTime.TryParse(value, new CultureInfo("it-IT"), DateTimeStyles.None, out parsedDate)
          || DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate)
          || DateTime.TryParse(value, new CultureInfo("en-US"), DateTimeStyles.None, out parsedDate))
      {
        return parsedDate;
      }

      throw new FormatException($"Formato data Satispay non riconosciuto: '{value}'.");
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
      int start = fullText.IndexOf("(EUR)to (EUR)") + 13;
      string movimentiString = fullText.Substring(start, fullText.Length - start);
      movimentiString = movimentiString.Replace("TipoTipo", "");
      movimentiString = movimentiString.Replace("ImporImportoto", "");
      movimentiString = movimentiString.Replace("DatData Ca Contontabileabile DescrizioneDescrizione", "");
      movimentiString = movimentiString.Replace("TTrransazioneansazione (EUR)(EUR)", "");
      movimentiString = movimentiString.Replace("TTrransazioneansazione", "");
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
        int anno = int.Parse(dataGiornoMeseImportoTipoList[2]);
        decimal importo = ParseDecimal(dataGiornoMeseImportoTipoList[4]);
        string tipoTransazione = "";//dataGiornoMeseImportoTipoList[2];
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
          tipoDescrizione = rigaSenzaData.Replace(match1.Value, "").Replace("null", "").Trim();
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

    public List<Operazione> EstraiGrigliaPdfConRigheMultiple(string pdfPath)
    {
      return EstraiGrigliaPdfConRigheMultiple(pdfPath, 95, 245, 445);
    }

    public List<Operazione> EstraiGrigliaPdfConRigheMultiple(string pdfPath, float limiteColonna1, float limiteColonna2, float limiteColonna3)
    {
      List<List<string>> macroRighe = EstraiMacroRighePdfConRigheMultiple(pdfPath, limiteColonna1, limiteColonna2, limiteColonna3);
      var operazioni = new List<Operazione>();

      foreach (List<string> macroRiga in macroRighe)
      {
        if (macroRiga.Count < 4)
        {
          continue;
        }
        var importo = ParseDecimal(macroRiga[3]);
        if (importo > 0)
        {
          continue;
        }
        operazioni.Add(new Operazione
        {
          Data = ParseDataMacroRiga(macroRiga[0]),
          Tipo = "",
          Descrizione = $"{macroRiga[1]} {macroRiga[2]}".Trim(),
          Importo = importo,
          IsContabilizzato = false,
          Banca = ING
        });
      }

      return operazioni;
    }

    public List<List<string>> EstraiMacroRighePdfConRigheMultiple(string pdfPath)
    {
      return EstraiMacroRighePdfConRigheMultiple(pdfPath, 95, 245, 445);
    }

    public List<List<string>> EstraiMacroRighePdfConRigheMultiple(string pdfPath, float limiteColonna1, float limiteColonna2, float limiteColonna3)
    {
      var macroRighe = new List<List<string>>();
      List<string>? macroRigaCorrente = null;
      List<List<PdfTextFragment>> righeFisiche = EstraiRigheFisichePdf(pdfPath);

      foreach (List<PdfTextFragment> rigaFisica in righeFisiche)
      {
        List<string> colonne = EstraiQuattroColonne(rigaFisica, limiteColonna1, limiteColonna2, limiteColonna3);
        NormalizzaDataSpezzataPrimaColonna(colonne);

        if (ContieneDataPrimaColonna(colonne))
        {
          macroRigaCorrente = colonne;
          macroRighe.Add(macroRigaCorrente);
        }
        else if (macroRigaCorrente != null)
        {
          AccodaSottorigaAMacroRiga(macroRigaCorrente, colonne);
        }
      }

      return macroRighe;
    }

    private DateTime ParseDataMacroRiga(string data)
    {
      string[] formatiData = { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "d MMM yyyy", "dd MMM yyyy", "d MMMM yyyy", "dd MMMM yyyy" };
      string dataNormalizzata = EstraiDataIniziale(data);

      if (DateTime.TryParseExact(dataNormalizzata, formatiData, new CultureInfo("it-IT"), DateTimeStyles.None, out DateTime dataParsed))
      {
        return dataParsed;
      }

      return DateTime.ParseExact(dataNormalizzata, formatiData, CultureInfo.InvariantCulture, DateTimeStyles.None);
    }

    private string EstraiDataIniziale(string testo)
    {
      Match match = Regex.Match(
          testo.Trim(),
          @"^(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{1,2}\s+[A-Za-zÀ-ÿ]+\s+\d{4})");

      return match.Success ? match.Groups[1].Value.Trim() : testo.Trim();
    }

    public List<PdfTextFragment> EstraiFrammentiPdfConCoordinate(string pdfPath)
    {
      var fragments = new List<PdfTextFragment>();

      using (var reader = new PdfReader(pdfPath))
      using (var pdf = new PdfDocument(reader))
      {
        for (int pageNumber = 1; pageNumber <= pdf.GetNumberOfPages(); pageNumber++)
        {
          var listener = new TextCoordinateExtractionListener(pageNumber);
          var processor = new PdfCanvasProcessor(listener);
          processor.ProcessPageContent(pdf.GetPage(pageNumber));
          fragments.AddRange(listener.Fragments);
        }
      }

      return fragments
          .OrderBy(f => f.Page)
          .ThenByDescending(f => f.Y)
          .ThenBy(f => f.X)
          .ToList();
    }

    private List<List<PdfTextFragment>> EstraiRigheFisichePdf(string pdfPath)
    {
      const float tolleranzaY = 2.5f;
      var righe = new List<List<PdfTextFragment>>();

      foreach (var pageGroup in EstraiFrammentiPdfConCoordinate(pdfPath).GroupBy(f => f.Page).OrderBy(g => g.Key))
      {
        foreach (PdfTextFragment fragment in pageGroup.OrderByDescending(f => f.Y).ThenBy(f => f.X))
        {
          List<PdfTextFragment>? riga = righe
              .Where(r => r[0].Page == fragment.Page)
              .FirstOrDefault(r => Math.Abs(r.Average(f => f.Y) - fragment.Y) <= tolleranzaY);

          if (riga == null)
          {
            riga = new List<PdfTextFragment>();
            righe.Add(riga);
          }

          riga.Add(fragment);
        }
      }

      return righe
          .OrderBy(r => r[0].Page)
          .ThenByDescending(r => r.Average(f => f.Y))
          .ToList();
    }

    private List<string> EstraiQuattroColonne(List<PdfTextFragment> riga, float limiteColonna1, float limiteColonna2, float limiteColonna3)
    {
      var colonne = new StringBuilder[] { new StringBuilder(), new StringBuilder(), new StringBuilder(), new StringBuilder() };
      var ultimoFrammentoPerColonna = new PdfTextFragment?[4];

      foreach (PdfTextFragment fragment in riga.OrderBy(f => f.X))
      {
        int indiceColonna = GetIndiceColonna(fragment.X, limiteColonna1, limiteColonna2, limiteColonna3);
        PdfTextFragment? ultimoFrammento = ultimoFrammentoPerColonna[indiceColonna];

        if (ultimoFrammento != null && fragment.X - ultimoFrammento.EndX > 2.2f && colonne[indiceColonna].Length > 0)
        {
          colonne[indiceColonna].Append(' ');
        }

        colonne[indiceColonna].Append(fragment.Text);
        ultimoFrammentoPerColonna[indiceColonna] = fragment;
      }

      return colonne
          .Select(c => Regex.Replace(c.ToString(), @"\s+", " ").Trim())
          .ToList();
    }

    private void NormalizzaDataSpezzataPrimaColonna(List<string> colonne)
    {
      if (colonne.Count < 2 || string.IsNullOrWhiteSpace(colonne[0]) || string.IsNullOrWhiteSpace(colonne[1]))
      {
        return;
      }

      Match primaColonna = Regex.Match(colonne[0], @"^(?<giorno>\d{1,2})\s+(?<mese>[A-Za-zÀ-ÿ]+)\s+(?<annoParziale>\d{2})$");
      Match secondaColonna = Regex.Match(colonne[1], @"^(?<restoAnno>\d{2})(?<resto>.*)$");
      if (!primaColonna.Success || !secondaColonna.Success)
      {
        return;
      }

      colonne[0] = $"{primaColonna.Groups["giorno"].Value} {primaColonna.Groups["mese"].Value} {primaColonna.Groups["annoParziale"].Value}{secondaColonna.Groups["restoAnno"].Value}";
      colonne[1] = secondaColonna.Groups["resto"].Value.Trim();
    }

    private int GetIndiceColonna(float x, float limiteColonna1, float limiteColonna2, float limiteColonna3)
    {
      if (x < limiteColonna1)
      {
        return 0;
      }

      if (x < limiteColonna2)
      {
        return 1;
      }

      if (x < limiteColonna3)
      {
        return 2;
      }

      return 3;
    }

    private bool ContieneDataPrimaColonna(List<string> colonne)
    {
      string[] formatiData = { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "d MMM yyyy", "dd MMM yyyy", "d MMMM yyyy", "dd MMMM yyyy" };
      if (colonne.Count == 0 || string.IsNullOrWhiteSpace(colonne[0]))
      {
        return false;
      }

      string dataNormalizzata = EstraiDataIniziale(colonne[0]);

      return DateTime.TryParseExact(
          dataNormalizzata,
          formatiData,
          CultureInfo.InvariantCulture,
          DateTimeStyles.None,
          out _)
          || DateTime.TryParseExact(
              dataNormalizzata,
              formatiData,
              new CultureInfo("it-IT"),
              DateTimeStyles.None,
              out _);
    }

    private void AccodaSottorigaAMacroRiga(List<string> macroRigaCorrente, List<string> colonne)
    {
      int colonneValorizzate = colonne.Count(c => !string.IsNullOrWhiteSpace(c));
      if (colonneValorizzate == 1 && !string.IsNullOrWhiteSpace(colonne[0]))
      {
        AccodaTestoInColonna(macroRigaCorrente, 2, colonne[0]);
        return;
      }

      for (int i = 0; i < Math.Min(4, colonne.Count); i++)
      {
        AccodaTestoInColonna(macroRigaCorrente, i, colonne[i]);
      }
    }

    private void AccodaTestoInColonna(List<string> colonne, int indiceColonna, string testo)
    {
      if (string.IsNullOrWhiteSpace(testo))
      {
        return;
      }

      colonne[indiceColonna] = $"{colonne[indiceColonna]} {testo}".Trim();
    }

    public class PdfTextFragment
    {
      public int Page { get; set; }
      public string Text { get; set; } = "";
      public float X { get; set; }
      public float EndX { get; set; }
      public float Y { get; set; }
    }

    private class TextCoordinateExtractionListener : IEventListener
    {
      private readonly int pageNumber;

      public TextCoordinateExtractionListener(int pageNumber)
      {
        this.pageNumber = pageNumber;
      }

      public List<PdfTextFragment> Fragments { get; } = new List<PdfTextFragment>();

      public void EventOccurred(IEventData data, EventType type)
      {
        if (type != EventType.RENDER_TEXT)
        {
          return;
        }

        var renderInfo = (TextRenderInfo)data;
        foreach (TextRenderInfo characterInfo in renderInfo.GetCharacterRenderInfos())
        {
          string text = characterInfo.GetText();
          if (string.IsNullOrWhiteSpace(text))
          {
            continue;
          }

          LineSegment baseline = characterInfo.GetBaseline();
          Vector startPoint = baseline.GetStartPoint();
          Vector endPoint = baseline.GetEndPoint();

          Fragments.Add(new PdfTextFragment
          {
            Page = pageNumber,
            Text = text,
            X = startPoint.Get(Vector.I1),
            EndX = endPoint.Get(Vector.I1),
            Y = startPoint.Get(Vector.I2)
          });
        }
      }

      public ICollection<EventType> GetSupportedEvents()
      {
        return new HashSet<EventType> { EventType.RENDER_TEXT };
      }
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
              Data = DateTime.ParseExact(list[i][0], "dd/MM/yyyy", CultureInfo.InvariantCulture),
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
          Importo = ParseDecimal(columns[3].Replace(".", ",")),
          ImportoRossella = ParseDecimal(columns[6].Replace(".", ",")),
          ImportoLuca = ParseDecimal(columns[5].Replace(".", ",")),
          IsContabilizzato = false,
          Banca = SPLITWISE
        });
      }
      return results;
    }

    private List<Operazione> EstraiMovimentiFromTrCsv(string fullName)
    {
      var results = new List<Operazione>();
      List<List<string>> rows = ParseCsvRows(fullName);
      if (rows.Count <= 1)
      {
        return results;
      }

      List<string> header = rows[0];
      int typeIndex = header.FindIndex(c => c.Equals("type", StringComparison.OrdinalIgnoreCase));

      for (int i = 1; i < rows.Count; i++)
      {
        List<string> columns = rows[i];
        if (columns.Count == 0 || columns.All(string.IsNullOrWhiteSpace))
        {
          continue;
        }

        if (columns.Count <= 16)
        {
          continue;
        }

        decimal importoTemp = ParseDecimal(columns[10].Replace(".", ","));
        if (importoTemp > 0)
        {
          continue;
        }

        string tipo = typeIndex >= 0 && typeIndex < columns.Count
            ? TraduciTipoTrCsv(columns[typeIndex])
            : "";

        results.Add(new Operazione
        {
          Data = Convert.ToDateTime(columns[1], CultureInfo.InvariantCulture),
          Descrizione = columns[17],
          Tipo = tipo,
          Importo = importoTemp,
          IsContabilizzato = false,
          Banca = TRADEREPUBBLIC
        });
      }

      return results;
    }

    private string TraduciTipoTrCsv(string type)
    {
      if (string.IsNullOrWhiteSpace(type))
      {
        return "";
      }

      string normalizedType = type.Trim();
      return normalizedType.ToUpperInvariant() switch
      {
        "CARD_TRANSACTION" => "Transazione con carta",
        "CARD_SUCCESSFUL" => "Transazione con carta",
        "CARD_REFUND" => "Rimborso carta",
        "CASH_INTEREST" => "Pagamento degli interessi",
        "INTEREST" => "Interessi",
        "INCOMING_TRANSFER" => "Bonifico in entrata",
        "OUTGOING_TRANSFER" => "Bonifico in uscita",
        "ORDER_EXECUTED" => "Commercio azioni",
        "SAVINGS_PLAN_EXECUTED" => "Piano di accumulo",
        "DIRECT_DEBIT" => "Addebito diretto",
        "DIVIDEND" => "Dividendo",
        "TAX" => "Imposta",
        "BENEFIT" => "Premio",
        "ROUND_UP" => "Arrotondamento",
        _ => CapitalizeWords(normalizedType.Replace("_", " "))
      };
    }

    private List<List<string>> ParseCsvRows(string fullName)
    {
      var rows = new List<List<string>>();
      var row = new List<string>();
      var field = new StringBuilder();
      bool inQuotes = false;
      string text = File.ReadAllText(fullName);

      for (int i = 0; i < text.Length; i++)
      {
        char current = text[i];

        if (current == '"')
        {
          if (inQuotes && i + 1 < text.Length && text[i + 1] == '"')
          {
            field.Append('"');
            i++;
          }
          else
          {
            inQuotes = !inQuotes;
          }
        }
        else if (current == ',' && !inQuotes)
        {
          row.Add(field.ToString());
          field.Clear();
        }
        else if ((current == '\r' || current == '\n') && !inQuotes)
        {
          if (current == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
          {
            i++;
          }

          row.Add(field.ToString());
          field.Clear();
          rows.Add(row);
          row = new List<string>();
        }
        else
        {
          field.Append(current);
        }
      }

      if (field.Length > 0 || row.Count > 0)
      {
        row.Add(field.ToString());
        rows.Add(row);
      }

      return rows;
    }

    // mesi italiani
    private readonly string[] Months = new[]
    {
        "gennaio","febbraio","marzo","aprile","maggio","giugno",
        "luglio","agosto","settembre","ottobre","novembre","dicembre"
    };
  }
}
