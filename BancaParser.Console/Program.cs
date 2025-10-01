using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using PdfSharpCore.Pdf.Content;
using PdfSharpCore.Pdf.Content.Objects;

class Operazione
{
    public DateTime Data { get; set; }
    public string Descrizione { get; set; } = "";
    public decimal Importo { get; set; }
    public string Mese => Data.ToString("yyyy-MM");
}

class Program
{
    static string PdfFolder = "estratti_pdf";
    static string OutputOperazioni = "operazioni.csv";
    static string OutputRiepilogo = "riepilogo_mensile.csv";

    static void Main()
    {
        var operazioni = new List<Operazione>();

        foreach (var file in Directory.GetFiles(PdfFolder, "*.pdf"))
        {
            operazioni.AddRange(EstrattiDaPdf(file));
        }

        if (operazioni.Count == 0)
        {
            Console.WriteLine("⚠️ Nessuna operazione trovata.");
            return;
        }

        // Salva tutte le operazioni in CSV
        using (var sw = new StreamWriter(OutputOperazioni))
        {
            sw.WriteLine("Data,Descrizione,Importo,Mese");
            foreach (var op in operazioni)
            {
                sw.WriteLine($"{op.Data:yyyy-MM-dd},\"{op.Descrizione}\",{op.Importo.ToString(CultureInfo.InvariantCulture)},{op.Mese}");
            }
        }

        // Riepilogo mensile
        var riepilogo = operazioni
            .GroupBy(o => o.Mese)
            .Select(g => new { Mese = g.Key, Totale = g.Sum(x => x.Importo) })
            .OrderBy(x => x.Mese)
            .ToList();

        using (var sw = new StreamWriter(OutputRiepilogo))
        {
            sw.WriteLine("Mese,Totale");
            foreach (var r in riepilogo)
            {
                sw.WriteLine($"{r.Mese},{r.Totale.ToString(CultureInfo.InvariantCulture)}");
            }
        }

        Console.WriteLine($"✅ File CSV creati:\n - {OutputOperazioni}\n - {OutputRiepilogo}");
    }

    static List<Operazione> EstrattiDaPdf(string pdfPath)
    {
        var operazioni = new List<Operazione>();

        using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.ReadOnly))
        {
            foreach (var page in document.Pages)
            {
                string testo = EstraiTestoDaPagina(page);

                foreach (var riga in testo.Split('\n'))
                {
                    if (Regex.IsMatch(riga, @"\d{2}\.\d{2}\.\d{4}") &&
                        (riga.Contains("Acquisto") || riga.Contains("Vendita") || riga.Contains("Dividendo") ||
                         riga.Contains("Deposito") || riga.Contains("Prelievo") || riga.Contains("Commissione")))
                    {
                        var parti = riga.Split(" ", StringSplitOptions.RemoveEmptyEntries);
                        try
                        {
                            var data = DateTime.ParseExact(parti[0], "dd.MM.yyyy", CultureInfo.InvariantCulture);

                            var importoStr = parti.Last()
                                .Replace("€", "")
                                .Replace(".", "")
                                .Replace(",", ".")
                                .Trim();

                            if (decimal.TryParse(importoStr, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal importo))
                            {
                                var descrizione = string.Join(" ", parti.Skip(1).Take(parti.Length - 2));
                                operazioni.Add(new Operazione
                                {
                                    Data = data,
                                    Descrizione = descrizione,
                                    Importo = importo
                                });
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        return operazioni;
    }

    static string EstraiTestoDaPagina(PdfSharpCore.Pdf.PdfPage page)
    {
        var content = ContentReader.ReadContent(page);
        return EstraiTesto(content);
    }

    static string EstraiTesto(CObject cObject)
    {
        string risultato = "";

        if (cObject is COperator op && op.Operands != null)
        {
            foreach (var operand in op.Operands)
            {
                risultato += EstraiTesto(operand);
            }
        }
        else if (cObject is CSequence seq)
        {
            foreach (var element in seq)
            {
                risultato += EstraiTesto(element);
            }
        }
        else if (cObject is CString str)
        {
            risultato += str.Value + " ";
        }

        return risultato;
    }
}