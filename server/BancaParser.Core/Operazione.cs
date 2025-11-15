using System.Globalization;

namespace BancaParser.Core
{
  public class Operazione
  {
    public DateTime Data { get; set; }
    public string Tipo { get; set; } = "";
    public string Descrizione { get; set; } = "";
    public decimal Importo { get; set; }
    public decimal ImportoRossella { get; set; }
    public decimal ImportoLuca { get; set; }
    public bool IsContabilizzato { get; set; }
    public string Banca { get; set; }
    public override string ToString()
    {
      var culture = new CultureInfo("it-IT");

      string s = "";
      s += $"{this.Data:yyyy-MM-dd};";
      s += $"\"{this.Tipo}\";";
      s += $"\"{this.Descrizione}\";";
      s += $"{this.ImportoRossella.ToString(culture)};";
      s += $"{this.ImportoLuca.ToString(culture)};";
      s += $"{this.Importo.ToString(culture)};";
      s += $"{this.IsContabilizzato:\"S\":\"N\"};";
      s += $"\"{this.Banca}\"";
      return s;
    }
  }
}
