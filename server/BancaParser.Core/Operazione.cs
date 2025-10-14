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
}
