using BancaParser.Core;
using Microsoft.Extensions.Configuration;

namespace BancaParser.Host.Helpers
{
  public class Config
  {
    public string Secret { get; set; }
    public string RelativeFolderInputFiles { get; set; }
    public string RelativeFolderOutputFiles { get; set; }
    public Dictionary<string, string> Descrizioni { get; set; }

    public IConfiguration AppSettings { get; }

    public Config()
    {

      string path = Path.Combine(Utility.PATH_SHARED_FOLDER, "dati", "banca-parser", "appsettings.json");
      FileInfo fileInfo = new FileInfo(path);
      if (File.Exists(path))
      {
        AppSettings = new ConfigurationBuilder()
          .SetBasePath(fileInfo.Directory.FullName)
          .AddJsonFile("appsettings.json")
          .Build();
      }
      else
      {
        AppSettings = new ConfigurationBuilder()
          .SetBasePath(AppContext.BaseDirectory)
          .AddJsonFile("appsettings.json")
          .Build();
      }
      var appSettingsSection = AppSettings.GetSection("AppSettings");
      Secret = appSettingsSection["Secret"];
      RelativeFolderInputFiles = appSettingsSection["RelativeFolderInputFiles"];
      RelativeFolderOutputFiles = appSettingsSection["RelativeFolderOutputFiles"];
      Descrizioni = appSettingsSection.GetSection("Descrizioni").Get<Dictionary<string, string>>();
    }
  }
}
