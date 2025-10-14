using System.Globalization;
using System.Reflection;
using System.Text.Json;

namespace BancaParser.Core
{
  public class Utility
  {
    public static string PATH_SHARED_FOLDER_DOCKER { get { return "/BancaParser.ApiHost/Volume_BancaParser_Host/"; } }

    public static string PATH_SHARED_FOLDER_LOCAL_MAC_OS { get { return "/Users/lucaurti/Develop/fileBrowserBancaParser/"; } }

    public static string PATH_SHARED_FOLDER_LOCAL_WIN_OS { get { return @"C:\temp\fileBrowserBancaParser\"; } }

    public static string getUnixTimeStamp(DateTime expiresDate)
    {
      DateTime startDate = new DateTime(1970, 1, 1, 0, 0, 0);
      string unixTimeStamp = Convert.ToInt64((expiresDate.Subtract(startDate)).TotalMilliseconds).ToString();
      return unixTimeStamp;
    }

    public static string PATH_SHARED_FOLDER
    {
      get
      {
        switch (Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT").ToUpper())
        {
          case "LOCAL_DEBUGGING_WIN_OS":
            return PATH_SHARED_FOLDER_LOCAL_WIN_OS;
          case "LOCAL_DEBUGGING_MAC_OS":
            return PATH_SHARED_FOLDER_LOCAL_MAC_OS;
          default:
            return PATH_SHARED_FOLDER_DOCKER;
        }
      }
    }

    public static bool IS_ENVIRONMENT_PRODUCTION
    {
      get
      {
        switch (Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT").ToUpper())
        {
          case "LOCAL_DEBUGGING_WIN_OS":
          case "LOCAL_DEBUGGING_MAC_OS":
          case "LOCAL_DOCKER":
            return false;
          default:
            return true;
        }
      }
    }



    public static string GetPathFile(string path)
    {
      string pathFile = path.Replace(Utility.PATH_SHARED_FOLDER_DOCKER, "").Replace(Utility.PATH_SHARED_FOLDER_LOCAL_WIN_OS, "").Replace(Utility.PATH_SHARED_FOLDER_LOCAL_MAC_OS, "");
      var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
      if (!String.IsNullOrEmpty(env))
      {
        switch (env.ToUpper())
        {
          case "LOCAL_DEBUGGING_WIN_OS":
            pathFile = $"{Utility.PATH_SHARED_FOLDER_LOCAL_WIN_OS}{pathFile}";
            pathFile = pathFile.Replace("/", "\\");
            return pathFile;
          case "LOCAL_DEBUGGING_MAC_OS":
            pathFile = $"{Utility.PATH_SHARED_FOLDER_LOCAL_MAC_OS}{pathFile}";
            pathFile = pathFile.Replace("\\", "/");
            break;
          default:
            pathFile = $"{Utility.PATH_SHARED_FOLDER_DOCKER}{pathFile}";
            pathFile = pathFile.Replace("\\", "/");
            break;
        }
      }
      return pathFile;
    }

    public static string GetRelativePathFile(string fullName)
    {
      string pathFile = "";
      var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
      if (!String.IsNullOrEmpty(env))
      {
        switch (env.ToUpper())
        {
          case "LOCAL_DEBUGGING_WIN_OS":
            pathFile = fullName.Replace(Utility.PATH_SHARED_FOLDER_LOCAL_WIN_OS, "");
            break;
          case "LOCAL_DEBUGGING_MAC_OS":
            pathFile = fullName.Replace(Utility.PATH_SHARED_FOLDER_LOCAL_MAC_OS,"");
            break;
          default:
            pathFile = fullName.Replace(Utility.PATH_SHARED_FOLDER_DOCKER, "");
            break;
        }
      }
      return pathFile.Replace("\\","/");
    }
  }
}