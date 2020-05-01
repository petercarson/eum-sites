using System;
using System.Configuration;
using System.Security;

namespace EITSiteProvisioning_API
{
  /// <summary>
  ///  App Settings
  /// </summary>
  public static class AppSettings
  {
    // Swagger
    internal static bool DisableSwagger = Convert.ToBoolean(ConfigurationManager.AppSettings["DisableSwagger"]);

    // SharePoint
    internal static string SPWebAppUrl = ConfigurationManager.AppSettings["SPWebAppUrl"];
    internal static string SPSitesListSiteCollectionPath = ConfigurationManager.AppSettings["SPSitesListSiteCollectionPath"];
    internal static string SPSitesListName = ConfigurationManager.AppSettings["SPSitesListName"];
    internal static string SPLandingSite = string.Concat(SPWebAppUrl, SPSitesListSiteCollectionPath);
    internal static string SPUsername = ConfigurationManager.AppSettings["SPUsername"];
    internal static SecureString SPPassword = ConfigurationManager.AppSettings["SPPassword"].ToSecureString();
    internal static int SPTimeoutMilliseconds = Convert.ToInt32(ConfigurationManager.AppSettings["SPTimeoutMilliseconds"]);
    internal static int SPListPageSize = Convert.ToInt32(ConfigurationManager.AppSettings["SPListPageSize"]);
    internal static string SPHideFromSiteListColumn = ConfigurationManager.AppSettings["SPHideFromSiteListColumn"];
    internal static string LogFile = ConfigurationManager.AppSettings["EITSiteProvisioningAPI-LogFile"];

    private static SecureString ToSecureString(this string Source)
    {
      if (string.IsNullOrWhiteSpace(Source))
      {
        return null;
      }
      else
      {
        SecureString Result = new SecureString();
        foreach (char c in Source.ToCharArray())
        {
          Result.AppendChar(c);
        }
        return Result;
      }
    }
  }
}