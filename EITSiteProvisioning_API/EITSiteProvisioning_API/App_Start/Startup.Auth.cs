using System;
using System.Collections.Generic;
using System.Configuration;
using System.IdentityModel.Tokens;
using System.Linq;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.ActiveDirectory;
using Owin;

namespace EITSiteProvisioning_API
{
  public partial class Startup
  {
    // For more information on configuring authentication, please visit https://go.microsoft.com/fwlink/?LinkId=301864
    public void ConfigureAuth(IAppBuilder app)
    {
      string azureADClientID = ConfigurationManager.AppSettings["AzureADClientID"];

      if (!string.IsNullOrWhiteSpace(azureADClientID))
      {
        app.UseWindowsAzureActiveDirectoryBearerAuthentication(
            new WindowsAzureActiveDirectoryBearerAuthenticationOptions
            {
              Tenant = ConfigurationManager.AppSettings["AzureADTenant"],
              TokenValidationParameters = new TokenValidationParameters
              {
                ValidAudience = azureADClientID
              },
            });
      }
    }
  }
}