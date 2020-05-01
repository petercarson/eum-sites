using Serilog;
using System.Configuration;
using System.Web.Http;
using System.Web.Http.Cors;

namespace EITSiteProvisioning_API
{
  public static class WebApiConfig
  {
    public static void Register(HttpConfiguration config)
    {
      // initialize logger
      Log.Logger = new LoggerConfiguration()
          .MinimumLevel.Debug()
          .WriteTo.Console()
          .WriteTo.File(AppSettings.LogFile, outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz}] [{Level}] {SourceContext}{NewLine}{Message}{NewLine}{Exception}{NewLine}",
            rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 500000)
          .CreateLogger();

      // Web API routes
      string origins = ConfigurationManager.AppSettings["Cors.AllowedOrigins"];
      if (!string.IsNullOrWhiteSpace(origins))
      {
        config.EnableCors(new EnableCorsAttribute(origins, "*", "*") { SupportsCredentials = true });
      }
      config.MapHttpAttributeRoutes();

      config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "v1/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
    }
  }
}
