using EITSiteProvisioning_API.Filters;
using EITSiteProvisioning_API.Models;
using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.Web.Http;

namespace EITSiteProvisioning_API.Controllers
{
  [Authorize]
  public class SitesController : ApiController
  {
    readonly SharePointHelpers sharePointHelpers = new SharePointHelpers();

    [HttpGet]
    public IHttpActionResult Get(string parentUrl = "")
    {
      Log.Debug($"SitesController - GET");
      try
      {
        if (string.IsNullOrWhiteSpace(parentUrl))
        {
          return Ok(sharePointHelpers.GetSiteListItems());
        }
        else
        {
          return Ok(sharePointHelpers.GetSiteListItemsFilteredByParent(parentUrl));
        }
      }
      catch (Exception ex)
      {
        Log.Error(ex, "Failed retrieving sites.");
        return BadRequest("Failed retrieving sites.");
      }
    }

    [HttpPost]
    public IHttpActionResult Post([FromBody]JToken siteRequest)
    {
      Log.Debug($"SitesController - POST");
      try
      {
        JObject siteRequestJObject = JObject.Parse(siteRequest.ToString());

        string requestor = User.Identity.Name?.Substring(User.Identity.Name.LastIndexOf('|') + 1);
        Log.Debug($"Requestor = {requestor}");
        sharePointHelpers.AddSiteRequest(siteRequestJObject, requestor);

        return Ok(siteRequest);
      }
      catch (Exception ex)
      {
        Log.Error(ex, "Failed saving site request.");
        return BadRequest("Failed saving site request.");
      }
    }
  }
}
