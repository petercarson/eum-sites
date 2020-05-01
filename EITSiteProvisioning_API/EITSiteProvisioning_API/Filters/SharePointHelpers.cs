using EITSiteProvisioning_API.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EITSiteProvisioning_API.Filters
{
  public class SharePointHelpers
  {
    public const string LookupFieldType = "Lookup";
    public const string LookupMultiFieldType = "LookupMulti";
    public const string TaxonomyFieldType = "Taxonomy";
    public const string TaxonomyMultiFieldType = "TaxonomyMulti";
    public const string UrlFieldType = "Url";
    public const string ChoiceMultiFieldType = "ChoiceMulti";
    public const string PersonFieldType = "Person";
    public const string PersonMultiFieldType = "PersonMulti";

    public ClientContext GetSPContext()
    {
      string spSiteUrl = string.Concat(AppSettings.SPWebAppUrl, AppSettings.SPSitesListSiteCollectionPath);
      bool isSharePointOnline = spSiteUrl.ToLowerInvariant().Contains("sharepoint.com");

      ClientContext spContext = null;
      AuthenticationManager authenticationManager = new AuthenticationManager();
      if (isSharePointOnline)
      {
        spContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(spSiteUrl, AppSettings.SPUsername, AppSettings.SPPassword);
      }
      else
      {
        spContext = authenticationManager.GetNetworkCredentialAuthenticatedContext(spSiteUrl, AppSettings.SPUsername, AppSettings.SPPassword, string.Empty);
      }

      spContext.RequestTimeout = AppSettings.SPTimeoutMilliseconds;

      return spContext;
    }

    public List<SiteListItem> GetSiteListItems()
    {
      List<SiteListItem> siteListItems = new List<SiteListItem>();

      ClientContext spContext = GetSPContext();

      string viewXml = $@"
        <View>
            <Query>
                <Where>
                  <And>
                    <IsNotNull>
                      <FieldRef Name='EUMSiteCreated'/>
                    </IsNotNull>
                    <Neq><FieldRef Name='{AppSettings.SPHideFromSiteListColumn}'/><Value Type='Text'>Hidden</Value></Neq>
                  </And>
                </Where>
            </Query>
            <OrderBy><FieldRef Name='Title' Ascending='TRUE' /></OrderBy>
            <RowLimit Paged='TRUE'>{AppSettings.SPListPageSize}</RowLimit>
        </View>";

      CamlQuery camlQuery = new CamlQuery
      {
        ViewXml = viewXml,
        DatesInUtc = true
      };

      List<ListItem> listItems = new List<ListItem>();
      ListItemCollectionPosition position = null;
      do
      {
        camlQuery.ListItemCollectionPosition = position;

        ListItemCollection spListItems = spContext.Web.Lists.GetByTitle(AppSettings.SPSitesListName).GetItems(camlQuery);
        spContext.Load(spListItems);
        spContext.ExecuteQueryRetry();

        position = spListItems.ListItemCollectionPosition;

        listItems.AddRange(spListItems.ToList());
      }
      while (position != null);

      if (listItems.Count > 0)
      {
        foreach (ListItem spListItem in listItems)
        {
          SiteListItem siteListItem = new SiteListItem();

          siteListItem.Id = spListItem.Id;
          siteListItem.Title = spListItem["Title"]?.ToString();
          siteListItem.EUMSiteURL = spListItem["EUMSiteURL"]?.ToString();
          siteListItem.EUMParentURL = spListItem["EUMParentURL"]?.ToString();
          siteListItem.EUMGroupSummary = spListItem["EUMGroupSummary"]?.ToString();
          siteListItem.EUMAlias = spListItem["EUMAlias"]?.ToString();
          siteListItem.EUMSiteVisibility = spListItem["EUMSiteVisibility"]?.ToString();
          siteListItem.SitePurpose = spListItem["SitePurpose"]?.ToString();
          siteListItem.EUMSiteCreated = spListItem["EUMSiteCreated"]?.ToString();

          if (spListItem["EUMSiteTemplate"] != null)
          {
            ILookupFieldValue siteTemplateLookup = new ILookupFieldValue();
            siteTemplateLookup.Id = ((FieldLookupValue)spListItem["EUMSiteTemplate"]).LookupId;
            siteTemplateLookup.Title = ((FieldLookupValue)spListItem["EUMSiteTemplate"]).LookupValue;

            siteListItem.EUMSiteTemplate = siteTemplateLookup;
          }

          if (spListItem["EUMDivision"] != null)
          {
            ILookupFieldValue divisionLookup = new ILookupFieldValue();
            divisionLookup.Id = ((FieldLookupValue)spListItem["EUMDivision"]).LookupId;
            divisionLookup.Title = ((FieldLookupValue)spListItem["EUMDivision"]).LookupValue;

            siteListItem.EUMDivision = divisionLookup;
          }

          siteListItems.Add(siteListItem);
        }
      }

      return siteListItems.OrderBy(s => s.Title).ToList();
    }

    public List<SiteListItem> GetSiteListItemsFilteredByParent(string parentUrl)
    {
      List<SiteListItem> siteListItems = new List<SiteListItem>();

      ClientContext spContext = GetSPContext();

      string siteUrlAbsolute = $"{AppSettings.SPWebAppUrl}{parentUrl}";

      string viewXml = $@"
        <View>
            <Query>
                <Where>
                  <And>
                    <And>
                      <IsNotNull>
                        <FieldRef Name='EUMSiteCreated'/>
                      </IsNotNull>
                      <Neq><FieldRef Name='{AppSettings.SPHideFromSiteListColumn}'/><Value Type='Text'>Hidden</Value></Neq>
                    </And>
                    <Or>
                      <Eq><FieldRef Name='EUMParentURL'/><Value Type='Text'>{parentUrl}</Value></Eq>
                      <Eq><FieldRef Name='EUMParentURL'/><Value Type='Text'>{siteUrlAbsolute}</Value></Eq>
                    </Or>
                  </And>
                </Where>
            </Query>
            <OrderBy><FieldRef Name='Title' Ascending='TRUE' /></OrderBy>
            <RowLimit Paged='TRUE'>{AppSettings.SPListPageSize}</RowLimit>
        </View>";

      CamlQuery camlQuery = new CamlQuery
      {
        ViewXml = viewXml,
        DatesInUtc = true
      };

      List<ListItem> listItems = new List<ListItem>();
      ListItemCollectionPosition position = null;
      do
      {
        camlQuery.ListItemCollectionPosition = position;

        ListItemCollection spListItems = spContext.Web.Lists.GetByTitle(AppSettings.SPSitesListName).GetItems(camlQuery);
        spContext.Load(spListItems);
        spContext.ExecuteQueryRetry();

        position = spListItems.ListItemCollectionPosition;

        listItems.AddRange(spListItems.ToList());
      }
      while (position != null);

      if (listItems.Count > 0)
      {
        foreach (ListItem spListItem in listItems)
        {
          SiteListItem siteListItem = new SiteListItem();

          siteListItem.Id = spListItem.Id;
          siteListItem.Title = spListItem["Title"]?.ToString();
          siteListItem.EUMSiteURL = spListItem["EUMSiteURL"]?.ToString();
          siteListItem.EUMParentURL = spListItem["EUMParentURL"]?.ToString();
          siteListItem.EUMGroupSummary = spListItem["EUMGroupSummary"]?.ToString();
          siteListItem.EUMAlias = spListItem["EUMAlias"]?.ToString();
          siteListItem.EUMSiteVisibility = spListItem["EUMSiteVisibility"]?.ToString();
          siteListItem.SitePurpose = spListItem["SitePurpose"]?.ToString();
          siteListItem.EUMSiteCreated = spListItem["EUMSiteCreated"]?.ToString();

          if (spListItem["EUMSiteTemplate"] != null)
          {
            ILookupFieldValue siteTemplateLookup = new ILookupFieldValue();
            siteTemplateLookup.Id = ((FieldLookupValue)spListItem["EUMSiteTemplate"]).LookupId;
            siteTemplateLookup.Title = ((FieldLookupValue)spListItem["EUMSiteTemplate"]).LookupValue;

            siteListItem.EUMSiteTemplate = siteTemplateLookup;
          }

          if (spListItem["EUMDivision"] != null)
          {
            ILookupFieldValue divisionLookup = new ILookupFieldValue();
            divisionLookup.Id = ((FieldLookupValue)spListItem["EUMDivision"]).LookupId;
            divisionLookup.Title = ((FieldLookupValue)spListItem["EUMDivision"]).LookupValue;

            siteListItem.EUMDivision = divisionLookup;
          }

          siteListItems.Add(siteListItem);
        }
      }

      return siteListItems.OrderBy(s => s.Title).ToList();
    }

    public void AddSiteRequest(JObject siteRequest, string requestor)
    {
      ClientContext spContext = GetSPContext();
      List spList = spContext.Web.Lists.GetByTitle(AppSettings.SPSitesListName);

      ListItemCreationInformation spListItemCreationInformation = new ListItemCreationInformation();
      ListItem spListItem = spList.AddItem(spListItemCreationInformation);

      string fieldName = string.Empty;
      foreach (var siteRequestField in siteRequest)
      {
        fieldName = siteRequestField.Key;

        var fieldValueObject = siteRequestField.Value;
        if (fieldValueObject?.GetType() == typeof(JValue))
        {
          var fieldValue = fieldValueObject.Value<string>();
          spListItem[fieldName] = fieldValue;
        }
        else if (fieldValueObject?.GetType() == typeof(JObject))
        {
          var fieldValue = fieldValueObject.Value<JToken>();

          string type = fieldValue["type"]?.ToString();

          ProcessField(fieldName, type, fieldValue, ref spListItem);
        }
      }

      spListItem["Author"] = spContext.Web.EnsureUser(requestor);
      spListItem.Update();

      spContext.ExecuteQuery();
    }

    private void ProcessField(string fieldName, string fieldType, JToken fieldValue, ref ListItem spListItem)
    {
      switch (fieldType)
      {
        case LookupFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            FieldLookupValue fieldLookupValue = new FieldLookupValue();
            fieldLookupValue.LookupId = Convert.ToInt32(fieldValue["value"]);

            spListItem[fieldName] = fieldLookupValue;
          }
          break;
        case LookupMultiFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            string[] lookupIds = fieldValue["value"].ToString().Split(',');
            FieldLookupValue[] fieldLookupValues = new FieldLookupValue[lookupIds.Length];
            for (int i = 0; i < lookupIds.Length; i++)
            {
              FieldLookupValue lookupValue = new FieldLookupValue();
              lookupValue.LookupId = Convert.ToInt32(lookupIds[i]);

              fieldLookupValues[i] = lookupValue;
            }

            spListItem[fieldName] = fieldLookupValues;
          }
          break;
        case ChoiceMultiFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["results"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            var choices = fieldValue["results"].Values<string>();
            spListItem[fieldName] = choices;
          }
          break;
        case TaxonomyFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["Label"].ToString()) || string.IsNullOrWhiteSpace(fieldValue["TermGuid"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            TaxonomyFieldValue taxonomyFieldValue = new TaxonomyFieldValue();
            taxonomyFieldValue.Label = fieldValue["Label"].ToString();
            taxonomyFieldValue.TermGuid = fieldValue["TermGuid"].ToString();
            taxonomyFieldValue.WssId = -1;

            spListItem[fieldName] = taxonomyFieldValue;
          }
          break;
        case TaxonomyMultiFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            spListItem[fieldName] = fieldValue["value"].ToString();
          }
          break;
        case UrlFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            FieldUrlValue fieldUrlValue = new FieldUrlValue();
            fieldUrlValue.Url = fieldValue["value"].ToString();
            fieldUrlValue.Description = fieldValue["value"].ToString();

            spListItem[fieldName] = fieldUrlValue;
          }
          break;
        case PersonFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            FieldUserValue fieldUserValue = new FieldUserValue();
            fieldUserValue.LookupId = Convert.ToInt32(fieldValue["value"]);

            spListItem[fieldName] = fieldUserValue;
          }
          break;
        case PersonMultiFieldType:
          if (string.IsNullOrWhiteSpace(fieldValue["value"].ToString()))
          {
            spListItem[fieldName] = null;
          }
          else
          {
            string[] userIds = fieldValue["value"].ToString().Split(',');
            FieldUserValue[] fieldUserValues = new FieldUserValue[userIds.Length];
            for (int i = 0; i < userIds.Length; i++)
            {
              FieldUserValue userValue = new FieldUserValue();
              userValue.LookupId = Convert.ToInt32(userIds[i]);

              fieldUserValues[i] = userValue;
            }

            spListItem[fieldName] = fieldUserValues;
          }
          break;
        default:
          spListItem[fieldName] = fieldValue;
          break;
      }

    }

  }
}