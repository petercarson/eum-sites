namespace EITSiteProvisioning_API.Models
{
  public class SiteListItem
  {
    public int Id { get; set; }
    public string Title { get; set; }
    public string EUMSiteURL { get; set; }
    public string EUMParentURL { get; set; }

    public ILookupFieldValue EUMDivision { get; set; }
    public string EUMGroupSummary { get; set; }
    public string EUMAlias { get; set; }
    public string EUMSiteVisibility { get; set; }
    public string SitePurpose { get; set; }
    public string EUMSiteCreated { get; set; }
    public ILookupFieldValue EUMSiteTemplate { get; set; }
  }

  public class ILookupFieldValue
  {
    public int Id { get; set; }
    public string Title { get; set; }
  }
}