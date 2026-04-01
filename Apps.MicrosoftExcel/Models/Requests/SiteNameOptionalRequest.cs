using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Apps.MicrosoftExcel.Models.Requests;

public class SiteNameOptionalRequest
{
    [Display("Sharepoint site name", Description = "Sharepoint site name")]
    [DataSource(typeof(SiteDataSourceHandler))]
    public string? SiteName { get; set; }
}
