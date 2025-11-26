using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftExcel.Models.Requests;

public class WorkbookRequest
{
    [Display("Workbook", Description = "Your Excel file")]
    [FileDataSource(typeof(WorkbookFileDataSourceHandler))]
    public string WorkbookId { get; set; }

    [Display("Sharepoint site name", Description = "Sharepoint site name")]
    [DataSource(typeof(SiteDataSourceHandler))]
    public string? SiteName { get; set; }
}
