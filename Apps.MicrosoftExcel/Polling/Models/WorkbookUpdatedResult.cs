using Blackbird.Applications.Sdk.Common;

public class WorkbookUpdatedResult
{
    [Display("Workbook ID")]
    public string WorkbookId { get; set; } 

    [Display("Workbook name")]
    public string WorkbookName { get; set; }

    [Display("Last modified date")]
    public DateTime LastModifiedDateTime { get; set; }
}