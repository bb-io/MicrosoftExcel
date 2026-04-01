using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftExcel.Models.Requests;

public class CreateWorkbookRequest
{
    [Display("Name", Description = "Workbook file name. .xlsx will be added automatically if missing.")]
    public string Name { get; set; } = default!;

    [Display("Parent folder ID", Description = "Optional OneDrive folder item id. If empty - workbook is created in root.")]
    [FileDataSource(typeof(FolderDataSourceHandler))]
    public string? ParentFolderId { get; set; }

    [Display("Overwrite", Description = "If true - replaces existing file with same name. If false - will rename on conflict.")]
    public bool? Overwrite { get; set; }
}
