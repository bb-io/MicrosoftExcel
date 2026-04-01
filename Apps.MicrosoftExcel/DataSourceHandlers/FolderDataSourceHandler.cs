using Apps.MicrosoftExcel.DataSourceHandlers.Base;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Apps.MicrosoftExcel.DataSourceHandlers;

public class FolderDataSourceHandler(
    InvocationContext invocationContext,
    [ActionParameter] SiteNameOptionalRequest siteName)
    : BaseWorkbookFolderPicker(invocationContext, siteName.SiteName), IAsyncFileDataSourceItemHandler
{
    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(FolderContentDataSourceContext context, CancellationToken ct)
    {
        return await GetFolderContentAsync(context, false, true);
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context, CancellationToken ct)
    {
        var path = await GetFolderPathAsync(context);
        var listPath = path.ToList();
        listPath.RemoveAt(listPath.Count - 1);

        return listPath;
    }
}
