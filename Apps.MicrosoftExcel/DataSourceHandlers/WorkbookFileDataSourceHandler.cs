using RestSharp;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Utils;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using File = Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File;

namespace Apps.MicrosoftExcel.DataSourceHandlers;

public class WorkbookFileDataSourceHandler(
    InvocationContext invocationContext,
    [ActionParameter] WorkbookRequest workbookRequest) :
    MicrosoftExcelInvocable(invocationContext), IAsyncFileDataSourceItemHandler
{
    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(
        FolderPathDataSourceContext context,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(context.FileDataItemId))
            return Enumerable.Empty<FolderPathItem>();

        var prefix = ResolvePrefix();
        var token = InvocationContext.AuthenticationCredentialsProviders
            .First(p => p.KeyName == "Authorization").Value;

        var path = new List<FolderPathItem>();
        var currentId = context.FileDataItemId;

        while (!string.IsNullOrEmpty(currentId))
        {
            var request = new RestRequest(
                $"{prefix}/drive/items/{currentId}?$select=id,name,parentReference,folder", 
                Method.Get
            );
            request.AddHeader("Authorization", token);

            var item = await Client.ExecuteWithHandling<FileMetadataDto>(request);

            if (item.Folder != null)
            {
                path.Add(new FolderPathItem
                {
                    Id = item.Id,
                    DisplayName = item.Name ?? "Root"
                });
            }

            currentId = item.ParentReference?.Id;
        }

        path.Reverse();
        return path;
    }

    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(
        FolderContentDataSourceContext context,
        CancellationToken cancellationToken)
    {
        var client = new MicrosoftExcelClient();
        var token = InvocationContext.AuthenticationCredentialsProviders
            .First(p => p.KeyName == "Authorization").Value;

        string prefix = ResolvePrefix();
        string folderId = !string.IsNullOrEmpty(context.FolderId) ? context.FolderId : "root";

        var items = new List<FileDataItem>();
        var endpoint = $"{prefix}/drive/items/{folderId}/children?$select=id,name,size,lastModifiedDateTime,folder&$top=200";

        while (endpoint != null)
        {
            var request = new RestRequest(endpoint, Method.Get);
            request.AddHeader("Authorization", token);
            request.AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

            var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(
                () => client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request)
            );

            foreach (var i in response.Value)
            {
                if (i.Folder != null)
                {
                    items.Add(new Folder
                    {
                        Id = i.Id,
                        DisplayName = i.Name,
                        Date = i.LastModifiedDateTime,
                        IsSelectable = false
                    });

                    continue;
                }

                if (IsExcelFile(i))
                {
                    items.Add(new File
                    {
                        Id = i.Id,
                        DisplayName = i.Name,
                        Date = i.LastModifiedDateTime,
                        Size = i.Size,
                        IsSelectable = true
                    });
                }
            }

            endpoint = response.ODataNextLink;
        }

        return items;
    }

    private string ResolvePrefix()
    {
        if (!string.IsNullOrEmpty(workbookRequest?.SiteName))
        {
            var token = InvocationContext.AuthenticationCredentialsProviders
                .First(p => p.KeyName == "Authorization").Value;

            var siteId = MicrosoftExcelRequest.GetSiteId(token, workbookRequest.SiteName) ??
                throw new PluginMisconfigurationException($"'{workbookRequest?.SiteName}' site was not found");
            return $"/sites/{siteId}";
        }

        return "/me";
    }

    private static bool IsExcelFile(FileMetadataDto file)
    {
        return file.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
               file.Name.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);
    }
}