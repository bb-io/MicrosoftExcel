using RestSharp;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Utils;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using File = Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.File;

namespace Apps.MicrosoftExcel.DataSourceHandlers.Base;

public class BaseWorkbookFolderPicker(InvocationContext invocationContext, string? sitename) 
    : MicrosoftExcelInvocable(invocationContext)
{
    private const string RootFolderName = "Root";

    public async Task<IEnumerable<FileDataItem>> GetFolderContentAsync(
        FolderContentDataSourceContext context,
        bool workbooksAreSelectable,
        bool foldersAreSelectable)
    {
        var client = new MicrosoftExcelClient();
        var token = InvocationContext.AuthenticationCredentialsProviders
            .First(p => p.KeyName == "Authorization").Value;

        string prefix = await PrefixResolver.ResolvePrefix(sitename, InvocationContext.AuthenticationCredentialsProviders);
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
                        IsSelectable = foldersAreSelectable
                    });

                    continue;
                }
                else
                {
                    items.Add(new File
                    {
                        Id = i.Id,
                        DisplayName = i.Name,
                        Date = i.LastModifiedDateTime,
                        Size = i.Size,
                        IsSelectable = workbooksAreSelectable == true && i.Name.HasExcelExtension()
                    });
                }
            }

            endpoint = response.ODataNextLink;
        }

        return items;
    }

    public async Task<IEnumerable<FolderPathItem>> GetFolderPathAsync(FolderPathDataSourceContext context)
    {
        if (string.IsNullOrEmpty(context.FileDataItemId))
            return [];

        string prefix = await PrefixResolver.ResolvePrefix(sitename, InvocationContext.AuthenticationCredentialsProviders);

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
                    DisplayName = item.Name ?? RootFolderName
                });
            }

            currentId = item.ParentReference?.Id;
        }

        path.Reverse();

        if (path.Count > 0)
            path[0].DisplayName = RootFolderName;

        return path;
    }
}
