using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using DocumentFormat.OpenXml.InkML;
using RestSharp;

namespace Apps.MicrosoftExcel.DataSourceHandlers;

public class WorkbookDataSourceHandler : BaseInvocable, IAsyncDataSourceHandler
{
    private WorkbookRequest WorkbookRequest { get; set; }

    public WorkbookDataSourceHandler(InvocationContext invocationContext, 
        [ActionParameter] WorkbookRequest workbookRequest) : base(invocationContext)
    {
        WorkbookRequest = workbookRequest;
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = "/me/drive/root/search(q='.xls')?$select=name,id&$top=100"; //$filter=fields/ContentType eq 'Document'&
        var filesDictionary = new Dictionary<string, string>();
        var filesAmount = 0;

        do
        {
            var request = new RestRequest(endpoint, Method.Get);
            request.AddHeader("Authorization", InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value);
            request.AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var files = await client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);
            var filteredFiles = files.Value
                .Select(i => new { i.Id, i.Name })
                .Where(i => i.Name.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase));
            
            foreach (var file in filteredFiles)
                filesDictionary.Add(file.Id, file.Name);
            
            filesAmount += filteredFiles.Count();
            endpoint = $"/me/drive{files.ODataNextLink?.Split("me/drive")[1]}";
        } while (filesAmount < 20 && endpoint != null);

        //foreach (var file in filesDictionary)
        //{
        //    var filePath = file.Value;
        //    if (filePath.Length > 40)
        //    {
        //        var filePathParts = filePath.Split("/");
        //        if (filePathParts.Length > 3)
        //        {
        //            filePath = string.Join("/", filePathParts[0], "...", filePathParts[^2], filePathParts[^1]);
        //            filesDictionary[file.Key] = filePath;
        //        }
        //    }
        //}

        if (!string.IsNullOrEmpty(WorkbookRequest?.SiteName))
        {
            var sharePointWorkbooks = await GetSharePointWorkbooks(context.SearchString, WorkbookRequest.SiteName);
            foreach(var sharePointWorkbook in sharePointWorkbooks)
            {
                filesDictionary.Add(sharePointWorkbook.Key, sharePointWorkbook.Value);
            }
        }

        return filesDictionary;
    }

    private async Task<Dictionary<string, string>> GetSharePointWorkbooks(string searchString, string siteId)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = $"/sites/{siteId}/drive/list/items?$select=id&$expand=driveItem($select=id,name,parentReference)&" +
                       "$filter=fields/ContentType eq 'Document'&$top=20";
        var filesDictionary = new Dictionary<string, string>();
        var filesAmount = 0;

        do
        {
            var request = new RestRequest(endpoint, Method.Get);
            request.AddHeader("Authorization", InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value);
            request.AddHeader("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var files = await client.ExecuteWithHandling<ListWrapper<DriveItemWrapper<FileMetadataDto>>>(request);
            var filteredFiles = files.Value
                .Select(w => w.DriveItem)
                .Select(i => new { i.Id, Path = GetFilePath(i) })
                .Where(i => i.Path.Contains(searchString, StringComparison.OrdinalIgnoreCase));

            foreach (var file in filteredFiles)
                filesDictionary.Add(file.Id, file.Path);

            filesAmount += filteredFiles.Count();
            endpoint = files.ODataNextLink == null ? null : $"/sites/{siteId}/drive" + files.ODataNextLink?.Split("drive")[^1];
        } while (filesAmount < 20 && endpoint != null);

        foreach (var file in filesDictionary)
        {
            var filePath = file.Value;
            if (filePath.Length > 40)
            {
                var filePathParts = filePath.Split("/");
                if (filePathParts.Length > 3)
                {
                    filePath = string.Join("/", filePathParts[0], "...", filePathParts[^2], filePathParts[^1]);
                    filesDictionary[file.Key] = filePath;
                }
            }
        }
        return filesDictionary;
    }

    private string GetFilePath(FileMetadataDto file)
    {
        var parentPath = file.ParentReference.Path.Split("root:");
        if (parentPath[1] == "")
            return file.Name;

        return $"{parentPath[1].Substring(1)}/{file.Name}";
    }
}