using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
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
        var oneDrivePrefix = "/me";
        var filesDictionary = await GetFilesFromService(oneDrivePrefix, context.SearchString);

        if (!string.IsNullOrEmpty(WorkbookRequest?.SiteName))
        {
            var siteId = MicrosoftExcelRequest.GetSiteId(InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value, WorkbookRequest.SiteName);
            var sharePointPrefix = $"/sites/{siteId}";
            var sharePointWorkbooks = await GetFilesFromService(sharePointPrefix, context.SearchString);
            foreach (var sharePointWorkbook in sharePointWorkbooks)
            {
                filesDictionary.Add(sharePointWorkbook.Key, sharePointWorkbook.Value);
            }
        }
        return filesDictionary;
    }

    private async Task<Dictionary<string, string>> GetFilesFromService(string servicePrefix, string searchString)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = $"{servicePrefix}/drive/root/search(q='.xls')?$select=name,id&$top=100"; //$filter=fields/ContentType eq 'Document'&
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
                .Where(i => i.Name.Contains(searchString, StringComparison.OrdinalIgnoreCase));

            foreach (var file in filteredFiles)
                filesDictionary.Add(file.Id, file.Name);

            filesAmount += filteredFiles.Count();
            var nextUrl = files.ODataNextLink?.Split(servicePrefix)[1];
            endpoint = nextUrl != null ? $"{servicePrefix}{nextUrl}" : null;
        } while (filesAmount < 20 && endpoint != null);
        return filesDictionary;
    }
}