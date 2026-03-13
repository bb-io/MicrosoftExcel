using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.Utils;
using Blackbird.Applications.Sdk.Common.Authentication;
using RestSharp;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelRequest : RestRequest
{
    public MicrosoftExcelRequest(
        string endpoint, 
        Method method,
        IEnumerable<AuthenticationCredentialsProvider> creds, 
        WorkbookRequest workbookRequest)
        : base(
            GetEndpointWithServicePath(endpoint, workbookRequest, creds).GetAwaiter().GetResult(), 
            method)
    {
        this.AddHeader("Authorization", creds.First(p => p.KeyName == "Authorization").Value);
    }

    private static async Task<string> GetEndpointWithServicePath(
        string endpoint, 
        WorkbookRequest workbookRequest, 
        IEnumerable<AuthenticationCredentialsProvider> creds)
    {
        var authHeader = creds.First(p => p.KeyName == "Authorization").Value;
        if (await IsOneDriveWorkbook(workbookRequest.WorkbookId, authHeader))
            return "/me/drive/" + endpoint.TrimStart('/');

        string siteId = await GetSiteId(authHeader, workbookRequest.SiteName);
        return $"/sites/{siteId}/drive/{endpoint.TrimStart('/')}";
    }

    public static async Task<string> GetSiteId(string accessToken, string? siteDisplayName)
    {
        var client = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var endpoint = "/sites?search=*";
        string siteId;

        do
        {
            var request = new RestRequest(endpoint);
            request.AddHeader("Authorization", $"{accessToken}");
            var response = await ErrorHandler.ExecuteWithErrorHandlingAsync(() => client.GetAsync(request));
            var resultSites = response.Content?.DeserializeResponseContent<ListWrapper<SiteDto>>();

            siteId = 
                resultSites?.Value
                .FirstOrDefault(site => site.DisplayName == siteDisplayName || site.WebUrl == siteDisplayName)?.Id;
            if (siteId != null)
                break;

            endpoint = resultSites?.ODataNextLink?.Split("v1.0")[^1];
        } while (endpoint != null);

        return siteId;
    }

    private static async Task<bool> IsOneDriveWorkbook(string workbookId, string authHeader)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = $"/me/drive/items/{workbookId}/workbook/worksheets";
        var request = new RestRequest(endpoint, Method.Get)
            .AddHeader("Authorization", authHeader)
            .AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

        var response = await client.ExecuteAsync(request);
        return response.IsSuccessStatusCode;
    }
}