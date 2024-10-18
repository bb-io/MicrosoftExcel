using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common.Authentication;
using RestSharp;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelRequest : RestRequest
{
    public MicrosoftExcelRequest(string endpoint, Method method,
        IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders, WorkbookRequest workbookRequest) : 
        base(GetEndpointWithServicePath(endpoint, workbookRequest, authenticationCredentialsProviders), method)
    {
        this.AddHeader("Authorization", authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value);
    }

    private static string GetEndpointWithServicePath(string endpoint, WorkbookRequest workbookRequest, 
        IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
    {
        var authHeader = authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value;
        if (IsOneDriveWorkbook(workbookRequest.WorkbookId, authHeader))
            return "/me/drive/" + endpoint.TrimStart('/');

        return $"/sites/{GetSiteId(authHeader, workbookRequest.SiteName)}/drive/{endpoint.TrimStart('/')}";
    }

    public static string? GetSiteId(string accessToken, string siteDisplayName)
    {
        var client = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var endpoint = "/sites?search=*";
        string siteId;

        do
        {
            var request = new RestRequest(endpoint);
            request.AddHeader("Authorization", $"{accessToken}");
            var response = client.Get(request);
            var resultSites = response.Content.DeserializeResponseContent<ListWrapper<SiteDto>>();
            siteId = resultSites.Value.FirstOrDefault(site => site.DisplayName == siteDisplayName)?.Id;

            if (siteId != null)
                break;

            endpoint = resultSites.ODataNextLink?.Split("v1.0")[^1];
        } while (endpoint != null);

        return siteId;
    }

    private static bool IsOneDriveWorkbook(string workbookId, string authHeader)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = $"/me/drive/items/{workbookId}/workbook/worksheets";
        var worksheetsDictionary = new Dictionary<string, string>();

        var request = new RestRequest(endpoint, Method.Get);
        request.AddHeader("Authorization", authHeader);
        request.AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
        var response = client.Execute(request);
        return response.IsSuccessStatusCode;
    }
}