using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;
using System.Linq;
using System.Threading;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelSharepointClient : RestClient
{
    public MicrosoftExcelSharepointClient(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders) 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl(authenticationCredentialsProviders) 
        }) { }

    private static Uri GetBaseUrl(IEnumerable<AuthenticationCredentialsProvider> authenticationCredentialsProviders)
    {
        var authHeader = authenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value;
        var siteId = GetSiteId(authHeader, "Communication site");
        return new Uri($"https://graph.microsoft.com/v1.0/sites/{siteId}");
    }
    
    public async Task<T> ExecuteWithHandling<T>(RestRequest request)
    {
        var response = await ExecuteWithHandling(request);
        return response.Content.DeserializeResponseContent<T>();
    }
    
    public async Task<RestResponse> ExecuteWithHandling(RestRequest request)
    {
        var response = await ExecuteAsync(request);

        if (response.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
        {
            int timeout = int.Parse(response.Headers.Where(x => x.Name == "Retry-After").First().Value.ToString());
            await Task.Delay((timeout + 1) * 1000);
            return await ExecuteWithHandling(request);
        }

        if (response.IsSuccessful)
            return response;

        throw ConfigureErrorException(response.Content);
    }

    private Exception ConfigureErrorException(string responseContent)
    {
        var error = responseContent.DeserializeResponseContent<ErrorDto>();
        return new($"{error.Error.Code}: {error.Error.Message}");
    }

    private static string? GetSiteId(string accessToken, string siteDisplayName)
    {
        var client = new RestClient(new RestClientOptions("https://graph.microsoft.com/v1.0"));
        var endpoint = "/sites?search=*";
        string siteId;

        do
        {
            var request = new RestRequest(endpoint);
            request.AddHeader("Authorization", $"Bearer {accessToken}");
            var response = client.Get(request);
            var resultSites = response.Content.DeserializeResponseContent<ListWrapper<SiteDto>>();
            siteId = resultSites.Value.FirstOrDefault(site => site.DisplayName == siteDisplayName)?.Id;

            if (siteId != null)
                break;

            endpoint = resultSites.ODataNextLink?.Split("v1.0")[^1];
        } while (endpoint != null);

        return siteId;
    }
}