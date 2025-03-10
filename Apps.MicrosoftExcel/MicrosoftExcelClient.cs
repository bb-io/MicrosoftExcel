using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;
using System.Linq;
using System.Threading;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelClient : RestClient
{
    public MicrosoftExcelClient() 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl() ,
            MaxTimeout = 200000
        }) { }

    private static Uri GetBaseUrl()
    {
        return new Uri("https://graph.microsoft.com/v1.0"); // me/drive or sites/{siteId}
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
        if (error.Error.Code?.Equals("InternalServerError", StringComparison.OrdinalIgnoreCase) == true)
        {
            return new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
        }
        return new PluginApplicationException($"{error.Error.Code} - {error.Error.Message}");
    }
}