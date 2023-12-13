using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using RestSharp;
using System.Linq;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelClient : RestClient
{
    public MicrosoftExcelClient() 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl() 
        }) { }

    private static Uri GetBaseUrl()
    {
        return new Uri("https://graph.microsoft.com/v1.0/me/drive");
    }
    
    public async Task<T> ExecuteWithHandling<T>(RestRequest request)
    {
        var response = await ExecuteWithHandling(request);
        return response.Content.DeserializeResponseContent<T>();
    }
    
    public async Task<RestResponse> ExecuteWithHandling(RestRequest request)
    {
        var response = await ExecuteAsync(request);

        if(response.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
        {
            int timeout = int.Parse(response.Headers.Where(x => x.Name == "Retry-After").First().Value.ToString());
            await Task.Delay(timeout * 1000);
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
}