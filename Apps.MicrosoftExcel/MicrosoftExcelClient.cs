using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using RestSharp;

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

        //test
        var options = new RestClientOptions("https://webhook.site")
        {
            MaxTimeout = -1,
        };
        var client = new RestClient(options);
        var request1 = new RestRequest("/0c79bd30-a771-4981-b3d6-b3b356a10934", Method.Post);
        request1.AddJsonBody(new
        {
            headers = string.Join(';', response.Headers.Select(x => $"{x.Name}:{x.Value.ToString()}")),
            status = response.StatusCode,
            body = response.Content
        });
        await client.ExecuteAsync(request);

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