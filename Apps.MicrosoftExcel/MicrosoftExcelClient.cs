using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Utils.RestSharp;
using RestSharp;
using System.Linq;
using System.Text.RegularExpressions;
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
        var response = new RestResponse();
        try
        {
            response = await ExecuteAsync(request);

            if (response.StatusCode == System.Net.HttpStatusCode.TooManyRequests ||
                (!string.IsNullOrEmpty(response.Content) && (response.Content.Contains("Internal Server Error", StringComparison.OrdinalIgnoreCase) || response.Content.Contains("UnknownError", StringComparison.OrdinalIgnoreCase))))
            {
                var retryAfterHeader = response.Headers.FirstOrDefault(x => x.Name == "Retry-After");
                if (retryAfterHeader != null && !string.IsNullOrEmpty(retryAfterHeader.Value?.ToString()))
                {
                    int timeout = int.Parse(retryAfterHeader.Value.ToString());
                    await Task.Delay((timeout + 1) * 1000);
                }
                else
                {
                    await Task.Delay(3000);
                }
                return await ExecuteWithHandling(request);
            }
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase))
            {
                throw new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
            }
            else
            {
                throw new PluginApplicationException(ex.Message);
            }   
        }


        if (response.IsSuccessful)
            return response;

        throw ConfigureErrorException(response);
    }

    private Exception ConfigureErrorException(RestResponse response)
    {
        var content = response.Content ?? string.Empty;
        var contentType = response.Headers
               .FirstOrDefault(h => string.Equals(h.Name, "Content-Type", StringComparison.OrdinalIgnoreCase))
               ?.Value?
               .ToString() ?? string.Empty;

        if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase)
        || content.TrimStart().StartsWith("<"))
        {
            var plainText = Regex.Replace(content, "<.*?>", string.Empty).Trim();
            return new PluginApplicationException($"HTTP {(int)response.StatusCode} — {plainText}");
        }

        var error = response?.Content?.DeserializeResponseContent<ErrorDto>();
        if ((error?.Error.Message?.Contains("Internal Server Error", StringComparison.OrdinalIgnoreCase) ?? false) || (error?.Error.Message?.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
        }
        return new PluginApplicationException($"{error?.Error.Code} - {error?.Error.Message}");
    }
}