using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Utils.RestSharp;
using RestSharp;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelClient : RestClient
{
    private const int MaxRetries = 5;
    private const int InitialDelayMs = 1000;
    public MicrosoftExcelClient() 
        : base(new RestClientOptions
        {
            ThrowOnAnyError = false, BaseUrl = GetBaseUrl() ,
            Timeout = TimeSpan.FromMilliseconds(200000)
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
        int delay = InitialDelayMs;
        RestResponse? response = null;

        for (int attempt = 1; attempt <= MaxRetries; attempt++)
        {
            response = await ExecuteAsync(request);

            if (response.IsSuccessful)
                return response;

            if (attempt < MaxRetries &&
                (response.StatusCode == HttpStatusCode.InternalServerError ||
                 response.StatusCode == HttpStatusCode.ServiceUnavailable))
            {
                await Task.Delay(delay);
                delay *= 2;
                continue;
            }
            break;
        }

        throw ConfigureErrorException(response);
    }

    private Exception ConfigureErrorException(RestResponse response)
    {
        if (string.IsNullOrEmpty(response.Content))
        {
            if(string.IsNullOrEmpty(response.ErrorMessage))
            {
                return new PluginApplicationException($"HTTP {(int)response.StatusCode} — {response.StatusDescription}");
            }
            
            return new PluginApplicationException(response.ErrorMessage);
        }
        
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
        if (response!.StatusCode == HttpStatusCode.InternalServerError || (error?.Error.Message?.Contains("Internal Server Error", StringComparison.OrdinalIgnoreCase) ?? false) || (error?.Error.Message?.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return new PluginApplicationException("An internal server error occurred. Please implement a retry policy and try again.");
        }

        if (response!.StatusCode == HttpStatusCode.ServiceUnavailable || (error?.Error.Message?.Contains("Service Unavailable", StringComparison.OrdinalIgnoreCase) ?? false) || (error?.Error.Message?.Contains("ServiceUnavailable", StringComparison.OrdinalIgnoreCase) ?? false))
        {
            return new PluginApplicationException("Server service unavailable error occurred. Please implement a retry policy and try again.");
        }

        return new PluginApplicationException($"{error?.Error.Code} - {error?.Error.Message}");
    }
}