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
        const int maxRetries = 5;
        int attempt = 0;

        while (true)
        {
            attempt++;
            var response = await ExecuteAsync(request);

            if (response.IsSuccessful)
                return response;

            bool isTooManyRequests = response.StatusCode == HttpStatusCode.TooManyRequests;
            bool isServerError = (int)response.StatusCode >= 500 && (int)response.StatusCode < 600;
            bool hasRetryableBody =
                !string.IsNullOrEmpty(response.Content) &&
                (response.Content.Contains("Internal Server Error", StringComparison.OrdinalIgnoreCase)
                 || response.Content.Contains("InternalServerError", StringComparison.OrdinalIgnoreCase)
                 || response.Content.Contains("UnknownError", StringComparison.OrdinalIgnoreCase));

            if ((isTooManyRequests || isServerError || hasRetryableBody) && attempt <= maxRetries)
            {
                var retryAfterHeader = response.Headers.FirstOrDefault(h => h.Name.Equals("Retry-After", StringComparison.OrdinalIgnoreCase));
                if (retryAfterHeader != null && int.TryParse(retryAfterHeader.Value?.ToString(), out int seconds))
                {
                    await Task.Delay((seconds + 1) * 1000);
                }
                else
                {
                    await Task.Delay((int)(Math.Pow(2, attempt) * 1000));
                }
                continue;
            }
            throw ConfigureErrorException(response);
        }
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