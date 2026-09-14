using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Blackbird.Applications.Sdk.Common.Exceptions;
using RestSharp;
using System.Net;
using System.Text.RegularExpressions;
using Polly;
using Polly.Retry;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelClient() : RestClient(new RestClientOptions
{
    ThrowOnAnyError = false,
    BaseUrl = new Uri("https://graph.microsoft.com/v1.0"),
    Timeout = TimeSpan.FromMilliseconds(200000),
})
{
    private static readonly ResiliencePipeline<RestResponse> RetryPipeline = new ResiliencePipelineBuilder<RestResponse>()
        .AddRetry(new RetryStrategyOptions<RestResponse>
        {
            MaxRetryAttempts = 5,
            UseJitter = true,
            Delay = TimeSpan.FromSeconds(1),
            BackoffType = DelayBackoffType.Exponential,
            ShouldHandle = new PredicateBuilder<RestResponse>().HandleResult(x => x.IsTransient()),
            DelayGenerator = args => new ValueTask<TimeSpan?>(args.Outcome.Result.GetRetryAfter())
        })
        .Build();
    
    public async Task<T> ExecuteWithHandling<T>(RestRequest request)
    {
        var response = await ExecuteWithHandling(request);
        return response.Content.DeserializeResponseContent<T>();
    }

    public async Task<RestResponse> ExecuteWithHandling(RestRequest request)
    {
        var response = await RetryPipeline.ExecuteAsync(async ct => await ExecuteAsync(request, ct), CancellationToken.None);
        return response.IsSuccessful ? response : throw ConfigureErrorException(response);
    }

    private static PluginApplicationException ConfigureErrorException(RestResponse response)
    {
        if (string.IsNullOrEmpty(response.Content))
        {
            if (string.IsNullOrEmpty(response.ErrorMessage))
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

        var error = response.Content?.DeserializeResponseContent<ErrorDto>();
        var errorMessages = new[] { "Internal Server Error", "InternalServerError", "Service Unavailable", "ServiceUnavailable" };
        if (response.StatusCode is HttpStatusCode.InternalServerError or HttpStatusCode.ServiceUnavailable ||
            errorMessages.Contains(error?.Error.Message, StringComparer.OrdinalIgnoreCase))
        {
            return new PluginApplicationException(
                "Microsoft Graph is temporarily unavailable. Retries were exhausted. Please try again later");
        }

        return new PluginApplicationException($"{error?.Error.Code} - {error?.Error.Message}");
    }
}