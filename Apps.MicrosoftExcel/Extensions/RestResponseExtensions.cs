using System.Net;
using Apps.MicrosoftExcel.Dtos;
using RestSharp;

namespace Apps.MicrosoftExcel.Extensions;

public static class RestResponseExtensions
{
    public static bool IsTransient(this RestResponse response)
    {
        return response.StatusCode 
                   is HttpStatusCode.InternalServerError
                   or HttpStatusCode.ServiceUnavailable
                   or HttpStatusCode.TooManyRequests
                   or HttpStatusCode.GatewayTimeout
               || response.ResponseStatus is ResponseStatus.TimedOut or ResponseStatus.Error
               || response.IsMaxRequestDurationExceeded();
    }
    
    public static bool IsMaxRequestDurationExceeded(this RestResponse response)
    {
        if (string.IsNullOrEmpty(response.Content)) 
            return false;

        try
        {
            var error = response.Content.DeserializeResponseContent<ErrorDto>();
            return string.Equals(error?.Error?.Code, "MaxRequestDurationExceeded", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }
    
    public static TimeSpan? GetRetryAfter(this RestResponse? response)
    {
        var value = response?.Headers?
            .FirstOrDefault(h => string.Equals(h.Name, "Retry-After", StringComparison.OrdinalIgnoreCase))?
            .Value
            .ToString();

        if (!int.TryParse(value, out var seconds))
            return null;

        return TimeSpan.FromSeconds(Math.Min(seconds, 60));
    }
}