using Apps.MicrosoftExcel.Utils.Drive;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Authentication;

namespace Apps.MicrosoftExcel.Utils;

public static class PrefixResolver
{
    public static async Task<string> ResolvePrefix(string? siteName, IEnumerable<AuthenticationCredentialsProvider> creds)
    {
        if (string.IsNullOrEmpty(siteName)) 
            return "/me";
        
        string token = creds.First(p => p.KeyName == "Authorization").Value;
        string siteId = await DrivePathResolver.GetSiteId(token, siteName) ?? 
                        throw new PluginMisconfigurationException($"'{siteName}' site was not found");

        return $"/sites/{siteId}";
    }
}
