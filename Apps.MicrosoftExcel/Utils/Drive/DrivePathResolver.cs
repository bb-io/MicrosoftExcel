using System.Collections.Concurrent;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common.Exceptions;
using RestSharp;

namespace Apps.MicrosoftExcel.Utils.Drive;

public static class DrivePathResolver
{
    private static readonly ConcurrentDictionary<DriveCacheKey, string> Cache = new();
    
    public static async Task<string> GetDrivePath(WorkbookRequest workbookRequest, string authHeader)
    {
        var cacheKey = new DriveCacheKey(workbookRequest.WorkbookId, workbookRequest.SiteName);
        if (Cache.TryGetValue(cacheKey, out string? drivePath))
            return drivePath;
        
        bool isOneDriveWorkbook = await IsOneDriveWorkbook(workbookRequest.WorkbookId, authHeader);
        if (isOneDriveWorkbook)
            drivePath = "/me/drive";
        else
        {
            string siteId = await GetSiteId(authHeader, workbookRequest.SiteName) ?? 
                            throw new PluginMisconfigurationException($"'{workbookRequest.SiteName}' site was not found");

            drivePath = $"/sites/{siteId}/drive";
        }
        
        Cache.TryAdd(cacheKey, drivePath);
        return drivePath;
    }

    public static async Task<string?> GetSiteId(string authHeader, string? siteName)
    {
        if (string.IsNullOrWhiteSpace(siteName))
            return null;

        return await FindSite(authHeader, siteName, Uri.EscapeDataString(siteName)) ?? 
               await FindSite(authHeader, siteName, "*");
    }
    
    private static async Task<string?> FindSite(string authHeader, string siteName, string searchQuery)
    {
        var client = new MicrosoftExcelClient();
        var endpoint = $"/sites?search={searchQuery}";

        while (endpoint != null)
        {
            var request = new RestRequest(endpoint).AddHeader("Authorization", authHeader);
            var sites = await client.ExecuteWithHandling<ListWrapper<SiteDto>>(request);

            string? siteId = sites.Value.FirstOrDefault(s => s.DisplayName == siteName || s.WebUrl == siteName)?.Id;
            if (!string.IsNullOrEmpty(siteId))
                return siteId;

            endpoint = sites.ODataNextLink?.Split("v1.0")[^1];
        }

        return null;
    }

    private static async Task<bool> IsOneDriveWorkbook(string workbookId, string authHeader)
    {
        var client = new MicrosoftExcelClient();
        var request = new RestRequest($"/me/drive/items/{workbookId}/workbook/worksheets")
            .AddHeader("Authorization", authHeader)
            .AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

        var response = await client.ExecuteAsync(request);
        return response.IsSuccessStatusCode;
    }
}