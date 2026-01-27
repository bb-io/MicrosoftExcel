using Apps.MicrosoftExcel;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.Polling.Models;
using Apps.MicrosoftExcel.Utils;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Polling;
using RestSharp;

public class ExcelPollingList(InvocationContext invocationContext)
    : BaseInvocable(invocationContext)
{
    [PollingEvent("On workbook updated", Description = "Triggered when a workbook is modified")]
    public async Task<PollingEventResponse<WorkbookUpdatedMemory, WorkbookUpdatedResult>>
        OnWorkbookUpdated(
            PollingEventRequest<WorkbookUpdatedMemory> request,
            [PollingEventParameter] WorkbookRequest workbookRequest)
    {
        var token = InvocationContext.AuthenticationCredentialsProviders
            .First(p => p.KeyName == "Authorization").Value;

        var client = new MicrosoftExcelClient();
        var prefix = ResolvePrefix(workbookRequest);

        var metadataRequest = new RestRequest(
            $"{prefix}/drive/items/{workbookRequest.WorkbookId}?$select=id,name,lastModifiedDateTime",
            Method.Get);

        metadataRequest.AddHeader("Authorization", token);

        var workbook = await ErrorHandler.ExecuteWithErrorHandlingAsync(
            () => client.ExecuteWithHandling<FileMetadataDto>(metadataRequest)
        );

        if (request.Memory == null)
        {
            var initialMemory = new WorkbookUpdatedMemory
            {
                LastModifiedDateTime = workbook.LastModifiedDateTime,
                LastPollingTime = DateTime.UtcNow,
                Triggered = false
            };

            return new PollingEventResponse<WorkbookUpdatedMemory, WorkbookUpdatedResult>
            {
                FlyBird = false,
                Memory = initialMemory,
                Result = null
            };
        }

        var memory = request.Memory;

        bool hasChanged =
            workbook.LastModifiedDateTime > memory.LastModifiedDateTime;

        memory.LastPollingTime = DateTime.UtcNow;
        memory.Triggered = hasChanged;
        memory.LastModifiedDateTime = workbook.LastModifiedDateTime;

        WorkbookUpdatedResult? result = null;

        if (hasChanged)
        {
            result = new WorkbookUpdatedResult
            {
                WorkbookId = workbook.Id,
                WorkbookName = workbook.Name,
                LastModifiedDateTime = workbook.LastModifiedDateTime!.Value
            };
        }

        return new PollingEventResponse<WorkbookUpdatedMemory, WorkbookUpdatedResult>
        {
            FlyBird = hasChanged,
            Memory = memory,
            Result = result
        };
    }

    private string ResolvePrefix(WorkbookRequest request)
    {
        if (!string.IsNullOrEmpty(request.SiteName))
        {
            var token = InvocationContext.AuthenticationCredentialsProviders
                .First(p => p.KeyName == "Authorization").Value;

            var siteId = MicrosoftExcelRequest.GetSiteId(token, request.SiteName)
                ?? throw new PluginMisconfigurationException(
                    $"'{request.SiteName}' site was not found");

            return $"/sites/{siteId}";
        }

        return "/me";
    }
}

