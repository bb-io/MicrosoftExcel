using Apps.MicrosoftExcel.Utils.Drive;
using RestSharp;
using RestSharp.Interceptors;

namespace Apps.MicrosoftExcel.Api.Interceptors;

public class DrivePathInterceptor : Interceptor
{
    public override async ValueTask BeforeRequest(RestRequest request, CancellationToken cancellationToken)
    {
        if (request is not MicrosoftExcelRequest excelRequest)
            return;

        string drivePath = await DrivePathResolver.GetDrivePath(excelRequest.WorkbookRequest, excelRequest.AuthHeader);
        excelRequest.Resource = $"{drivePath}/{excelRequest.Endpoint.TrimStart('/')}";
    }
}