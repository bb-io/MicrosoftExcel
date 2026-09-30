using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common.Authentication;
using RestSharp;

namespace Apps.MicrosoftExcel;

public class MicrosoftExcelRequest : RestRequest
{
    public string Endpoint { get; }
    public string AuthHeader { get; }
    public WorkbookRequest WorkbookRequest { get; }

    public MicrosoftExcelRequest(
        string endpoint,
        Method method,
        IEnumerable<AuthenticationCredentialsProvider> creds,
        WorkbookRequest workbookRequest) 
        : base(endpoint, method)
    {
        Endpoint = endpoint;
        WorkbookRequest = workbookRequest;
        AuthHeader = creds.First(p => p.KeyName == "Authorization").Value;
        this.AddHeader("Authorization", AuthHeader);
    }
}