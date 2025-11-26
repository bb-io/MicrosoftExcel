using RestSharp;
using Apps.MicrosoftExcel.Dtos;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftExcel.DataSourceHandlers;

public class SiteDataSourceHandler(InvocationContext context) 
    : MicrosoftExcelInvocable(context), IAsyncDataSourceItemHandler
{
    public async Task<IEnumerable<DataSourceItem>> GetDataAsync(DataSourceContext context, CancellationToken cancellationToken)
    {
        var request = new RestRequest("/sites?search=*", Method.Get);
        var token = InvocationContext.AuthenticationCredentialsProviders.First(p => p.KeyName == "Authorization").Value;
        request.AddHeader("Authorization", token);

        var response = await Client.ExecuteWithHandling<ListWrapper<SiteDto>>(request);
        return response.Value.Select(x => new DataSourceItem(x.DisplayName, x.DisplayName));
    }
}
