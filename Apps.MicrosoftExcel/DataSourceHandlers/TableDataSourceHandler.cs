using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common;
using RestSharp;

namespace Apps.MicrosoftExcel.DataSourceHandlers
{
    public class TableDataSourceHandler : BaseInvocable, IAsyncDataSourceHandler
    {
        private WorkbookRequest WorkbookRequest { get; set; }

        public TableDataSourceHandler(InvocationContext invocationContext,
            [ActionParameter] WorkbookRequest workbookRequest) : base(invocationContext)
        {
            WorkbookRequest = workbookRequest;
        }

        public async Task<Dictionary<string, string>> GetDataAsync(DataSourceContext context,
            CancellationToken cancellationToken)
        {
            if (WorkbookRequest == null || string.IsNullOrEmpty(WorkbookRequest.WorkbookId))
            {
                throw new ArgumentException("Please, select the workbook first");
            }
            var client = new MicrosoftExcelClient();
            var endpoint = $"/items/{WorkbookRequest.WorkbookId}/workbook/tables";
            var worksheetsDictionary = new Dictionary<string, string>();

            var request = new MicrosoftExcelRequest(endpoint, Method.Get,
                    InvocationContext.AuthenticationCredentialsProviders, WorkbookRequest);
            request.AddHeader("prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
            var files = await client.ExecuteWithHandling<ListWrapper<TableDto>>(request);
            var filteredFiles = files.Value
                .Select(i => new { i.Id, i.Name });

            foreach (var file in filteredFiles)
                worksheetsDictionary.Add(file.Id, file.Name);

            return worksheetsDictionary;
        }
    }
}
