using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;

namespace Apps.MicrosoftExcel.Actions
{
    [ActionList]
    public class TableActions : MicrosoftExcelInvocable
    {
        public TableActions(InvocationContext invocationContext) : base(invocationContext)
        {
        }

        [Action("List table rows", Description = "List table rows")]
        public async Task<SimplerRowsDto> ListTableRows(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest)
        {
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows",
                Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);

            var result = await Client.ExecuteWithHandling<ListWrapper<MultipleListWrapper<List<string>>>>(request);
            var allRows = result.Value.ToList();
            return new SimplerRowsDto() { Rows = allRows.Select(x => x.Values.First()).ToList() };
        }

        [Action("Get table row", Description = "Get table row")]
        public async Task<RowDto> GetTableRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] GetTableRowRequest tableRowRequest)
        {
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/itemAt(index={tableRowRequest.RowIndex})",
                Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);

            var result = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }

        [Action("Update table row", Description = "Update table row")]
        public async Task<RowDto> UpdateTableRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] GetTableRowRequest tableRowRequest,
        [ActionParameter] UpdateRowRequest updateTableRowRequest)
        {
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/itemAt(index={tableRowRequest.RowIndex})",
                Method.Patch, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
            request.AddJsonBody(new
            {
                values = new[]
                {
                    updateTableRowRequest.Row
                }
            });
            var result = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }

        [Action("Create table", Description = "Create table")]
        public async Task<TableDto> CreateTable(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] CreateTableRequest createTableRequest)
        {
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/tables/add",
                Method.Post, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
            request.AddJsonBody(new
            {
                address = $"{createTableRequest.ColumnRow1}:{createTableRequest.ColumnRow2}",
                hasHeaders = createTableRequest.HasHeaders ?? true
            });
            return await Client.ExecuteWithHandling<TableDto>(request);
        }

        [Action("Add new table row", Description = "Add new table row")]
        public async Task<RowDto> AddRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] TableRowOptionalRequest rowRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
        {
            var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/add",
                Method.Post, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
            request.AddJsonBody(new
            {
                index = rowRequest.RowIndex,
                values = new[] { updateRowRequest.Row }
            });
            var result = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }
    }
}
