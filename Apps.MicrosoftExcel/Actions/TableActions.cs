using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using RestSharp;
using System.Xml.Linq;

namespace Apps.MicrosoftExcel.Actions
{
    [ActionList]
    public class TableActions : BaseInvocable
    {
        public TableActions(InvocationContext invocationContext) : base(invocationContext)
        {
        }

        [Action("List table rows", Description = "List table rows")]
        public async Task<RowsDto> ListTableRows(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest)
        {
            var client = new MicrosoftExcelClient();
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows",
                Method.Get, InvocationContext.AuthenticationCredentialsProviders);

            var result = await client.ExecuteWithHandling<ListWrapper<MultipleListWrapper<List<string>>>>(request);
            return new RowsDto() { Rows = result.Value.Select(x => x.Values.First()).ToList() };
        }

        [Action("Get table row", Description = "Get table row")]
        public async Task<RowDto> GetTableRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] GetTableRowRequest tableRowRequest)
        {
            var client = new MicrosoftExcelClient();
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/itemAt(index={tableRowRequest.RowIndex})",
                Method.Get, InvocationContext.AuthenticationCredentialsProviders);

            var result = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }

        [Action("Update table row", Description = "Update table row")]
        public async Task<RowDto> UpdateTableRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] GetTableRowRequest tableRowRequest,
        [ActionParameter] UpdateRowRequest updateTableRowRequest)
        {
            var client = new MicrosoftExcelClient();
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/itemAt(index={tableRowRequest.RowIndex})",
                Method.Patch, InvocationContext.AuthenticationCredentialsProviders);
            request.AddJsonBody(new
            {
                values = new[]
                {
                    updateTableRowRequest.Row
                }
            });
            var result = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }

        [Action("Create table", Description = "Create table")]
        public async Task<TableDto> CreateTable(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] CreateTableRequest createTableRequest)
        {
            var client = new MicrosoftExcelClient();
            var request = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/tables/add",
                Method.Post, InvocationContext.AuthenticationCredentialsProviders);
            request.AddJsonBody(new
            {
                address = $"{createTableRequest.ColumnRow1}:{createTableRequest.ColumnRow2}",
                hasHeaders = createTableRequest.HasHeaders ?? true
            });
            return await client.ExecuteWithHandling<TableDto>(request);
        }

        [Action("Add new table row", Description = "Add new table row")]
        public async Task<RowDto> AddRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] TableRequest tableRequest,
        [ActionParameter] TableRowOptionalRequest rowRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
        {
            var client = new MicrosoftExcelClient();
            var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/tables/{tableRequest.Table}/rows/add",
                Method.Post, InvocationContext.AuthenticationCredentialsProviders);
            request.AddJsonBody(new
            {
                index = rowRequest.RowIndex,
                values = new[] { updateRowRequest.Row }
            });
            var result = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
            return new RowDto() { Row = result.Values.First() };
        }
    }
}
