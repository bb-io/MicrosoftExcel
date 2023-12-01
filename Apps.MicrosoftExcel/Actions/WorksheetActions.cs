using System.Net.Mime;
using Apps.MicrosoftExcel.DataSourceHandlers;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Newtonsoft.Json.Linq;
using RestSharp;
using File = Blackbird.Applications.Sdk.Common.Files.File;

namespace Apps.MicrosoftExcel.Actions;

[ActionList]
public class WorksheetActions : BaseInvocable
{
    public WorksheetActions(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    [Action("Get cell", Description = "Get cell by address")]
    public async Task<CellDto> GetCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.Column}{cellRequest.Row}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var cellValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new CellDto(){ Value = cellValue.Values.First().First() };
    }

    [Action("Update cell", Description = "Update cell by address")]
    public async Task<CellDto> UpdateCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest,
        [ActionParameter] UpdateCellRequest updateCellRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.Column}{cellRequest.Row}')", 
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            values = new[] { new[] { updateCellRequest.Value } }
        });
        var cellValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new CellDto() { Value = cellValue.Values.First().First() };
    }

    [Action("Get row", Description = "Get row by address")]
    public async Task<RowDto> GetRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetRowRequest rowRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{rowRequest.Column1}{rowRequest.RowIndex}:{rowRequest.Column2}{rowRequest.RowIndex}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var rowValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new RowDto() { Row = rowValue.Values.First() };
    }

    [Action("Update row", Description = "Update row by address")]
    public async Task<RowDto> UpdateRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetRowRequest rowRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{rowRequest.Column1}{rowRequest.RowIndex}:{rowRequest.Column2}{rowRequest.RowIndex}')",
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            values = new[] { updateRowRequest.Row }
        });
        var rowValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new RowDto() { Row = rowValue.Values.First() };
    }

    [Action("Create worksheet", Description = "Create worksheet")]
    public async Task<WorksheetDto> CreateWorksheet(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] CreateWorksheetRequest createWorksheetRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets",
            Method.Post, InvocationContext.AuthenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            name = createWorksheetRequest.Name
        });
        return await client.ExecuteWithHandling<WorksheetDto>(request);
    }

    [Action("Get used range", Description = "Get used range")]
    public async Task<RowsDto> GetUsedRange(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/usedRange",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var rowValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new RowsDto() { Rows = rowValue.Values.ToList() };
    }
}