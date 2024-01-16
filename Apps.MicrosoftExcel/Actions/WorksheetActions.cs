using System.Net.Mime;
using System.Text;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Glossaries.Utils.Converters;
using Blackbird.Applications.Sdk.Glossaries.Utils.Dtos;
using RestSharp;

namespace Apps.MicrosoftExcel.Actions;

[ActionList]
public class WorksheetActions : BaseInvocable
{
    private readonly IFileManagementClient _fileManagementClient;
    
    public WorksheetActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient) 
        : base(invocationContext)
    {
        _fileManagementClient = fileManagementClient;
    }

    [Action("Get sheet cell", Description = "Get cell by address")]
    public async Task<CellDto> GetCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.CellAddress}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var cellValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new CellDto(){ Value = cellValue.Values.First().First() };
    }

    [Action("Update sheet cell", Description = "Update cell by address")]
    public async Task<CellDto> UpdateCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest,
        [ActionParameter] UpdateCellRequest updateCellRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.CellAddress}')", 
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            values = new[] { new[] { updateCellRequest.Value } }
        });
        var cellValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new CellDto() { Value = cellValue.Values.First().First() };
    }

    [Action("Get sheet row", Description = "Get row by address")]
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

    [Action("Add new sheet row", Description = "Adds a new row to the first empty line of the sheet")]
    public async Task<RowDto> AddRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] InsertRowRequest insertRowRequest)
    {        
        var range = await GetUsedRange(workbookRequest, worksheetRequest);
        var newRowIndex = range.Rows.First().All(x => string.IsNullOrWhiteSpace(x)) ? 1 : range.Rows.Count + 1;

        var startColumn = insertRowRequest.ColumnAddress ?? "A";

        //var client = new MicrosoftExcelClient();
        //var endColumn = (startColumn.ToExcelColumnIndex() + insertRowRequest.Row.Count - 1).ToExcelColumnAddress();

        //var request = new MicrosoftExcelRequest(
        //    $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{startColumn}{newRowIndex}:{endColumn}{newRowIndex}')/insert",
        //    Method.Post, InvocationContext.AuthenticationCredentialsProviders);
        //request.AddJsonBody(new
        //{
        //    shift = "Down",

        //});
        //await client.ExecuteWithHandling(request);
        return await UpdateRow(workbookRequest, worksheetRequest, new UpdateRowRequest { Row = insertRowRequest.Row, CellAddress = startColumn + newRowIndex});
    }

    [Action("Update sheet row", Description = "Update row by start address")]
    public async Task<RowDto> UpdateRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
    {
        var client = new MicrosoftExcelClient();
        var (startColumn, row) = updateRowRequest.CellAddress.ToExcelColumnAndRow();
        var endColumn = startColumn + updateRowRequest.Row.Count - 1;
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{startColumn.ToExcelColumnAddress()}{row}:{endColumn.ToExcelColumnAddress()}{row}')",
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

    [Action("Get sheet range", Description = "Get a specific range of rows and columns in a sheet")]
    public async Task<RowsDto> GetRange(
    [ActionParameter] WorkbookRequest workbookRequest,
    [ActionParameter] WorksheetRequest worksheetRequest,
    [ActionParameter] GetRangeRequest rangeRequest)
    {
        if (!rangeRequest.Range.IsValidExcelRange())
            throw new Exception($"{rangeRequest.Range} is not a valid range. Please use the Excel format e.g. 'A1:F9'.");

        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{rangeRequest.Range}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var rowValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        return new RowsDto() { Rows = allRows.Select(x => x.ToList()).ToList() };
    }

    [Action("Get sheet used range", Description = "Get used range in a sheet")]
    public async Task<RowsDto> GetUsedRange(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest)
    {
        var client = new MicrosoftExcelClient();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/usedRange",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders);
        var rowValue = await client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        return new RowsDto() { Rows = allRows.Select(x => x.ToList()).ToList() };
    }

    [Action("Download sheet CSV file", Description = "Download CSV file")]
    public async Task<FileResponse> DownloadCSV(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest)
    {
        var rows = await GetUsedRange(workbookRequest, worksheetRequest);
        var csv = new StringBuilder();
        rows.Rows.ForEach(row =>
        {
            csv.AppendLine(string.Join(",", row));
        });

        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(csv.ToString()));
        var csvFile = await _fileManagementClient.UploadAsync(stream, MediaTypeNames.Text.Csv, "Table.csv");
        return new(csvFile);
    }

    [Action("Import glossary", Description = "Import glossary as Excel worksheet")]
    public async Task<WorksheetDto> ImportGlossary([ActionParameter] WorkbookRequest workbookRequest, 
        [ActionParameter] GlossaryRequest glossary)
    {
        const string term = "Term";
        const string variations = "Variations";
        const string notes = "Notes";
        const string id = "ID";
        const string subjectField = "Subject field";
        const string definition = "Definition";

        static string? GetColumnValue(string columnName, GlossaryLanguageSection languageSection)
        {
            var languageCode = languageSection.LanguageCode;

            if (columnName == $"{term} ({languageCode})")
                return languageSection.Terms.First().Term;

            if (columnName == $"{variations} ({languageCode})")
            {
                var variations = languageSection.Terms.Skip(1).Select(term => term.Term);
                return string.Join("; ", variations);
            }

            if (columnName == $"{notes} ({languageCode})")
            {
                var notes = languageSection.Terms.Select(term =>
                    term.Notes == null ? string.Empty : term.Term + ": " + string.Join(';', term.Notes));
                return string.Join("; ", notes.Where(note => note != string.Empty));
            }

            return null;
        }
        
        var glossaryStream = await _fileManagementClient.DownloadAsync(glossary.Glossary);
        var blackbirdGlossary = await glossaryStream.ConvertFromTBX();

        var worksheet = await CreateWorksheet(workbookRequest,
            new() { Name = blackbirdGlossary.Title ?? Path.GetFileNameWithoutExtension(glossary.Glossary.Name)! });

        var languagesPresent = blackbirdGlossary.ConceptEntries
            .SelectMany(entry => entry.LanguageSections)
            .Select(section => section.LanguageCode)
            .Distinct();
        
        var languageRelatedColumns = languagesPresent
            .SelectMany(language => new[] { term, variations, notes }
            .Select(suffix => $"{suffix} ({language})"))
            .ToList();

        await AddRow(workbookRequest, new() { Worksheet = worksheet.Id }, new()
        {
            Row = new List<string>(new[] { id, definition, subjectField, notes }.Concat(languageRelatedColumns))
        });

        foreach (var entry in blackbirdGlossary.ConceptEntries)
        {
            var languageRelatedValues = (IEnumerable<string>)entry.LanguageSections
                .SelectMany(languageSection => languageRelatedColumns
                    .Select(column => GetColumnValue(column, languageSection)))
                .Where(value => value != null);
            
            await AddRow(workbookRequest, new() { Worksheet = worksheet.Id }, new()
            {
                Row = new List<string>(new[]
                {
                    entry.Id, entry.Definition ?? "", entry.SubjectField ?? "",
                    string.Join(';', entry.Notes ?? Enumerable.Empty<string>())
                }.Concat(languageRelatedValues))
            });
        }

        return worksheet;
    }
}