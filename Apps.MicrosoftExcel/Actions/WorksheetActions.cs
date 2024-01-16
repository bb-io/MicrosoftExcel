using System.Globalization;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using Apps.MicrosoftExcel.Dtos;
using Apps.MicrosoftExcel.Extensions;
using Apps.MicrosoftExcel.Models;
using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Files;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Glossaries.Utils.Converters;
using Blackbird.Applications.Sdk.Glossaries.Utils.Dtos;
using Blackbird.Applications.Sdk.Glossaries.Utils.Parsers;
using Blackbird.Applications.Sdk.Utils.Extensions.Files;
using CsvHelper;
using CsvHelper.Configuration;
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
        var newRowIndex = range.Rows.First().Columns.All(x => string.IsNullOrWhiteSpace(x)) ? 1 : range.Rows.Count + 1;

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
        return new RowsDto() { Rows = allRows.Select(x => new ColumnDto() { Columns = x.ToList() }).ToList() };
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
        return new RowsDto() { Rows = allRows.Select(x => new ColumnDto() { Columns = x.ToList() }).ToList() };
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

    #region Glossaries
    
    private const string Term = "Term";
    private const string Variations = "Variations";
    private const string Notes = "Notes";
    private const string Id = "ID";
    private const string SubjectField = "Subject field";
    private const string Definition = "Definition";

    [Action("Import glossary", Description = "Import glossary as Excel worksheet")]
    public async Task<WorksheetDto> ImportGlossary([ActionParameter] WorkbookRequest workbookRequest, 
        [ActionParameter] GlossaryWrapper glossary)
    {
        static string? GetColumnValue(string columnName, GlossaryLanguageSection languageSection)
        {
            var languageCode = languageSection.LanguageCode;

            if (columnName == $"{Term} ({languageCode})")
                return languageSection.Terms.First().Term;

            if (columnName == $"{Variations} ({languageCode})")
            {
                var variations = languageSection.Terms.Skip(1).Select(term => term.Term);
                return string.Join(';', variations);
            }

            if (columnName == $"{Notes} ({languageCode})")
            {
                var notes = languageSection.Terms.Select(term =>
                    term.Notes == null ? string.Empty : term.Term + ": " + string.Join(';', term.Notes));
                return string.Join(";; ", notes.Where(note => note != string.Empty));
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
            .SelectMany(language => new[] { Term, Variations, Notes }
            .Select(suffix => $"{suffix} ({language})"))
            .ToList();

        var rowsToAdd = new List<List<string>>();
        rowsToAdd.Add(new List<string>(new[] { Id, Definition, SubjectField, Notes }.Concat(languageRelatedColumns)));

        foreach (var entry in blackbirdGlossary.ConceptEntries)
        {
            var languageRelatedValues = (IEnumerable<string>)entry.LanguageSections
                .SelectMany(languageSection => languageRelatedColumns
                    .Select(column => GetColumnValue(column, languageSection)))
                .Where(value => value != null);
            
            rowsToAdd.Add(new List<string>(new[]
            {
                entry.Id, entry.Definition ?? "", entry.SubjectField ?? "",
                string.Join(';', entry.Notes ?? Enumerable.Empty<string>())
            }.Concat(languageRelatedValues)));
        }
        
        var startColumn = 1;
        var startRow = 1;
        
        var endColumn = startColumn + rowsToAdd[0].Count - 1;
        var addRowsRequest = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheet.Id}/range(address='{startColumn.ToExcelColumnAddress()}{startRow}:{endColumn.ToExcelColumnAddress()}{rowsToAdd.Count}')",
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders);
        addRowsRequest.AddJsonBody(new { values = rowsToAdd });
        await new MicrosoftExcelClient().ExecuteWithHandling(addRowsRequest);

        return worksheet;
    }

    [Action("Export glossary", Description = "Export glossary from Excel worksheet")]
    public async Task<GlossaryWrapper> ExportGlossary([ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] [Display("Title")] string? title,
        [ActionParameter] [Display("Source description")] string? sourceDescription)
    {
        var rows = await GetUsedRange(workbookRequest, worksheetRequest);
        var maxLength = rows.Rows.Max(list => list.Columns.Count);

        var parsedGlossary = new Dictionary<string, List<string>>();

        for (var i = 0; i < maxLength; i++)
        {
            parsedGlossary[rows.Rows[0].Columns[i]] = new List<string>(rows.Rows.Skip(1)
                .Select(row => i < row.Columns.Count ? row.Columns[i] : string.Empty));
        }
        
        var glossaryConceptEntries = new List<GlossaryConceptEntry>();

        var entriesCount = rows.Rows.Count - 1;
        
        for (var i = 0; i < entriesCount; i++)
        {
            string entryId = null;
            string? entryDefinition = null;
            string? entrySubjectField = null;
            List<string>? entryNotes = null;
            
            var languageSections = new List<GlossaryLanguageSection>();

            foreach (var column in parsedGlossary)
            {
                var columnName = column.Key;
                var columnValues = column.Value;
                
                switch (columnName)
                {
                    case Id:
                        entryId = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entryId))
                            entryId = Guid.NewGuid().ToString();
                        
                        break;
                    
                    case Definition:
                        entryDefinition = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entryDefinition))
                            entryDefinition = null;
                        
                        break;
                    
                    case SubjectField:
                        entrySubjectField = i < columnValues.Count ? columnValues[i].Trim() : string.Empty;

                        if (string.IsNullOrWhiteSpace(entrySubjectField))
                            entrySubjectField = null;
                        
                        break;
                    
                    case Notes:
                        entryNotes = (i < columnValues.Count ? columnValues[i] : string.Empty).Split(';')
                            .Select(value => value.Trim()).ToList();
                        
                        if (entryNotes.All(string.IsNullOrWhiteSpace))
                            entryNotes = null;
                        
                        break;
                    
                    case var languageTerm when new Regex($@"{Term} \(.*?\)").IsMatch(languageTerm):
                        var languageCode = new Regex($@"{Term} \((.*?)\)").Match(languageTerm).Groups[1].Value;
                        if (i < columnValues.Count)
                            languageSections.Add(new(languageCode,
                                new List<GlossaryTermSection>(new GlossaryTermSection[]
                                    { new(columnValues[i].Trim()) })));
                        else
                            languageSections.Add(new(languageCode,
                                new List<GlossaryTermSection>(new GlossaryTermSection[] { new(string.Empty) }))); 
                        break;
                    
                    case var termVariations when new Regex($@"{Variations} \(.*?\)").IsMatch(termVariations):
                        if (i < columnValues.Count)
                        {
                            languageCode = new Regex($@"{Variations} \((.*?)\)").Match(termVariations).Groups[1].Value;
                            var targetLanguageSectionIndex =
                                languageSections.FindIndex(section => section.LanguageCode == languageCode);

                            languageSections[targetLanguageSectionIndex].Terms.AddRange(columnValues[i].Split(';')
                                .Select(term => new GlossaryTermSection(term.Trim())));
                        }
                        break;
                    
                    case var termNotes when new Regex($@"{Notes} \(.*?\)").IsMatch(termNotes):
                        if (i < columnValues.Count)
                        {
                            languageCode = new Regex($@"{Notes} \((.*?)\)").Match(termNotes).Groups[1].Value;
                            var targetLanguageSectionIndex =
                                languageSections.FindIndex(section => section.LanguageCode == languageCode);
                            
                            var notesDictionary = columnValues[i]
                                .Split(";; ")
                                .Select(note => new { Term = note.Split(": ")[0], Notes = note.Split(": ")[1] })
                                .ToDictionary(value => value.Term.Trim(), 
                                    value => value.Notes.Split(';').Select(note => note.Trim()));
                        
                            foreach (var termNotesPair in notesDictionary)
                            {
                                var targetTermIndex = languageSections[targetLanguageSectionIndex].Terms
                                    .FindIndex(term => term.Term == termNotesPair.Key);
                                languageSections[targetLanguageSectionIndex].Terms[targetTermIndex].Notes =
                                    termNotesPair.Value.ToList();
                            }
                        }
                        
                        break;
                }
            }

            var entry = new GlossaryConceptEntry(entryId, languageSections)
            {
                Definition = entryDefinition,
                Notes = entryNotes,
                SubjectField = entrySubjectField
            };
            glossaryConceptEntries.Add(entry);
        }

        if (title == null)
        {
            var client = new MicrosoftExcelClient();
            var getWorksheetRequest =
                new MicrosoftExcelRequest(
                    $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}", Method.Get,
                    InvocationContext.AuthenticationCredentialsProviders);
            var worksheet = await client.ExecuteWithHandling<WorksheetDto>(getWorksheetRequest);
            title = worksheet.Name;
        }
        
        var glossary = new Glossary(glossaryConceptEntries)
        {
            Title = title, 
            SourceDescription = sourceDescription 
                                ?? $"Glossary export from Microsoft Excel on {DateTime.Now.ToLocalTime().ToString("F")}" 
        };

        var glossaryStream = glossary.ConvertToTBX();
        var glossaryFileReference =
            await _fileManagementClient.UploadAsync(glossaryStream, MediaTypeNames.Text.Xml, $"{title}.tbx");
        return new() { Glossary = glossaryFileReference };
    }
    
    private static async Task<Dictionary<string, List<string>>> ParseCsvFile(Stream csvFileStream)
    {
        using var reader = new StreamReader(csvFileStream);
        using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));
        
        var csvDictionary = new Dictionary<string, List<string>>();
        var records = csv.GetRecords<dynamic>().ToList();

        foreach (var record in records)
        {
            var recordDictionary =
                (record as IDictionary<string, object>)!.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToString());

            foreach (var kvp in recordDictionary)
            {
                if (!csvDictionary.ContainsKey(kvp.Key))
                    csvDictionary[kvp.Key] = new List<string>();

                csvDictionary[kvp.Key].Add(kvp.Value ?? "");
            }
        }
        
        return csvDictionary;
    }

    #endregion
}