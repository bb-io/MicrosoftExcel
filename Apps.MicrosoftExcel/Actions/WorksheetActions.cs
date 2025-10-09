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
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Glossaries.Utils.Converters;
using Blackbird.Applications.Sdk.Glossaries.Utils.Dtos;
using Blackbird.Applications.Sdk.Utils.Extensions.Http;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Apps.MicrosoftExcel.Utils;

namespace Apps.MicrosoftExcel.Actions;

[ActionList("Worksheets")]
public class WorksheetActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient)
    : MicrosoftExcelInvocable(invocationContext)
{
    [Action("Get sheet cell", Description = "Get cell by address")]
    public async Task<CellDto> GetCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        cellRequest.CellAddress = cellRequest.CellAddress.ToUpper();
        ValidateCellAddressParameter(cellRequest);
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.CellAddress}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var cellValue = await ErrorHandler.ExecuteWithErrorHandlingAsync(() => Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request));
        return new CellDto() { Value = cellValue.Values.First().First() };
    }

    [Action("Update sheet cell", Description = "Update cell by address")]
    public async Task<CellDto> UpdateCell(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetCellRequest cellRequest,
        [ActionParameter] UpdateCellRequest updateCellRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        cellRequest.CellAddress = cellRequest.CellAddress.ToUpper();
        ValidateCellAddressParameter(cellRequest);
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{cellRequest.CellAddress}')",
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        request.AddJsonBody(new
        {
            values = new[] { new[] { updateCellRequest.Value } }
        });
        var cellValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new CellDto() { Value = cellValue.Values.First().First() };
    }

    [Action("Get sheet row", Description = "Get row by address")]
    public async Task<RowDto> GetRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] GetRowRequest rowRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{rowRequest.Column1}{rowRequest.RowIndex}:{rowRequest.Column2}{rowRequest.RowIndex}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new RowDto() { Row = rowValue.Values.First() };
    }

    [Action("Add new sheet row", Description = "Adds a new row to the first empty line of the sheet")]
    public async Task<RowDto> AddRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] InsertRowRequest insertRowRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var range = await GetUsedRange(workbookRequest, worksheetRequest);
        var newRowIndex = range.Rows.First().Values.All(x => string.IsNullOrWhiteSpace(x)) ? 1 : range.Rows.Count + 1;

        var startColumn = insertRowRequest.ColumnAddress ?? "A";

        //var endColumn = (startColumn.ToExcelColumnIndex() + insertRowRequest.Row.Count - 1).ToExcelColumnAddress();

        //var request = new MicrosoftExcelRequest(
        //    $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{startColumn}{newRowIndex}:{endColumn}{newRowIndex}')/insert",
        //    Method.Post, InvocationContext.AuthenticationCredentialsProviders);
        //request.AddJsonBody(new
        //{
        //    shift = "Down",

        //});
        //await Client.ExecuteWithHandling(request);
        return await UpdateRow(workbookRequest, worksheetRequest, new UpdateRowRequest { Row = insertRowRequest.Row, CellAddress = startColumn + newRowIndex });
    }

    [Action("Update sheet row", Description = "Update row by start address")]
    public async Task<RowDto> UpdateRow(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] UpdateRowRequest updateRowRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var (startColumn, row) = updateRowRequest.CellAddress.ToExcelColumnAndRow();
        var endColumn = startColumn + updateRowRequest.Row.Count - 1;
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{startColumn.ToExcelColumnAddress()}{row}:{endColumn.ToExcelColumnAddress()}{row}')",
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        request.AddJsonBody(new
        {
            values = new[] { updateRowRequest.Row }
        });
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        return new RowDto() { Row = rowValue.Values.First() };
    }

    [Action("Create worksheet", Description = "Create worksheet")]
    public async Task<WorksheetDto> CreateWorksheet(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] CreateWorksheetRequest createWorksheetRequest)
    {
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets",
            Method.Post, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        request.AddJsonBody(new
        {
            name = createWorksheetRequest.Name
        });
        return await Client.ExecuteWithHandling<WorksheetDto>(request);
    }

    [Action("Get sheet range", Description = "Get a specific range of rows and columns in a sheet")]
    public async Task<RowsDto> GetRange(
    [ActionParameter] WorkbookRequest workbookRequest,
    [ActionParameter] WorksheetRequest worksheetRequest,
    [ActionParameter] GetRangeRequest rangeRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        if (!rangeRequest.Range.IsValidExcelRange())
            throw new PluginMisconfigurationException($"{rangeRequest.Range} is not a valid range. Please use the Excel format e.g. 'A1:F9'.");
        var (startColumn, startCell) = rangeRequest.Range.Split(":")[0].ToExcelColumnAndRow();
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{rangeRequest.Range}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        var rangeIDs = GetIdsRange(startCell, startCell + rowValue.Values.Count() - 1);

        return new RowsDto()
        {
            Rows = rangeIDs.Zip(allRows, (id, rowvalues) => new _row { RowId = id, Values = rowvalues }).ToList(),
            RowsCount = (double)rowValue.Values.Count()
        };
    }


    [Action("Get sheet used range", Description = "Get used range in a sheet")]
    public async Task<RowsDto> GetUsedRange(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/usedRange",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        var rangeIDs = GetIdsRange(1, rowValue.Values.Count());

        return new RowsDto()
        {
            Rows = rangeIDs.Zip(allRows, (id, rowvalues) => new _row { RowId = id, Values = rowvalues }).ToList(),
            RowsCount = (double)rowValue.Values.Count()
        };
    }

    [Action("Find sheet row", Description = "Providing a column address and a value, return row number where said value is located")]
    public async Task<string?> FindRow([ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest, [ActionParameter] FindRowRequest input)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var range = await GetUsedRange(workbookRequest, worksheetRequest);
        var maxRowIndex = range.Rows.Count;
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/range(address='{input.ColumnAddress}1:{input.ColumnAddress}{maxRowIndex}')",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        var columnValues = allRows.Select(subList => subList.First()).ToList();
        var index = columnValues.IndexOf(input.Value);
        index = index + 1;
        return index == 0 ? null : index.ToString();
    }

    [Action("Download sheet CSV file", Description = "Download CSV file")]
    public async Task<FileResponse> DownloadCSV(
        [ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest)
    {
        ValidateWorksheetParameter(worksheetRequest);
        var rows = await GetUsedRange(workbookRequest, worksheetRequest);
        var csv = new StringBuilder();
        rows.Rows.ForEach(row =>
        {
            csv.AppendLine(string.Join(",", row.Values));
        });

        using var stream = new MemoryStream(Encoding.ASCII.GetBytes(csv.ToString()));
        var csvFile = await fileManagementClient.UploadAsync(stream, MediaTypeNames.Text.Csv, "Table.csv");
        return new(csvFile);
    }

    #region Utils

    private List<int> GetIdsRange(int start, int end)
    {
        var myList = new List<int>();
        for (var i = start; i <= end; i++)
        {
            myList.Add(i);
        }
        return myList;
    }

    private async Task<SimplerRowsDto> GetGlossaryUsedRange(WorkbookRequest workbookRequest, WorksheetRequest worksheetRequest)
    {
        var request = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}/usedRange",
            Method.Get, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var rowValue = await Client.ExecuteWithHandling<MultipleListWrapper<List<string>>>(request);
        var allRows = rowValue.Values.ToList();
        return new SimplerRowsDto()
        {
            Rows = allRows.Select(x => x.ToList()).ToList()
        };
    }
    #endregion


    #region Glossaries

    private const string Term = "Term";
    private const string Variations = "Variations";
    private const string Notes = "Notes";
    private const string Id = "ID";
    private const string SubjectField = "Subject field";
    private const string Definition = "Definition";

    [Action("Import glossary", Description = "Import glossary as Excel worksheet")]
    public async Task<WorksheetDto> ImportGlossary([ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] GlossaryWrapper glossary,
        [ActionParameter] [Display("Overwrite existing sheet",
            Description = "Overwrite an existing sheet if it has the same title as the glossary")]
        bool? overwriteSheet)
    {
        static string? GetColumnValue(string columnName, GlossaryConceptEntry entry, string languageCode)
        {
            var languageSection = entry.LanguageSections.FirstOrDefault(ls => ls.LanguageCode == languageCode);

            if (languageSection != null)
            {
                if (columnName == $"{Term} ({languageCode})")
                    return languageSection.Terms.FirstOrDefault()?.Term;

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

            if (columnName == $"{Term} ({languageCode})" || columnName == $"{Variations} ({languageCode})" ||
                columnName == $"{Notes} ({languageCode})")
                return string.Empty;

            return null;
        }

        await using var glossaryStream = await fileManagementClient.DownloadAsync(glossary.Glossary);
        var blackbirdGlossary = await glossaryStream.ConvertFromTbx();
        var sheetName = blackbirdGlossary.Title ?? Path.GetFileNameWithoutExtension(glossary.Glossary.Name)!;

        var listWorksheetsRequest =
            new MicrosoftExcelRequest($"/items/{workbookRequest.WorkbookId}/workbook/worksheets", Method.Get,
                InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        var listWorksheetsResponse = await Client.ExecuteWithHandling<ListWorksheetsResponse>(listWorksheetsRequest);
        var worksheet = listWorksheetsResponse.Value.FirstOrDefault(sheet => sheet.Name == sheetName);

        if (worksheet != null && (overwriteSheet == null || overwriteSheet.Value == false))
            sheetName += $" {DateTime.Now.ToString("dd-MM-yyyy")}";

        if (worksheet == null || (worksheet != null && (overwriteSheet == null || overwriteSheet.Value == false)))
        {
            const int maxAllowedSheetNameLength = 31;

            if (sheetName.Length > maxAllowedSheetNameLength)
                sheetName = sheetName.Substring(0, maxAllowedSheetNameLength);

            worksheet = await CreateWorksheet(workbookRequest, new() { Name = sheetName });
        }
        else
        {
            var getUsedRangeRequest = new MicrosoftExcelRequest(
                $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{sheetName}/usedRange", Method.Get,
                InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
            var rangeAddress = await Client.ExecuteWithHandling<RangeAddressDto>(getUsedRangeRequest);

            var clearWorksheetRequest = new MicrosoftExcelRequest(
                    $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{sheetName}/range(address='{rangeAddress.Address}')/clear",
                    Method.Post, InvocationContext.AuthenticationCredentialsProviders, workbookRequest)
                .WithJsonBody(new { applyTo = "Contents" });

            await Client.ExecuteWithHandling(clearWorksheetRequest);
        }

        var languagesPresent = blackbirdGlossary.ConceptEntries
            .SelectMany(entry => entry.LanguageSections)
            .Select(section => section.LanguageCode)
            .Distinct()
            .ToList();

        var languageRelatedColumns = languagesPresent
            .SelectMany(language => new[] { Term, Variations, Notes }
            .Select(suffix => $"{suffix} ({language})"))
            .ToList();

        var rowsToAdd = new List<List<string>>();
        rowsToAdd.Add(new List<string>(new[] { Id, Definition, SubjectField, Notes }.Concat(languageRelatedColumns)));

        foreach (var entry in blackbirdGlossary.ConceptEntries)
        {
            var languageRelatedValues = (IEnumerable<string>)languagesPresent
                .SelectMany(languageCode =>
                    languageRelatedColumns
                        .Select(column => GetColumnValue(column, entry, languageCode)))
                .Where(value => value != null);

            rowsToAdd.Add(new List<string>(new[]
            {
                string.IsNullOrWhiteSpace(entry.Id) ? Guid.NewGuid().ToString() : entry.Id,
                entry.Definition ?? "",
                entry.SubjectField ?? "",
                string.Join(';', entry.Notes ?? Enumerable.Empty<string>())
            }.Concat(languageRelatedValues)));
        }

        var startColumn = 1;
        var startRow = 1;

        var endColumn = startColumn + rowsToAdd[0].Count - 1;
        var addRowsRequest = new MicrosoftExcelRequest(
            $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheet.Id}/range(address='{startColumn.ToExcelColumnAddress()}{startRow}:{endColumn.ToExcelColumnAddress()}{rowsToAdd.Count}')",
            Method.Patch, InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
        addRowsRequest.AddJsonBody(new { values = rowsToAdd });
        await Client.ExecuteWithHandling(addRowsRequest);

        return worksheet;
    }

    [Action("Export glossary", Description = "Export glossary from Excel worksheet")]
    public async Task<GlossaryWrapper> ExportGlossary([ActionParameter] WorkbookRequest workbookRequest,
        [ActionParameter] WorksheetRequest worksheetRequest,
        [ActionParameter] ExportGlossaryRequest input)
    {
        var rows = await GetGlossaryUsedRange(workbookRequest, worksheetRequest);
        var maxLength = rows.Rows.Max(list => list.Count);

        var parsedGlossary = new Dictionary<string, List<string>>();

        for (var i = 0; i < maxLength; i++)
        {
            parsedGlossary[rows.Rows[0][i]] = new List<string>(rows.Rows.Skip(1)
                .Select(row => i < row.Count ? row[i] : string.Empty));
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
                        if (i < columnValues.Count && !string.IsNullOrWhiteSpace(columnValues[i]))
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
                                .Select(note => note.Split(": "))
                                .Where(note => note.Length > 1)
                                .Select(note => new { Term = note[0], Notes = note[1] })
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

        var title = input.Title;

        if (title == null)
        {
            var getWorksheetRequest =
                new MicrosoftExcelRequest(
                    $"/items/{workbookRequest.WorkbookId}/workbook/worksheets/{worksheetRequest.Worksheet}", Method.Get,
                    InvocationContext.AuthenticationCredentialsProviders, workbookRequest);
            var worksheet = await Client.ExecuteWithHandling<WorksheetDto>(getWorksheetRequest);
            title = worksheet.Name;
        }

        var glossary = new Glossary(glossaryConceptEntries)
        {
            Title = title,
            SourceDescription = input.SourceDescription
                                ?? $"Glossary export from Microsoft Excel on {DateTime.Now.ToLocalTime().ToString("F")}"
        };

        var glossaryStream = glossary.ConvertToTbx();
        var glossaryFileReference =
            await fileManagementClient.UploadAsync(glossaryStream, MediaTypeNames.Text.Xml, $"{title}.tbx");
        return new() { Glossary = glossaryFileReference };
    }

    #endregion
}