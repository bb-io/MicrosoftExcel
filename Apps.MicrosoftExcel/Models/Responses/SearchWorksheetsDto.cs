using Apps.MicrosoftExcel.Dtos;

namespace Apps.MicrosoftExcel.Models.Responses;

public record SearchWorksheetsDto(List<WorksheetDto> Worksheets);