using Apps.MicrosoftExcel.Dtos;

namespace Apps.MicrosoftExcel.Models.Responses;

public class ListWorksheetsResponse
{
    public IEnumerable<WorksheetDto> Value { get; set; }
}