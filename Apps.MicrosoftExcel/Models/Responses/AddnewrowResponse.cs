using Apps.MicrosoftExcel.Dtos;
using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftExcel.Models.Responses;

public class AddnewrowResponse
{
    [Display("Row values")]
    public RowDto Row { get; set; }

    [Display("Row number")]
    public double RowNumber {get; set;}
}
