using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftExcel.Models.Requests;

public class ExportGlossaryRequest
{
    [Display("Title", Description = "The name of the exported glossary")]
    public string? Title { get; set; }
    
    [Display("Source description", Description = "Information or metadata about the source or origin of the " +
                                                 "terminology data contained in the glossary")]
    public string? SourceDescription { get; set; }
}