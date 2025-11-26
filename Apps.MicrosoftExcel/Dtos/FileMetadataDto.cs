namespace Apps.MicrosoftExcel.Dtos;

public class FileMetadataDto
{
    public string Id { get; set; }    
    public string Name { get; set; }
    public object? Folder { get; set; }
    public long? Size { get; set; }
    public DateTime? LastModifiedDateTime { get; set; }
    public ParentReferenceDto ParentReference { get; set; }
}