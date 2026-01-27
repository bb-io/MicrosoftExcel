namespace Apps.MicrosoftExcel.Polling.Models;

public class WorkbookUpdatedMemory
{
    public DateTime? LastModifiedDateTime { get; set; }
    public DateTime LastPollingTime { get; set; }
    public bool Triggered { get; set; }
}

