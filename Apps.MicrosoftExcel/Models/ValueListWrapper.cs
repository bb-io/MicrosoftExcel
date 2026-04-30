namespace Apps.MicrosoftExcel.Models;

public class ValueListWrapper<T>
{
    public IEnumerable<T> Value { get; set; }
}