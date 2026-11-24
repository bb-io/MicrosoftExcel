namespace Apps.MicrosoftExcel.Extensions;

public static class CellValueExtensions
{
    public const string EmptyCellPlaceholder = "<EMPTY>";

    public static List<string> ReplaceEmptyCellPlaceholders(this IEnumerable<string> cellValues)
    {
        return cellValues.Select(value => value.IsEmptyCellPlaceholder() ? string.Empty : value).ToList();
    }

    private static bool IsEmptyCellPlaceholder(this string? value)
    {
        return string.Equals(value?.Trim(), EmptyCellPlaceholder, StringComparison.OrdinalIgnoreCase);
    }
}
