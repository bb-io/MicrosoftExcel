using Apps.MicrosoftExcel.Extensions;

namespace Tests.MicrosoftExcel;

[TestClass]
public class CellValueExtensionsTests
{
    [TestMethod]
    public void ReplaceEmptyCellPlaceholders_RowWithPlaceholders_ReplacesThemWithEmptyValues()
    {
        var row = new List<string> { "Job 1", "<EMPTY>", " <empty> ", "Done" };

        var result = row.ReplaceEmptyCellPlaceholders();

        CollectionAssert.AreEqual(new List<string> { "Job 1", string.Empty, string.Empty, "Done" }, result);
    }

    [TestMethod]
    public void ReplaceEmptyCellPlaceholders_RowWithoutPlaceholders_KeepsValuesUnchanged()
    {
        var row = new List<string> { "-", "EMPTY", "<EMPTY> value", "" };

        var result = row.ReplaceEmptyCellPlaceholders();

        CollectionAssert.AreEqual(new List<string> { "-", "EMPTY", "<EMPTY> value", "" }, result);
    }
}
