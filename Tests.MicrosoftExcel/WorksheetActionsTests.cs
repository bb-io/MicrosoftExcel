using Apps.MicrosoftExcel.Actions;
using Apps.MicrosoftExcel.Models.Requests;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Newtonsoft.Json;

namespace Tests.MicrosoftExcel;

[TestClass]
public class WorksheetActionsTests : TestBase
{
    [TestMethod]
    public async Task FindSheetRow_IsSuccess()
    {
        var action = new WorksheetActions(InvocationContext, FileManager);

        var response = await action.FindRow(new WorkbookRequest { WorkbookId = "016FYB3YJRWLXMAN5Z5ZAY3J2FIUZ7MGRG"},
            new WorksheetRequest { Worksheet= "{00000000-0001-0000-0000-000000000000}" },
            new FindRowRequest { ColumnAddress="A", Value= "Netherlands (Dutch)" });

        var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        Console.WriteLine(json);
        Assert.IsNotNull(response);
    }

    [TestMethod]
    public async Task AddRow_CorrectColumnAddress_IsSuccess()
    {
        // Arrange
        var action = new WorksheetActions(InvocationContext, FileManager);
        string correctCellAddress = "A";

        // Act
        var response = await action.AddRow(
            new WorkbookRequest { WorkbookId = "016FYB3YJRWLXMAN5Z5ZAY3J2FIUZ7MGRG" },
            new WorksheetRequest { Worksheet = "{00000000-0001-0000-0000-000000000000}" },
            new InsertRowRequest { ColumnAddress = correctCellAddress, Row = ["HELLo", "123"] }
        );

        // Assert
        var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        Console.WriteLine(json);
        Assert.IsNotNull(response);
    }

    [TestMethod]
    public async Task AddRow_IncorrectColumnAddress_ThrowsException()
    {
        // Arrange
        var action = new WorksheetActions(InvocationContext, FileManager);
        string incorrectCellAddress = "a";

        // Act & Assert
        var ex = await Assert.ThrowsExactlyAsync<PluginMisconfigurationException>(async () =>
            await action.AddRow(
                new WorkbookRequest { WorkbookId = "01WKPATKT7UH5XNZ3DSBFK3YFY5TXMGQV2" },
                new WorksheetRequest { Worksheet = "{00000000-0001-0000-0000-000000000000}" },
                new InsertRowRequest { ColumnAddress = incorrectCellAddress, Row = ["HELLo", "123"] })
        );

        StringAssert.Contains(ex.Message, "is not a valid cell address");
    }

    [TestMethod]
    public async Task UpdateRow_CorrectColumnAddress_IsSuccess()
    {
        // Arrange
        var action = new WorksheetActions(InvocationContext, FileManager);
        string correctCellAddress = "A4";

        // Act
        var response = await action.UpdateRow(
            new WorkbookRequest { WorkbookId = "01WKPATKT7UH5XNZ3DSBFK3YFY5TXMGQV2" },
            new WorksheetRequest { Worksheet = "{00000000-0001-0000-0000-000000000000}" },
            new UpdateRowRequest { CellAddress = correctCellAddress, Row = ["world", "456"] }
        );

        // Assert
        var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        Console.WriteLine(json);
        Assert.IsNotNull(response);
    }

    [TestMethod]
    public async Task UpdateRow_IncorrectColumnAddress_IsSuccess()
    {
        // Arrange
        var action = new WorksheetActions(InvocationContext, FileManager);
        string incorrectCellAddress = "a4";

        // Act & Assert
        var ex = await Assert.ThrowsExactlyAsync<PluginMisconfigurationException>(async () => await action.UpdateRow(
            new WorkbookRequest { WorkbookId = "01WKPATKT7UH5XNZ3DSBFK3YFY5TXMGQV2" },
            new WorksheetRequest { Worksheet = "{00000000-0001-0000-0000-000000000000}" },
            new UpdateRowRequest { CellAddress = incorrectCellAddress, Row = ["world", "456"] })
        );

        // Assert
        StringAssert.Contains(ex.Message, "is not a valid cell address");
    }

    [TestMethod]
    public async Task DownloadCsv_IsSuccess()
    {
        var action = new WorksheetActions(InvocationContext, FileManager);

        var response = await action.DownloadCSV(new WorkbookRequest { WorkbookId = "01ICRCSNYTBQDEZT6R7FEYV4R3V2DOHP5U", SiteName = "MTPlatforms_Blackbird" },
            new WorksheetRequest { Worksheet = "{56448A07-EE25-4E9E-A58E-7499D686EF0B}" });

        var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        Console.WriteLine(json);
        Assert.IsNotNull(response);
    }

    [TestMethod]
    public async Task DownloadPdf_IsSuccess()
    {
        var action = new WorksheetActions(InvocationContext, FileManager);

        var response = await action.DownloadWorkbookPdf(new WorkbookRequest { WorkbookId = "016FYB3YJRWLXMAN5Z5ZAY3J2FIUZ7MGRG"});

        var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        Console.WriteLine(json);
        Assert.IsNotNull(response);
    }
}