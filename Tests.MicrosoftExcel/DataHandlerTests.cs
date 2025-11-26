using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Tests.MicrosoftExcel;

[TestClass]
public class DataHandlerTests : TestBase
{
    [TestMethod]
    public async Task WorkbookFileDataSourceHandler_IsSuccess()
    {
        // Arrange
        var request = new WorkbookRequest { SiteName = "" };
        var handler = new WorkbookFileDataSourceHandler(InvocationContext, request);

        // Act
        var response = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = "" }, 
            CancellationToken.None
        );

        // Assert
        foreach (var item in response)
            Console.WriteLine($"Id: {item.Id}, Type: {(item.Type == 0 ? "Folder" : "File")}, Name: {item.DisplayName}");

        Assert.IsNotNull(response);
    }
}
