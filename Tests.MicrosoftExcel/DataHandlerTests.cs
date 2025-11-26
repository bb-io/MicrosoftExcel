using Apps.MicrosoftExcel.Models.Requests;
using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Tests.MicrosoftExcel;

[TestClass]
public class DataHandlerTests : TestBase
{
    [TestMethod]
    public async Task WorkbookFileDataSourceHandler_ReturnsFilesAndFolders()
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

    [TestMethod]
    public async Task WorkbookFileDataSourceHandler_IncorrectSite_ThrowsMisconfigException()
    {
        // Arrange
        var request = new WorkbookRequest { SiteName = "incorrect" };
        var handler = new WorkbookFileDataSourceHandler(InvocationContext, request);

        // Act
        var ex = await Assert.ThrowsExactlyAsync<PluginMisconfigurationException>(async () => 
            await handler.GetFolderContentAsync(
                new FolderContentDataSourceContext { FolderId = "" },
                CancellationToken.None
            )
        );

        Assert.Contains("was not found", ex.Message);
    }

    [TestMethod]
    public async Task SiteDataSourceHandler_ReturnsSiteNames()
    {
        // Arrange
        var handler = new SiteDataSourceHandler(InvocationContext);

        // Act
        var response = await handler.GetDataAsync(
            new DataSourceContext { SearchString = "" },
            CancellationToken.None
        );

        // Assert
        foreach (var item in response)
            Console.WriteLine($"Id: {item.Value}, Name: {item.DisplayName}");

        Assert.IsNotNull(response);
    }
}
