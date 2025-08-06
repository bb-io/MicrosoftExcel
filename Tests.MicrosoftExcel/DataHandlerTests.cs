using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common.Dynamic;

namespace Tests.MicrosoftExcel
{
    [TestClass]
    public class DataHandlerTests : TestBase
    {
        [TestMethod]
        public async Task WorkbookDataSourceHandler_IsSuccess()
        {
            var handler = new WorkbookDataSourceHandler(InvocationContext, new Apps.MicrosoftExcel.Models.Requests.WorkbookRequest { });

            var response = await handler.GetDataAsync(new DataSourceContext { SearchString = "" }, CancellationToken.None);

            foreach (var item in response)
            {
                Console.WriteLine($"Id: {item.Key}, Name: {item.Value}");
            }

            Assert.IsNotNull(response);
        }
    }
}
