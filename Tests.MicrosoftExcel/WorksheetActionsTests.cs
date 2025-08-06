using Apps.MicrosoftExcel.Actions;
using Apps.MicrosoftExcel.Models.Requests;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace Tests.MicrosoftExcel
{
    [TestClass]
    public class WorksheetActionsTests : TestBase
    {
        [TestMethod]
        public async Task FindSheetRow_IsSuccess()
        {
            var action = new WorksheetActions(InvocationContext, FileManager);

            var response = await action.FindRow(new WorkbookRequest { WorkbookId = "01ICRCSNYTBQDEZT6R7FEYV4R3V2DOHP5U", SiteName= "MTPlatforms_Blackbird" },
                new WorksheetRequest { Worksheet= "{56448A07-EE25-4E9E-A58E-7499D686EF0B}" },
                new FindRowRequest { ColumnAddress="B", Value= "Netherlands (Dutch)" });

            var json = JsonConvert.SerializeObject(response, Formatting.Indented);
            Console.WriteLine(json);
            Assert.IsNotNull(response);
        }


        //[TestMethod]
        //public async Task GetUsefRowCount_IsSuccess()
        //{
        //    var action = new WorksheetActions(InvocationContext, FileManager);

        //    var response = await action.GetUsedRowCount(new WorkbookRequest { WorkbookId = "01ICRCSNYTBQDEZT6R7FEYV4R3V2DOHP5U", SiteName = "MTPlatforms_Blackbird" },
        //        new WorksheetRequest { Worksheet = "{56448A07-EE25-4E9E-A58E-7499D686EF0B}" });

        //    var json = JsonConvert.SerializeObject(response, Formatting.Indented);
        //    Console.WriteLine(json);
        //    Assert.IsNotNull(response);
        //}

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

    }
}
