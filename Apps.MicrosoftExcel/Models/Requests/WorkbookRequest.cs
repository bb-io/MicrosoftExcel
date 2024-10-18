using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class WorkbookRequest
    {
        [Display("Workbook", Description = "Your Excel file")]
        [DataSource(typeof(WorkbookDataSourceHandler))]
        public string WorkbookId { get; set; }

        [Display("Sharepoint site name", Description = "Sharepoint site name")]
        public string? SiteName { get; set; }
    }
}
