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
    public class WorksheetRequest
    {
        [Display("Worksheet")]
        [DataSource(typeof(WorksheetDataSourceHandler))]
        public string Worksheet { get; set; }
    }
}
