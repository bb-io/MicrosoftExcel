using Apps.MicrosoftExcel.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class TableRequest
    {
        [DataSource(typeof(TableDataSourceHandler))]
        public string Table { get; set; }
    }
}
