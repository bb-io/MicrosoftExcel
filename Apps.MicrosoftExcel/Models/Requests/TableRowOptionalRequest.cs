using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class TableRowOptionalRequest
    {
        [Display("Table row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
        public int? RowIndex { get; set; }
    }
}
