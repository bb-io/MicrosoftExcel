using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class GetRowRequest
    {
        [Display("Row index")]
        public string RowIndex { get; set; }

        [Display("Start column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column1 { get; set; }

        [Display("End column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column2 { get; set; }
    }
}
