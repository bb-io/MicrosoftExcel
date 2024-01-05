using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class InsertRowRequest
    {
        public List<string> Row { get; set; }

        [Display("Start column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string? ColumnAddress { get; set; }
    }
}
