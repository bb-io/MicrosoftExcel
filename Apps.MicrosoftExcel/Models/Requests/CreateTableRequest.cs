using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class CreateTableRequest
    {
        [Display("Start address", Description = "Column address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string ColumnRow1 { get; set; }

        [Display("End address", Description = "Column address (e.g. \"A2\", \"B2\", \"C2\")")]
        public string ColumnRow2 { get; set; }

        [Display("Has headers")]
        public bool? HasHeaders { get; set; }
    }
}
