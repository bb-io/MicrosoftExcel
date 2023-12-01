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
        [Display("Start column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column1 { get; set; }

        [Display("Start row address", Description = "Row address (e.g. \"1\", \"2\", \"3\")")]
        public string Row1 { get; set; }

        [Display("End column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column2 { get; set; }

        [Display("End row address", Description = "Row address (e.g. \"1\", \"2\", \"3\")")]
        public string Row2 { get; set; }

        [Display("Has headers")]
        public bool? HasHeaders { get; set; }
    }
}
