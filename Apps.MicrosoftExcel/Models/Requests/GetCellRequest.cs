using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class GetCellRequest
    {
        [Display("Column address", Description = "Cell column address (e.g. \"A\", \"B\", \"C\")")]
        public string Column { get; set; }

        [Display("Row address", Description = "Cell row address (e.g. \"1\", \"2\", \"3\")")]
        public string Row { get; set; }
    }
}
