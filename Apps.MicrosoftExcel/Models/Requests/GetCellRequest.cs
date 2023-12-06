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
        [Display("Cell address", Description = "Cell address (e.g. \"A1\", \"B2\", \"C3\")")]
        public string CellAddress { get; set; }
    }
}
