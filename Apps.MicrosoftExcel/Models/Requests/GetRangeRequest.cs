using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class GetRangeRequest
    {
        [Display("Range", Description = "(e.g. \"A1:F3\", \"C5:C9\", \"B10:Z12\")")]
        public string Range { get; set; }
    }
}
