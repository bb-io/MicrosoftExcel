using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Dtos
{
    public class RowsDto
    {
        public List<List<string>> Rows { get; set; }

        [Display("Rows Count")]
        public double RowsCount { get; set; }
    }
}
