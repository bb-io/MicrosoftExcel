using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class GetTableRowRequest
    {
        [Display("Row index")]
        public string RowIndex { get; set; }
    }
}
