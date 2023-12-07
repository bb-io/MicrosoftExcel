using Blackbird.Applications.Sdk.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class AddRowRequest
    {
        [Display("Row", Description = "Row number (e.g. \"1\", \"2\", \"3\")")]
        public string RowIndex { get; set; }

        public List<string> Row { get; set; }
    }
}
