
using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftExcel.Models.Requests
{
    public class FindRowRequest
    {
        [Display("Column address", Description = "Column address (e.g. \"A\", \"B\", \"C\")")]
        public string ColumnAddress { get; set; }

        public string Value { get; set; }
    }
}
