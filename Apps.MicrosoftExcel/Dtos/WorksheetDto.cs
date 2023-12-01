using Blackbird.Applications.Sdk.Common;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Dtos
{
    public class WorksheetDto
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public int Position { get; set; }

        public string Visibility { get; set; }
    }
}
