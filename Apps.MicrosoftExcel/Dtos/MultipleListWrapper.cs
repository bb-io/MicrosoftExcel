using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Apps.MicrosoftExcel.Dtos
{
    public class MultipleListWrapper<T>
    {
        public IEnumerable<T> Values { get; set; }
    }
}
