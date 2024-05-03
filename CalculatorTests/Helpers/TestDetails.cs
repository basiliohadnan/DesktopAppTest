using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Consinco.Helpers
{
    public class TestDetails
    {
        public Dictionary<string, object> Properties { get; set; }

        public TestDetails()
        {
            Properties = new Dictionary<string, object>();
        }
    }
}
