using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin.Utils
{
    public class Metadata
    {
        public string  Author { get; set; }
        public string Summary { get; set; }
    }
    public class RootMeta
    {
        public Metadata Meta { get; set; }
    }
}
