using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin.Utils
{
    public class ComboboxItem
    {
        public string Text { get; set; }
        public int index { get; set; }
        public string itemType { get; set; }
        public override string ToString()
        {
            return Text;
        }
    }
}
