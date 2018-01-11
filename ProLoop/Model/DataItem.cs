using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin.Model
{
    class DataItem
    {
        public string FileName { get; set; }
        public string Author { get; set; }
        public string Database { get; set; }
        public string Client { get; set; }
        public string Matter { get; set; }
        public string DocuNm { get; set; }
        public string Version { get; set; }
        public string Editdate { get; set; }

    }

    class MetaDataInfo
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string Keywords { get; set; }
        public string EditorName { get; set; }
       

    }
}
