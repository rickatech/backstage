using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin
{
    public class ProLoopFile
    {
        public string Path
        {
            get;
            set;
        }

        public string Name
        {
            get;
            set;
        }
        public string LockingUserId
        {
            get;
            set;
        }
        public string VersionId
        {
            get;
            set;
        }
        public string Id { get; set; }
        [Newtonsoft.Json.JsonProperty("dir")]
        public bool isDirctory
        {
            get; set;
        }
    }
}
