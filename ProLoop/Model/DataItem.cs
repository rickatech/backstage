using Newtonsoft.Json;
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

        [JsonProperty(PropertyName = "username")]
        public string EditorName { get; set; }

        [JsonProperty(PropertyName="versionId")]
        public string VersionId { get; set; }

    }
    class History
    {
        public string Id { get; set; }
        public string Username
        { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
    }
    class DataInfoRequest
    {
        public List<History> History { get; set; }
    }
    public class FileMetadataInfo
    {
        public string ProjectName { get; set; }
        public string ClientName { get; set; }
        public string MatterName { get; set; }
    }
}