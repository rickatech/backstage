using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.WordAddin.Utils
{
    class MetadataHandler
    {
        public static void GenerateNewMetadataFile(string filePath,string content)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    string newPath = ProcessFilePath(filePath);
                    File.Encrypt(filePath);
                    File.WriteAllText(newPath, content);
                }
            }
            catch (Exception ex)
            {
                //TODO: need to write log.
            }
        }
        public static List<T> GetMetaDataInfo<T>(string filePath)
        {
            try
            {
                using (WebClient client = new WebClient())
                {
                    string localurl = AddinModule.CurrentInstance.ProLoopUrl + "/api/filetags" + filePath;
                    string response = client.DownloadString(localurl);
                    var result = JsonConvert.DeserializeObject<List<T>>(response);
                    return result;
                }
            }
            catch (Exception ex)
            {

            }
            return null;
        }
        public static void PostMetaDataInfo(string data, string filePath)
        {
            using (WebClient client = new WebClient())
            {
                string localurl = AddinModule.CurrentInstance.ProLoopUrl + "/api/filetags" + filePath+"?"+data;
                var resposen = client.UploadString(localurl, data);
            }
        }

        private static string ProcessFilePath(string filePath)
        {
            string parentDirctory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            fileName = Path.Combine(fileName,"_info.txt");
            return Path.Combine(parentDirctory, fileName);
        }
    }
}
