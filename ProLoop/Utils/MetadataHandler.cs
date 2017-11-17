using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        public static string GetMetadatOfFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    string newPath = ProcessFilePath(filePath);
                    string content = File.ReadAllText(newPath);
                    return content;
                }
                catch (Exception ex)
                {

                }
            }
            return string.Empty;
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
