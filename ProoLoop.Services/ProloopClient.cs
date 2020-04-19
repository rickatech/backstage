using ProLoop.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ProLoop.Services
{
    public class ProloopClient : IProloopClient
    {
        private string loginUrl;
        public void GetClientAsync()
        {

        }

        public void GetFilesAsync()
        {
            throw new NotImplementedException();
        }

        public void GetFolderAsync()
        {
            throw new NotImplementedException();
        }

        public void GetMatterAsync()
        {
            throw new NotImplementedException();
        }

        public void GetProjectAsync()
        {
            throw new NotImplementedException();
        }

        public async Task LoginAsync(string username, string password, string url)
        {
            Dictionary<string, string> nameValueCollection = new Dictionary<string, string>
            {
                {
                    "username",
                    username
                },
                {
                    "password",
                    password
                }
            };
            FormUrlEncodedContent content = new FormUrlEncodedContent(nameValueCollection);
            try
            {
                string endPoint = "/api/dir/authenticate";
                HttpClient httpClient = new HttpClient();
                httpClient.BaseAddress = new Uri(url);
                try
                {
                    var response = await httpClient.PostAsync(endPoint, content);
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {

                    }
                }
                catch(Exception ex)
                {

                }
                
            }
            catch (AggregateException ex)
            {
               
            }
        }

        public void UploadFile()
        {
            throw new NotImplementedException();
        }
    }
}
