using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Serilog;
using Newtonsoft.Json;

namespace ProLoop.WordAddin.Service
{
    class APIHelper
    {
        private static HttpClient localHttpClient =null;
      public static List<Organization> GetOrganizations()
        {
            Log.Debug("GetOrganizations() -- Begin");
           
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<Organization> list = new List<Organization>();
            string text = "/api/organizations?token=" + AddinModule.CurrentInstance.ProLoopToken;
            Log.Debug<string>("URL: {0}", text);
            list = PerformGetOperation<Organization>(text);
            return list;
        }
        public static List<Project> GetProjects()
        {
            Log.Debug("GetProjects() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<Project> list = new List<Project>();
            string text = "/api/projects?token=" + AddinModule.CurrentInstance.ProLoopToken;
            list = PerformGetOperation<Project>(text);
            return list;
        }

        public static List<Client> GetClients(int orgId)
        {
            Log.Debug("GetProjects() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<Client> list = new List<Client>();
            string text = string.Concat(new object[]
             {
                "/api/organizations/",
                orgId,
                "/clients?token=",
                AddinModule.CurrentInstance.ProLoopToken
        });
            list = PerformGetOperation<Client>(text);
            return list;
        }

        public static List<Matter> GetMatters(int orgId,int clientId)
        {
            Log.Debug("GetMatters() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<Matter> list = new List<Matter>();
            string text = string.Concat(new object[]
            {
                "/api/organizations/",
                orgId,
                "/clients/",
               clientId,
                "/matters?token=",
               AddinModule.CurrentInstance.ProLoopToken
            });
            list = PerformGetOperation<Matter>(text);
            return list;
        }

        public static List<ProLoopFolder> GetFolders(string projectName,string orgName,string matterName, string clientName)
        {
            Log.Debug("GetFolders() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<ProLoopFolder> list = new List<ProLoopFolder>();
            string text = string.Empty;
            
            if (!string.IsNullOrEmpty(projectName))
            {
                text = "/api/files/Projects/" + projectName;
                text = text + "/*?token=" + AddinModule.CurrentInstance.ProLoopToken;
            }
            else
            {
                
                if (!string.IsNullOrEmpty(matterName))
                {
                    text = string.Concat(new string[]
                    {
                        "/api/files/Organizations/",
                        orgName,
                        "/",
                        clientName,
                        "/",
                        matterName
                    });
                    text = text + "/*?token=" + AddinModule.CurrentInstance.ProLoopToken;
                }
                else
                {
                    text = "/api/files/Organizations/" +orgName + "/" + clientName;
                    text = text + "/*?token=" + AddinModule.CurrentInstance.ProLoopToken;
                }
            }
            string text2 = Uri.EscapeUriString(text);
            list = PerformGetOperation<ProLoopFolder>(text2);
            return list;
        }

        public static List<ProLoopFolder> GetFolders(string folderPath)
        {
            Log.Debug("GetFolders() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<ProLoopFolder> list = new List<ProLoopFolder>();
            string text = folderPath + "/*?token=" + AddinModule.CurrentInstance.ProLoopToken;
            Log.Debug<string>("URL: {0}", text);
            list = PerformGetOperation<ProLoopFolder>(text);
            Log.Debug("GetFolders() -- End");
            return list;
        }

        public static List<ProLoopFile> GetFiles(string folderPath)
        {
            Log.Debug("GetFiles() -- Begin");
            if (!IsTokenValid())
            {
                GetSignInToken();
            }
            List<ProLoopFile> list = new List<ProLoopFile>();
            folderPath += AddinModule.CurrentInstance.ProLoopToken;
            string text = Uri.EscapeUriString(folderPath);
            Log.Debug<string>("URL: {0}", text);
            list = PerformGetOperation<ProLoopFile>(text);
            Log.Debug("GetFiles() -- End");
            return list;
        }


        private static HttpResponseMessage PerfomPostOperation(string url, FormUrlEncodedContent content)
        {
            HttpResponseMessage result = localHttpClient.PostAsync(url, content).Result;
            return result;
        }


        private static List<T> PerformGetOperation<T>(string text)
        {
            try
            {
                HttpResponseMessage result = localHttpClient.GetAsync(text).Result;

                Log.Debug<bool>("API Response IsSuccessStatusCode: {0}", result.IsSuccessStatusCode);
                bool isSuccessStatusCode = result.IsSuccessStatusCode;
                if (isSuccessStatusCode)
                {
                    string response = result.Content.ReadAsStringAsync().Result;
                    var ObjectList = JsonConvert.DeserializeObject<List<T>>(response);
                    return ObjectList;
                }
                else
                {
                    Log.Debug<HttpStatusCode>("API Response Status Code: {0}", result.StatusCode);
                    Log.Debug<string>("API Response Phrase: {0}", result.ReasonPhrase);
                    Log.Debug<HttpResponseHeaders>("API Response Headers: {0}", result.Headers);
                }
            }
            catch (AggregateException ex)
            {
                Log.Error<string>("Aggregate Exception while calling the API: {0}", ex.Message);
                Log.Error<string>("Inner Exception details: {0}", ex.InnerException.Message);
            }
            catch (Exception ex2)
            {
                Log.Error<string>("Exception while calling the API: {0}", ex2.Message);
                Log.Error<string>("Inner Exception details: {0}", ex2.InnerException.Message);
            }
            Log.Debug("GetOrganizations() -- End");
            return null;
        }
        private static void GetSignInToken()
        {
            Log.Debug("GetSignInToken() -- Begin");
            Dictionary<string, string> nameValueCollection = new Dictionary<string, string>
            {
                {
                    "username",
                    AddinModule.CurrentInstance.ProLoopUsername
                },
                {
                    "password",
                    AddinModule.CurrentInstance.ProLoopPassword
                }
            };
            FormUrlEncodedContent content = new FormUrlEncodedContent(nameValueCollection);
            try
            {
                string endPoint = "/api/authenticate";
                HttpResponseMessage result = PerfomPostOperation(endPoint,content);
                bool isSuccessStatusCode = result.IsSuccessStatusCode;
                if (isSuccessStatusCode)
                {
                    string jsonResp = result.Content.ReadAsStringAsync().Result;
                    var profile = JsonConvert.DeserializeObject<Profile>(jsonResp);
                    AddinModule.CurrentInstance.ProLoopToken = profile.Token; //Extensions.Value<string>(jObject.SelectToken("token"));
                    Log.Debug("Acquired ProLoop Token: {0}", AddinModule.CurrentInstance.ProLoopToken);
                }
                else
                {
                   // MessageBox.Show("Unable to sign in to the ProLoop. Please check the username and password.", "Sign in error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            catch (AggregateException ex)
            {
                Log.Error<string>("Aggregate Exception while calling the API: {0}", ex.Message);
               
                if (ex.InnerException != null)
                {
                    Log.Error<string>("Inner Exception details: {0}", ex.InnerException.Message);
                }
            }
            catch (Exception ex2)
            {
                Log.Error<string>("Exception while calling the API: {0}", ex2.Message);
                
                if (ex2.InnerException!=null)
                {
                    Log.Error<string>("Inner Exception details: {0}", ex2.InnerException.Message);
                }
            }
            Log.Debug("GetSignInToken() -- End");
        }
        private static bool IsTokenValid()
        {
            string text = "/api/organizations?token=" + AddinModule.CurrentInstance.ProLoopToken;
            Log.Debug<string>("URL: {0}", text);
            bool result2;
            try
            {
                HttpResponseMessage result = localHttpClient.GetAsync(text).Result;
                Log.Debug<bool>("API Response IsSuccessStatusCode: {0}", result.IsSuccessStatusCode);
                bool isSuccessStatusCode = result.IsSuccessStatusCode;
                if (isSuccessStatusCode)
                {
                   Log.Debug("IsTokenValid() -- End");
                    result2 = true;
                    return result2;
                }
               Log.Debug<HttpStatusCode>("API Response Status Code: {0}", result.StatusCode);
               Log.Debug<string>("API Response Phrase: {0}", result.ReasonPhrase);
                bool flag = result.ReasonPhrase == "Unauthorized";
                if (flag)
                {
                   Log.Debug("IsTokenValid() -- End");
                    result2 = false;
                    return result2;
                }
            }
            catch (AggregateException ex)
            {
                Log.Error<string>("Aggregate Exception while calling the API: {0}", ex.Message);
                bool flag2 = ex.InnerException != null;
                if (flag2)
                {
                   Log.Error<string>("Inner Exception details: {0}", ex.InnerException.Message);
                }
            }
            catch (Exception ex2)
            {
                Log.Error<string>("Exception while calling the API: {0}", ex2.Message);
                bool flag3 = ex2.InnerException != null;
                if (flag3)
                {
                    Log.Error<string>("Inner Exception details: {0}", ex2.InnerException.Message);
                }
            }
           // Log.Debug("IsTokenValid() -- End");
            result2 = false;
            return result2;

        }

        public static void InitClient( Uri url)
        {
            if(localHttpClient==null)
            {
               
                localHttpClient = new HttpClient
                {
                    BaseAddress = url,
                    Timeout = TimeSpan.FromSeconds(30.0)
                };
                localHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            }
            else
            {
                localHttpClient.BaseAddress = url;
            }
            GetSignInToken();
        }
    }
    
}
