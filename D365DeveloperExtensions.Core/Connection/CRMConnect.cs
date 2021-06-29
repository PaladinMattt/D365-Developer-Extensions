using System;
using System.Collections.Generic;
using Microsoft.IdentityModel;
using Microsoft.IdentityModel.Clients;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json.Linq;

namespace D365DeveloperExtensions.Core.Connection
{
    public class CrmConnect : IDisposable
    {
        const string token_endpoint = @"https://login.microsoftonline.com/common/oauth2/token";
        private string client_credential_token_endpoint = String.Empty;

        public delegate bool EntityHandler(Entity entity);

        private CrmServiceClient service;
        private OrganizationWebProxyClient sdkService;


        public string Url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }

        public string clientId { get; set; }
        public string clientSecret { get; set; }

        public bool IsReady
        {
            get
            {
                try
                {
                    return Service?.IsReady ??false;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
        }

        public CrmServiceClient Service
        {
            get
            {

                if (service == null || !TokenIsValid(service.CurrentAccessToken))
                {
                    CreateService();
                }
                return service;

            }
            private set
            {
            }
        }

        public bool DateValid(DateTime? dt)
        {
            return (dt ?? DateTime.Now.AddYears(-1)) > DateTime.UtcNow.AddMinutes(30);
        }

        public CrmConnect()
        {

        }

        public bool TokenIsValid(string token)
        {
            if (String.IsNullOrEmpty(token))
            {
                return false;
            }

            try
            {
                var jwtHandler = new JwtSecurityTokenHandler();

                //Check if readable token (string is in a JWT format)
                var readableToken = jwtHandler.CanReadToken(token);

                if (readableToken != true)
                {
                    return false;
                }
                else
                {
                    var jwt = jwtHandler.ReadJwtToken(token);
                    return DateValid(jwt.ValidTo);
                }
            }
            catch
            {
                return false;
            }
        }

        public void LoadSettings(string url, string Username, string Password, string _cId, string _cSecret)
        {

            Url = url;
            userName = Username;
            password = Password;
            clientId = _cId;
            clientSecret = _cSecret;

            CreateService();
        }


        private void CreateService()
        {

            try
            {
                string orguri = Url + @"/XRMServices/2011/Organization.svc";

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;


                string token = ObtainNewToken();

                orguri += "/web";
                Uri org = new Uri(orguri);
                OrganizationWebProxyClient proxy = new OrganizationWebProxyClient(org, false);
                proxy.HeaderToken = token;

                SecureString pass = new NetworkCredential("", password).SecurePassword;
                string region = "crm.dynamics.com";
                service = new CrmServiceClient(userName, pass, region, Url, false, null, null, clientId, null, null, proxy, PromptBehavior.Never, false);


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /*
        private string GetToken()
        {
            Uri resourceURL = new Uri(Url + "/api/data/");

            //Create the Client credentials to pass for authentication
            ClientCredential clientcred = new ClientCredential(clientId, clientSecret);

            //get the authentication parameters
            AuthenticationParameters authParam = AuthenticationParameters.CreateFromUrlAsync(resourceURL).Result;

            //Generate the authentication context - this is the azure login url specific to the tenant
            string authority = authParam.Authority;
            authority = authority.Replace("oauth2/authorize", "");

            AuthenticationContext context = new AuthenticationContext(authority);
            //request token
            AuthenticationResult authenticationResult = context.AcquireTokenAsync(Url, clientcred).Result;

            //get the token              
            string token = authenticationResult.AccessToken;
            return token;
        }
        */

        private string ObtainNewToken()
        {
            string parameters = "grant_type=password&username=" + HttpUtility.UrlEncode(userName)
                      + "&password=" + HttpUtility.UrlEncode(password)
                      + "&client_id=" + HttpUtility.UrlEncode(clientId)
                      + "&client_secret=" + HttpUtility.UrlEncode(clientSecret)
                      + "&resource=" + HttpUtility.UrlEncode(Url);

            byte[] pdata = Encoding.UTF8.GetBytes(parameters);
            //First we need to get the ID for the part
            HttpWebRequest webrequest = (HttpWebRequest)WebRequest.Create(token_endpoint);
            webrequest.Method = "POST";

            //Although not required headers, Microsoft *strongly* recommends adding the odata version
            webrequest.Headers.Add("OData-MaxVersion", "4.0");
            webrequest.Headers.Add("OData-Version", "4.0");

            webrequest.ContentType = "application/x-www-form-urlencoded";
            webrequest.ContentLength = pdata.Length;

            //write the parameters to the data stream
            using (Stream dataStream = webrequest.GetRequestStream())
            {
                dataStream.Write(pdata, 0, pdata.Length);
                using (HttpWebResponse webresponse = webrequest.GetResponse() as HttpWebResponse)
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Token));
                    Token token = ser.ReadObject(webresponse.GetResponseStream()) as Token;
                    return token.access_token;
                }
            }
        }


        /// <summary>
        /// Obtains a new token from CRM
        /// 
        /// </summary>
  

        public Guid GetCallerId()
        {
            IOrganizationService osp = Service;
            if (osp is OrganizationServiceProxy)
            {
                return ((OrganizationServiceProxy)osp).CallerId;
            }
            else if (osp is OrganizationWebProxyClient)
            {
                return ((OrganizationWebProxyClient)osp).CallerId;
            }
            return new Guid();
        }

        public void SetCallerId(Guid callId)
        {
            IOrganizationService osp = Service;
            if (osp is OrganizationServiceProxy)
            {
                ((OrganizationServiceProxy)osp).CallerId = callId;
            }
            else if (osp is OrganizationWebProxyClient)
            {
                ((OrganizationWebProxyClient)osp).CallerId = callId;
            }
            else if (osp is CrmServiceClient)
            {
                ((CrmServiceClient)osp).CallerId = callId;
            }

        }

        public List<Entity> Fetch(string fetchxml, IProgress<ProgressReportData> report = null, int maxRecords = 0)
        {
            //We need to strip the "fetch" part off the xml;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(fetchxml);
            XmlNode root = doc.DocumentElement;
            string entityname = "";
            string existingAttributes = "";
            if (root.Name == "fetch")
            {
                //We extract the fetchcore so we can retrieve all records
                fetchxml = root.FirstChild.OuterXml;
                entityname = root.FirstChild.Attributes["name"].Value.ToString();
                if (root.Attributes != null && root.Attributes.Count > 0)
                {
                    foreach (XmlAttribute att in root.Attributes)
                    {
                        existingAttributes += $" {att.Name}='{att.Value}' ";
                    }
                }
            }

            if (maxRecords <= 12000 && maxRecords != 0)
            {
                existingAttributes += $" top='{maxRecords}' ";
            }
            else
            {
                existingAttributes += "  returntotalrecordcount='true' ";
            }

            fetchxml = "<fetch {0} {1} >" + fetchxml.Replace("{", "{{").Replace("}", "}}") + "</fetch>";

            List<Entity> Entities = new List<Entity>();
            try
            {
                var moreRecords = false;
                int page = 1;
                var cookie = string.Empty;
                string totalRecords = "?";
                do
                {
                    if (report != null)
                    {
                        ProgressReportData rdata = new ProgressReportData();
                        rdata.ProgressPercent = 0;
                        rdata.Text = $"Gathering {entityname}: {Entities.Count} out of {totalRecords}";
                        report.Report(rdata);
                    }

                    var xml = string.Format(fetchxml, cookie, existingAttributes);
                    var collection = Service.RetrieveMultiple(new FetchExpression(xml));

                    if (collection.Entities.Count >= 0)
                    {
                        Entities.AddRange(collection.Entities);
                    }

                    if (maxRecords > 0 && Entities.Count >= maxRecords)
                    {
                        break;
                    }

                    moreRecords = collection.MoreRecords;
                    if (moreRecords)
                    {
                        totalRecords = collection.TotalRecordCount.ToString();

                        page++;
                        cookie = string.Format("paging-cookie='{0}' page='{1}'", System.Security.SecurityElement.Escape(collection.PagingCookie), page);
                    }
                } while (moreRecords);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }

            return Entities;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="funcHandler"></param>
        /// <param name="report"></param>
        /// <returns>If there were any errors.</returns>
        public bool ExecuteOnQuery(QueryExpression expression, EntityHandler funcHandler, IProgress<ProgressReportData> report = null)
        {
            var moreRecords = false;
            bool anyErrors = false;
            int current = 0;
            int count = 0;
            do
            {
                if (report != null)
                {
                    ProgressReportData rdata = new ProgressReportData();
                    rdata.ProgressPercent = 0;
                    rdata.Text = $"Running Functions on {expression.EntityName}: {count} out of {count}";
                    report.Report(rdata);
                }
                var collection = Service.RetrieveMultiple(expression);

                if (collection.Entities.Count >= 0)
                {
                    count += collection.Entities.Count;
                    List<Task> tasks = new List<Task>();
                    foreach (Entity ent in collection.Entities)
                    {
                        if (report != null && current % 100 == 0)
                        {
                            ProgressReportData rdata = new ProgressReportData();
                            rdata.ProgressPercent = 0;
                            rdata.Text = $"Runing Functions on {expression.EntityName}: {current} out of {count}";
                            report.Report(rdata);
                        }
                        current++;

                        if (funcHandler != null)
                        {
                            tasks.Add(Task.Run(() =>
                            {
                                if (!funcHandler.Invoke(ent) && !anyErrors)
                                {
                                    anyErrors = true;
                                }
                            }));

                        }

                        if (tasks.Count > 1000)
                        {
                            Task.WaitAll(tasks.ToArray());
                            tasks.Clear();
                        }
                    }
                    Task.WaitAll(tasks.ToArray());
                    tasks.Clear();
                }

                moreRecords = collection.MoreRecords;
                if (moreRecords)
                {
                    expression.PageInfo.PagingCookie = collection.PagingCookie;
                    expression.PageInfo.PageNumber++;
                }
            } while (moreRecords);
            return anyErrors;
        }

        public List<Entity> Query(QueryExpression expression, IProgress<ProgressReportData> report = null)
        {
            //We need to strip the "fetch" part off the xml;          

            List<Entity> Entities = new List<Entity>();
            var moreRecords = false;

            string totalRecords = "?";
            do
            {
                if (report != null)
                {
                    ProgressReportData rdata = new ProgressReportData();
                    rdata.ProgressPercent = 0;
                    rdata.Text = $"Gathering {expression.EntityName}: {Entities.Count} out of {totalRecords}";
                    report.Report(rdata);
                }

                var collection = Service.RetrieveMultiple(expression);

                if (collection.Entities.Count >= 0)
                {
                    Entities.AddRange(collection.Entities);
                }

                moreRecords = collection.MoreRecords;
                if (moreRecords)
                {
                    expression.PageInfo.PagingCookie = collection.PagingCookie;
                    expression.PageInfo.PageNumber++;

                    totalRecords = collection.TotalRecordCount.ToString();

                }
            } while (moreRecords);


            return Entities;
        }

        public void Dispose()
        {
            service.Dispose();
        }
    }

    public class ProgressReportData
    {
        public int ProgressPercent { get; set; }
        public string Text { get; set; }

    }
}
