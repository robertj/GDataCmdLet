using System;
using System.Diagnostics;
using System.Management.Automation;
using System.ComponentModel;
using System.Collections.Generic;
using Google.Contacts;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Apps;
using Google.GData.Extensions.Apps;
using Google.GData.Calendar;
using System.Net;
using System.Text;
using System.Linq;
using System.IO;
using System.Web;
using System.Xml;
using System.Xml.Linq;


namespace Microsoft.PowerShell.GData
{

    public class Dgc
    {
        #region ParseXML

        public class ParseXML
        {
            public IEnumerable<XElement> Parse(string XMLString)
            {
                XElement Elem = XElement.Parse(XMLString);
                XNamespace Ns = "http://schemas.google.com/apps/2006";

                IEnumerable<XElement> list = from c in Elem.DescendantsAndSelf()
                    select c.Element(Ns + "property");

                
            }
            
        }

        #endregion
        
        #region GoogleAppService

        public class GoogleAppService
        {
            #region GetDomain

            public string GetDomain(string AdminUser)
            {
                char[] delimiterChars = { '@' };
                
                
                string[] temp = AdminUser.Split(delimiterChars);
                var Domain = temp[1];
                return Domain;

            }

            #endregion GetDomain

            #region GetDomainFromAppService

            public string GetDomainFromAppService(AppsService UserService)
            {
                string Domain = UserService.Domain;
                return Domain;

            }

            #endregion GetDomainFromAppService

            #region GetCustomerId

            public string GetCustomerId(AppsService CustIdService)
            {
                var Token = CustIdService.Groups.QueryClientLoginToken();
                var uri = new Uri("https://apps-apis.google.com/a/feeds/customer/2.0/customerId");

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token);

                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();
                var Xml = new ParseXML();
                var CustomerId = Xml.Parse(_Result);
                return CustomerId;
                

            }

            #endregion

            #region CreateUserAlias

            public string CreateUserAlias(string ID, AppsService UserService, string UserAlias)
            {
                var Domain = UserService.Domain.ToString();
                char[] DelimiterChars = { '@' };


                string[] temp = UserAlias.Split(DelimiterChars);
                var AliasDomain = temp[1];
                
                var Token = UserService.Groups.QueryClientLoginToken();

                var uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + AliasDomain);

                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.Method = "POST";
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token);

                string Post = "<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'><apps:property name=\"aliasEmail\" value=\"" + UserAlias + "\" /><apps:property name=\"userEmail\" value=\"" + ID + "@" + Domain + "\" /></atom:entry>";

                byte[] Bytes = Encoding.ASCII.GetBytes(Post);
                Stream OS = null;
                WebRequest.ContentLength = Bytes.Length;   
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);      
                

                WebResponse WebResponse = WebRequest.GetResponse();
                
                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
                    
 
            }

            #endregion CreateUserAlias

            #region CreateOrganizationUnit

            public string CreateOrganizationUnit(AppsService OrgUnitService, string OrgUnitName, string OrgUnitDescription, string OrgUnitParentOrgUnitPath, bool OrgUnitBlockInheritance)
            {

                return null;

            }

            #endregion CreateOrganizationUnit

            #region RemoveUserAlias

            public string RemoveUserAlias(AppsService UserService, string UserAlias)
            {
                var Domain = UserService.Domain.ToString();
                char[] DelimiterChars = { '@' };


                string[] temp = UserAlias.Split(DelimiterChars);
                var AliasDomain = temp[1];

                var Token = UserService.Groups.QueryClientLoginToken();

                var uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + Domain + "/" + UserAlias);

                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.Method = "DELETE";
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token);

                //string Post = "<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'><apps:property name=\"aliasEmail\" value=\"" + UserAlias + "\" /><apps:property name=\"userEmail\" value=\"" + ID + "@" + Domain + "\" /></atom:entry>";
                string Post = "";
                byte[] Bytes = Encoding.ASCII.GetBytes(Post);
                Stream OS = null;
                WebRequest.ContentLength = Bytes.Length;
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);


                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;


            }

            #endregion RemoveUserAlias

            #region RetriveUserAlias

            public string RetriveUserAlias(string ID, AppsService UserService)
            {
                var Domain = UserService.Domain.ToString();
                char[] DelimiterChars = { '@' };

                var Token = UserService.Groups.QueryClientLoginToken();

                var uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + Domain + "?userEmail=" + ID);

                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token);

                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;


            }

            #endregion RetriveUserAlias

            #region RetriveAllUserAlias

            public string RetriveAllUserAlias(AppsService UserService)
            {
                var Domain = UserService.Domain.ToString();
                char[] DelimiterChars = { '@' };

                var Token = UserService.Groups.QueryClientLoginToken();

                var uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + Domain + "?start=alias@" + Domain);

                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + Token);

                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;


            }

            #endregion RetriveAllUserAlias

        }

        #endregion GoogleAppService

        #region GoogleContactsService

        public class GoogleContactsService
        {
            public string GetDomain(ContactsService ContactService)
            {

                char[] delimiterChars = { '@' };
                string[] temp = ContactService.Credentials.Username.ToString().Split(delimiterChars);
                var Domain = temp[1];

                return Domain;
            }
        }

        #endregion GoogleContactsService

        #region GoogleCalendarsService

        public class GoogleCalendarsService
        {
            public string GetDomain(CalendarService CalendarService)
            {

                char[] delimiterChars = { '@' };
                string[] _temp = CalendarService.Credentials.Username.ToString().Split(delimiterChars);
                var _Domain = _temp[1];

                return _Domain;
            }
        }

        #endregion GoogleCalendarsService

        #region GoogleResourceService

        public class GoogleResourceService
        {

            public class ResourceService
            {
                public string AdminUser { get; set; }
                public string Token { get; set; }
                public string Domain { get; set; }
            }

            #region GetAuthToken

            public ResourceService GetAuthToken(string AdminUser, string AdminPassword)
            {
                var uri = new Uri("https://www.google.com/accounts/ClientLogin");

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/x-www-form-urlencoded";
                WebRequest.Method = "POST";

                byte[] Bytes = Encoding.ASCII.GetBytes("&Email=" + AdminUser + "&Passwd=" + AdminPassword.Replace("@", "%40") + "&accountType=HOSTED_OR_GOOGLE&service=apps&source=companyName-applicationName-versionID");
                Stream OS = null;

                WebRequest.ContentLength = Bytes.Length;   //Count bytes to send
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);         //Send it

                OS.Close();


                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var Result = SR.ReadToEnd().Trim();


                char[] _Deliminator = { ';' };

                var ResE = Result.Replace("Auth=", ";");
                var ResS = ResE.Split(_Deliminator);
                var Token = ResS[1].ToString();

                char[] delimiterChars = { '@' };


                string[] temp = AdminUser.Split(delimiterChars);
                var Domain = temp[1];

                
                var _Entry = new ResourceService
                {
                    AdminUser = AdminUser,
                    Token = Token,
                    Domain = Domain,
                };
                
               

                return _Entry;
            }

            #endregion GetAuthToken

            #region NewResource

            public string NewResource(ResourceService ResourceService, string ResourceId, string ResourceType, string ResurceDescription)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + Domain);

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "POST";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);
                byte[] Bytes = Encoding.ASCII.GetBytes("<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'><apps:property name='resourceId' value='" + ResourceId + "'/><apps:property name='resourceCommonName' value='" + ResourceId + "'/><apps:property name='resourceDescription' value='" + ResurceDescription + "'/><apps:property name='resourceType' value='" + ResourceType + "'/></atom:entry>");
                Stream OS = null;
                WebRequest.ContentLength = Bytes.Length;   
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);         

                OS.Close();

                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion NewResource

            #region RetriveResource

            public string RetriveResource(ResourceService ResourceService, string ResourceId)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + Domain + "/" + ResourceId);

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "GET";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);
             

                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion RetriveResource

            #region RetriveResource

            public string RetriveAllResources(ResourceService ResourceService)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + Domain + "/");

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "GET";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);


                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion RetriveAllResources

            #region RemoveResource

            public string RemoveResources(ResourceService ResourceService, string ResourceID)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + Domain + "/" + ResourceID);

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "DELETE";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);

                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion RemoveResources

            #region SetResource

            public string SetResource(ResourceService ResourceService, string ResourceID, string ResourceDescription, string ResourceType)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + Domain + "/" + ResourceID);

                // parameters: name1=value1&name2=value2

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "PUT";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);
                byte[] Bytes = Encoding.ASCII.GetBytes("<atom:entry xmlns:atom='http://www.w3.org/2005/Atom'><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceCommonName' value='" + ResourceID + "'/><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceDescription' value='" + ResourceDescription + "'/><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceType' value='" + ResourceType + "'/></atom:entry>");
                Stream OS = null;
                WebRequest.ContentLength = Bytes.Length;
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);

                OS.Close();

                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion SetResources

        }
        #endregion GoogleResourceService
        
        #region GoogleProfileService
        
        public class GoogleProfileService
        {

            public class ProfileService
            {
                public string AdminUser { get; set; }
                public string Token { get; set; }
                public string Domain { get; set; }
            }

            public string PostalAddress
            {
                set;
                get;
            }
          

            #region GetAuthToken

            public ProfileService GetAuthToken(string AdminUser, string AdminPassword)
            {
                var uri = new Uri("https://www.google.com/accounts/ClientLogin");

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/x-www-form-urlencoded";
                WebRequest.Method = "POST";

                byte[] Bytes = Encoding.ASCII.GetBytes("&Email=" + AdminUser + "&Passwd=" + AdminPassword.Replace("@", "%40") + "&accountType=HOSTED_OR_GOOGLE&service=apps&source=companyName-applicationName-versionID");
                Stream OS = null;

                WebRequest.ContentLength = Bytes.Length;   //Count bytes to send
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);         //Send it

                OS.Close();


                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var Result = SR.ReadToEnd().Trim();


                char[] _Deliminator = { ';' };

                var ResE = Result.Replace("Auth=", ";");
                var ResS = ResE.Split(_Deliminator);
                var Token = ResS[1].ToString();

                char[] delimiterChars = { '@' };


                string[] temp = AdminUser.Split(delimiterChars);
                var Domain = temp[1];


                var _Entry = new ProfileService
                {
                    AdminUser = AdminUser,
                    Token = Token,
                    Domain = Domain,
                };



                return _Entry;
            }

            #endregion GetAuthToken

            #region SetProfile

            public string SetProfile(ProfileService ProfileService, string ID, string ProfilePostalAddress)
            {
                /*
                var DGCGoogleAppsService = new Dgc.GoogleAppService();
                var Domain = DGCGoogleAppsService.GetDomain(ResourceService.AdminUser);
                */

                var Domain = ResourceService.Domain;

                var uri = new Uri("https://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID);

                // parameters: name1=value1&name2=value2

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "PUT";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ProfileService.Token);

                if (ProfilePostalAddress != null)
                {
                    PostalAddress = "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'><gd:formattedAddress>" + ProfilePostalAddress + "</gd:formattedAddress></gd:structuredPostalAddress>";
                }

                var Post = "<entry xmlns='http://www.w3.org/2005/Atom'xmlns:gContact='http://schemas.google.com/contact/2008'xmlns:batch='http://schemas.google.com/gdata/batch'xmlns:gd='http://schemas.google.com/g/2005'<category scheme='http://schemas.google.com/g/2005#kind'term='http://schemas.google.com/contact/2008#profile' /><id>http://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID + "</id><link rel='self' type='application/atom+xml'href='http://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID + "' /><link rel='edit' type='application/atom+xml'href='http://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID + "' />" + PostalAddress + "</entry>";
                
                byte[] Bytes = Encoding.ASCII.GetBytes("");
                Stream OS = null;
                WebRequest.ContentLength = Bytes.Length;
                OS = WebRequest.GetRequestStream();
                OS.Write(Bytes, 0, Bytes.Length);

                OS.Close();

                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion SetProfile

        }
        
        #endregion GoogleProfileService
    }
}

