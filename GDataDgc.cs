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
        #region XmlReturn

        public class XmlReturn
        {
            public string name { get; set; }
            public string value { get; set; }
        }

        #endregion XmlReturn

        #region ParseXML

        public class ParseXML
        {
            public IEnumerable<XmlReturn> ListFormat;
            public ParseXML(string XMLString)
            {
                XElement Elem = XElement.Parse(XMLString);
                XNamespace Ns = "http://schemas.google.com/apps/2006";


                ListFormat = from c in Elem.Elements(Ns + "property")
                             select new XmlReturn
                             {
                                 name = c.Attribute("name").Value.ToString(),
                                 value = c.Attribute("value").Value.ToString()
                             };


            }

        }

        #endregion ParseXML
        
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
                var Xml = new ParseXML(_Result);
                
                foreach (var x in Xml.ListFormat)
                {
                    if (x.name == "CustomerID")
                    {
                        CustomerId = x.value;
                    }

                }

                return CustomerId;

            }
            private string CustomerId;

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

            public string PostalAddress;


            #region GetAuthToken

            public ProfileService GetAuthToken(string AdminUser, string AdminPassword)
            {
                var uri = new Uri("https://www.google.com/accounts/ClientLogin");

                // parameters: name1=value1&name2=value2    
                WebRequest WebRequest = WebRequest.Create(uri);

                WebRequest.ContentType = "application/x-www-form-urlencoded";
                WebRequest.Method = "POST";
                WebRequest.Headers.Add("GData-Version: 3.0");
                byte[] Bytes = Encoding.ASCII.GetBytes("&Email=" + AdminUser.Replace("@", "%40") + "&Passwd=" + AdminPassword + "&accountType=HOSTED&service=cp&source=dgctest.com-GDataCmdLet-v0508");
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

            #region GetProfile

            public string GetProfile(ProfileService ProfileService, string ID)
            {

                var Domain = ProfileService.Domain;

                var uri = new Uri("https://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID);



                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "Get";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ProfileService.Token);
                WebRequest.Headers.Add("GData-Version: 3.0");


                WebResponse WebResponse = WebRequest.GetResponse();


                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion GetProfile

            #region SetProfile
            public string Res;
            public XElement Elem;
            public XNamespace Ns;
            public string NewXml;

            public string SetProfile(ProfileService ProfileService, string ID, string PostalAddress, string PhoneNumber, string MobilePhoneNumber, string OtherPhoneNumber, string HomePostalAddress)
            {
            

                
                var Domain = ProfileService.Domain;

                var uri = new Uri("http://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + ID);

                var GoogleProfileService = new GoogleProfileService();
                Res = GoogleProfileService.GetProfile(ProfileService, ID);



                Elem = XElement.Parse(Res.ToString().Trim());

                Ns = "http://schemas.google.com/g/2005";

                string FormatedRes = Res.Replace("\"", "'");

                if (PostalAddress != null)
                {
                    if (FormatedRes.Contains("gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'"))
                    {

                        foreach (var Element in Elem.Elements(Ns + "structuredPostalAddress"))
                        {
                            if (Element.FirstAttribute.Value == "http://schemas.google.com/g/2005#work")
                            {

                                Element.SetElementValue(Ns + "formattedAddress", PostalAddress);
                            }


                        }
                        NewXml = Elem.ToString();
                    }
                    else
                    {
                        NewXml = Elem.ToString().Replace("</entry>", "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'><gd:formattedAddress>" + PostalAddress + "</gd:formattedAddress></gd:structuredPostalAddress></entry>");
                        Elem = XElement.Parse(NewXml.ToString().Trim());
                    }
                }

                if (HomePostalAddress != null)
                {
                    if (FormatedRes.Contains("gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#home'"))
                    {

                        foreach (var Element in Elem.Elements(Ns + "structuredPostalAddress"))
                        {

                            if (Element.FirstAttribute.Value == "http://schemas.google.com/g/2005#home")
                            {
                                //Element.SetValue(PhoneNumber);

                                Element.SetElementValue(Ns + "formattedAddress", HomePostalAddress);
                            }

                        }
                        NewXml = Elem.ToString();
                    }
                    else
                    {
                        NewXml = Elem.ToString().Replace("</entry>", "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#home'><gd:formattedAddress>" + HomePostalAddress + "</gd:formattedAddress></gd:structuredPostalAddress></entry>");
                        Elem = XElement.Parse(NewXml.ToString().Trim());
                    }
                }

                if (PhoneNumber != null)
                {
                    if (Res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#work'"))
                    {

                        foreach (var Element in Elem.Elements(Ns + "phoneNumber"))
                        {

                            if (Element.FirstAttribute.Value == "http://schemas.google.com/g/2005#work")
                            {
                                Element.SetValue(PhoneNumber);
                            }
                            NewXml = Elem.ToString();
                        }
                    }
                    else
                    {
                        NewXml = Elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#work' primary='true'>" + PhoneNumber + "</gd:phoneNumber></entry>");
                        Elem = XElement.Parse(NewXml.ToString().Trim());
                    }
                }

                if (MobilePhoneNumber != null)
                {
                    if (Res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#mobile'"))
                    {
                        foreach (var Element in Elem.Elements(Ns + "phoneNumber"))
                        {
                            if (Element.FirstAttribute.Value == "http://schemas.google.com/g/2005#mobile")
                            {
                                Element.SetValue(MobilePhoneNumber);
                            }
                        }
                        NewXml = Elem.ToString();
                    }
                    else
                    {
                        NewXml = Elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#mobile'>" + MobilePhoneNumber + "</gd:phoneNumber></entry>");
                        Elem = XElement.Parse(NewXml.ToString().Trim());
                    }
                }

                if (OtherPhoneNumber != null)
                {
                    if (Res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#other'"))
                    {
                        foreach (var Element in Elem.Elements(Ns + "phoneNumber"))
                        {
                            if (Element.FirstAttribute.Value == "http://schemas.google.com/g/2005#other")
                            {
                                Element.SetValue(OtherPhoneNumber);
                            }
                            NewXml = Elem.ToString();
                        }
                    }
                    else
                    {
                        NewXml = Elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#other'>" + MobilePhoneNumber + "</gd:phoneNumber></entry>");
                        Elem = XElement.Parse(NewXml.ToString().Trim());
                    }
                }


                // parameters: name1=value1&name2=value2

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.ContentType = "application/atom+xml";
                WebRequest.Method = "PUT";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + ProfileService.Token);
                WebRequest.Headers.Add("GData-Version: 3.0");

                var Post = NewXml;

                byte[] Bytes = Encoding.ASCII.GetBytes(Post);
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
                StreamReader SR1 = new StreamReader(WebResponse.GetResponseStream());

                var _Result1 = SR1.ReadToEnd().Trim();

                //return _Result1;


                return NewXml;

            }

            #endregion SetProfile

        }
        
        #endregion GoogleProfileService
    }
}

