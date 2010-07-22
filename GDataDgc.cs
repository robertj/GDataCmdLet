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


    }
}

