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
using Google.GData.Apps.GoogleMailSettings;
using System.Net;
using System.Text;
using System.Linq;
using System.IO;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Google.AccessControl;
using Google.GData.AccessControl;
using System.Security;
using System.Runtime.InteropServices;

namespace Microsoft.PowerShell.GData
{

    public class Dgc
    {
        #region XmlReturn
        public class XmlReturn
        {
            public string Name { get; set; }
            public string Value { get; set; }
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
                                 Name = c.Attribute("name").Value.ToString(),
                                 Value = c.Attribute("value").Value.ToString()
                             };

                
            }

        }

        #endregion ParseXML

        #region ConvertToUnsecureString

        public class ConvertToUnsecureString
        {
            public string PlainString;
            public ConvertToUnsecureString(SecureString securePassword)
            {
                if (securePassword == null)
                {
                    throw new ArgumentNullException("securePassword");
                }

                IntPtr _unmanagedString = IntPtr.Zero;
                _unmanagedString = Marshal.SecureStringToGlobalAllocUnicode(securePassword);
                PlainString = Marshal.PtrToStringUni(_unmanagedString);
            }
        }

        #endregion ConvertToUnsecureString

        #region GoogleAppService

        public class GoogleAppService
        {
            #region GetDomain

            public string GetDomain(string adminUser)
            {
                char[] _delimiterChars = { '@' };
                string[] _temp = adminUser.Split(_delimiterChars);
                var _domain = _temp[1];
                return _domain;
            }

            #endregion GetDomain

            #region GetDomainFromAppService

            public string GetDomainFromAppService(AppsService userService)
            {
                string _domain = userService.Domain;
                return _domain;

            }

            #endregion GetDomainFromAppService

            #region CreateGroupEntrys

            public GDataTypes.GDataGroupEntrys CreateGroupEntrys(AppsExtendedFeed groupEntrys)
            {

                var _gDataGroupEntrys = new GDataTypes.GDataGroupEntrys();
                foreach (AppsExtendedEntry _groupEntry in groupEntrys.Entries)
                {
                    var _gDataGroupEntry = new GDataTypes.GDataGroupEntry();
                    _gDataGroupEntry.SelfUri = _groupEntry.SelfUri.ToString();
                    foreach (var _setting in _groupEntry.Properties)
                    {
                        if (_setting.Name == "groupId")
                        {
                            _gDataGroupEntry.GroupId = _setting.Value;
                        }
                        else if (_setting.Name == "groupName")
                        {
                            _gDataGroupEntry.GroupName = _setting.Value;
                        }
                        else if (_setting.Name == "description")
                        {
                            _gDataGroupEntry.Description = _setting.Value;
                        }
                        else if (_setting.Name == "emailPermission")
                        {
                            _gDataGroupEntry.EmailPermission = _setting.Value;
                        }

                    }
                    _gDataGroupEntrys.Add(_gDataGroupEntry);
                }

                return _gDataGroupEntrys;
            }

            #endregion CreateGroupEntrys

            #region CreateGroupEntry

            public GDataTypes.GDataGroupEntry CreateGroupEntry(AppsExtendedEntry groupEntry)
            {
                var _gDataGroupEntry = new GDataTypes.GDataGroupEntry();
                _gDataGroupEntry.SelfUri = groupEntry.SelfUri.ToString();
                foreach (var _setting in groupEntry.Properties)
                {
                    if (_setting.Name == "groupId")
                    {
                        _gDataGroupEntry.GroupId = _setting.Value;
                    }
                    else if (_setting.Name == "groupName")
                    {
                        _gDataGroupEntry.GroupName = _setting.Value;
                    }
                    else if (_setting.Name == "description")
                    {
                        _gDataGroupEntry.Description = _setting.Value;
                    }
                    else if (_setting.Name == "emailPermission")
                    {
                        _gDataGroupEntry.EmailPermission = _setting.Value;
                    }

                }
                return _gDataGroupEntry;
            }

            #endregion CreateGroupEntry

            #region CreateGroupMemberEntrys

            public GDataTypes.GDataGroupMemberEntrys CreateGroupMemberEntrys(AppsExtendedFeed groupMemberEntrys)
            {

                var _gDataGroupMemberEntrys = new GDataTypes.GDataGroupMemberEntrys();
                foreach (AppsExtendedEntry _groupMemberEntry in groupMemberEntrys.Entries)
                {
                    var _gDataGroupMemberEntry = new GDataTypes.GDataGroupMemberEntry();
                    foreach (var _setting in _groupMemberEntry.Properties)
                    {
                        if (_setting.Name == "memberId")
                        {
                            _gDataGroupMemberEntry.MemberId = _setting.Value;
                        }
                        else if (_setting.Name == "memberType")
                        {
                            _gDataGroupMemberEntry.MemberType = _setting.Value;
                        }
                        else if (_setting.Name == "directMember")
                        {
                            _gDataGroupMemberEntry.DirectMember = _setting.Value;
                        }
                        
                    }
                    _gDataGroupMemberEntrys.Add(_gDataGroupMemberEntry);
                }

                return _gDataGroupMemberEntrys;
            }

            #endregion CreateGroupMemberEntrys

            #region CreateGroupMemberEntry

            public GDataTypes.GDataGroupMemberEntry CreateGroupMemberEntry(AppsExtendedEntry groupMemberEntry)
            {
                var _gDataGroupMemberEntry = new GDataTypes.GDataGroupMemberEntry();
                foreach (var _setting in groupMemberEntry.Properties)
                {
                    if (_setting.Name == "memberId")
                    {
                        _gDataGroupMemberEntry.MemberId = _setting.Value;
                    }
                    else if (_setting.Name == "memberType")
                    {
                        _gDataGroupMemberEntry.MemberType = _setting.Value;
                    }
                    else if (_setting.Name == "directMember")
                    {
                        _gDataGroupMemberEntry.DirectMember = _setting.Value;
                    }

                }
                return _gDataGroupMemberEntry;
            }

            #endregion CreateGroupMemberEntry

            #region CreateGroupOwnerEntrys

            public GDataTypes.GDataGroupOwnerEntrys CreateGroupOwnerEntrys(AppsExtendedFeed groupOwnerEntrys)
            {
                var _gDataGroupOwnerEntrys = new GDataTypes.GDataGroupOwnerEntrys();
                foreach (AppsExtendedEntry _groupOwnerEntry in groupOwnerEntrys.Entries)
                {
                    var _gDataGroupOwnerEntry = new GDataTypes.GDataGroupOwnerEntry();
                    foreach (var _setting in _groupOwnerEntry.Properties)
                    {
                        if (_setting.Name == "email")
                        {
                            _gDataGroupOwnerEntry.MemberId = _setting.Value;
                        }
                        else if (_setting.Name == "type")
                        {
                            _gDataGroupOwnerEntry.MemberType = _setting.Value;
                        }

                    }
                    _gDataGroupOwnerEntrys.Add(_gDataGroupOwnerEntry);
                }
                return _gDataGroupOwnerEntrys;
            }

            #endregion CreateGroupOwnerEntrys

            #region CreateGroupOwnerEntry

            public GDataTypes.GDataGroupOwnerEntry CreateGroupOwnerEntry(AppsExtendedEntry groupOwnerEntry)
            {
                var _gDataGroupOwnerEntry = new GDataTypes.GDataGroupOwnerEntry();
                foreach (var _setting in groupOwnerEntry.Properties)
                {
                    if (_setting.Name == "email")
                    {
                        _gDataGroupOwnerEntry.MemberId = _setting.Value;
                    }
                    else if (_setting.Name == "type")
                    {
                        _gDataGroupOwnerEntry.MemberType = _setting.Value;
                    }
                }
                return _gDataGroupOwnerEntry;
            }

            #endregion CreateGroupOwnerEntry

            #region CreateSenderAddressEntry

            public GDataTypes.GDataSenderAddress CreateSenderAddressEntry(GoogleMailSettingsEntry mailSettingsEntry)
            {
                var _gDataMailSettingsEntry = new GDataTypes.GDataSenderAddress();
                foreach (var _setting in mailSettingsEntry.Properties)
                {
                    if(_setting.Name == "address")
                    {
                        _gDataMailSettingsEntry.Address = _setting.Value;
                    }
                    else if (_setting.Name == "replyTo")
                    {
                        _gDataMailSettingsEntry.ReplyTo = _setting.Value;
                    }
                    else if (_setting.Name == "name")
                    {
                        _gDataMailSettingsEntry.Name = _setting.Value;
                    }
                    else if (_setting.Name == "makeDefault")
                    {
                        _gDataMailSettingsEntry.Default = _setting.Value;
                    }
                }
                return _gDataMailSettingsEntry;
            }

            #endregion CreateSenderAddressEntry

            #region CreateLanguageEntry

            public GDataTypes.GDataLanguage CreateLanguageEntry(GoogleMailSettingsEntry mailSettingsEntry)
            {
                var _gDataLanguageEntry = new GDataTypes.GDataLanguage();
                foreach (var _setting in mailSettingsEntry.Properties)
                {
                    if (_setting.Name == "language")
                    {
                        _gDataLanguageEntry.Language = _setting.Value;
                    }
                }
                return _gDataLanguageEntry;
            }

            #endregion CreateLanguageEntry

            #region CreateWebclipsEntry

            public GDataTypes.GDataWebClips CreateWebClipsEntry(GoogleMailSettingsEntry mailSettingsEntry)
            {
                var _gDataWebClips = new GDataTypes.GDataWebClips();
                foreach (var _setting in mailSettingsEntry.Properties)
                {
                    if (_setting.Name == "enable")
                    {
                        _gDataWebClips.Enabled = _setting.Value;
                    }
                }
                return _gDataWebClips;
            }

            #endregion CreateWebclipsEntry

            #region CreateCreateIMapEntry

            public GDataTypes.GDataImap CreateIMapEntry(GoogleMailSettingsEntry mailSettingsEntry)
            {
                var _gDataImapEntry = new GDataTypes.GDataImap();
                foreach (var _setting in mailSettingsEntry.Properties)
                {
                    if (_setting.Name == "enable")
                    {
                        _gDataImapEntry.Enabled = _setting.Value;
                    }
                }
                return _gDataImapEntry;
            }

            #endregion CreateCreateIMapEntry

            #region CreateCreatePop3Entry

            public GDataTypes.GDataPop3 CreatePop3Entry(GoogleMailSettingsEntry mailSettingsEntry)
            {
                var _pop3Entry = new GDataTypes.GDataPop3();

                foreach (var _setting in mailSettingsEntry.Properties)
                {
                    if (_setting.Name == "enable")
                    {
                        _pop3Entry.Enabled = _setting.Value;
                    }
                    else if (_setting.Name == "action")
                    {
                        _pop3Entry.Action = _setting.Value;
                    }
                    else if (_setting.Name == "enableFor")
                    {
                        _pop3Entry.EnabledFor = _setting.Value;
                    }
                }

                return _pop3Entry;
            }

            #endregion CreateCreatePop3Entry

            #region AppendUserAliasEntry

            public GDataTypes.GDataUserAliasEntry AppendUserAliasEntry(string _Xml, GDataTypes.GDataUserAliasEntry userAliasEntry)
            {
                var _paresdXml = new GDataTypes.ParseXML(_Xml);
                var _userSingelAliasEntry = new GDataTypes.GDataAliasEntry();
                foreach (var _entry in _paresdXml.ListFormat)
                {
                    var _aliasEntry = new GDataTypes.GDataAliasEntry();
                    foreach (var _attribute in _entry.at)
                    {
                        if (_attribute.Value == "aliasEmail" || _attribute.Value == "userEmail")
                        {
                            if (_attribute.Value == "aliasEmail")
                            {
                                _userSingelAliasEntry.aliasEmail = _attribute.NextAttribute.Value;
                            }
                            if (_attribute.Value == "userEmail")
                            {
                                _userSingelAliasEntry.UserName = _attribute.NextAttribute.Value;
                            }
                            if (_userSingelAliasEntry.UserName != null && _userSingelAliasEntry.aliasEmail != null)
                            {
                                userAliasEntry.Add(_userSingelAliasEntry);
                            }
                        }
                    }
                    foreach (var _subEntry in _entry.sub)
                    {
                        foreach (var _attribute in _subEntry.at)
                        {

                            if (_attribute.Value == "aliasEmail" || _attribute.Value == "userEmail")
                            {
                                if (_attribute.Value == "aliasEmail")
                                {
                                    _aliasEntry.aliasEmail = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "userEmail")
                                {
                                    _aliasEntry.UserName = _attribute.NextAttribute.Value;
                                }
                                if (_aliasEntry.UserName != null && _aliasEntry.aliasEmail != null)
                                {
                                    userAliasEntry.Add(_aliasEntry);
                                }
                            }
                        }
                    }
                }
                return userAliasEntry;
            }

            #endregion AppendUserAliasEntry

            #region CreateUserAliasEntry

            public GDataTypes.GDataUserAliasEntry CreateUserAliasEntry(string xml)
            {
                var _paresdXml = new GDataTypes.ParseXML(xml);
                var _userAliasEntry = new GDataTypes.GDataUserAliasEntry();
                var _userSingelAliasEntry = new GDataTypes.GDataAliasEntry();
                foreach (var _entry in _paresdXml.ListFormat)
                {
                    var _aliasEntry = new GDataTypes.GDataAliasEntry();
                    foreach (var _attribute in _entry.at)
                    {
                        if (_attribute.Value == "aliasEmail" || _attribute.Value == "userEmail")
                        {
                            if (_attribute.Value == "aliasEmail")
                            {
                                _userSingelAliasEntry.aliasEmail = _attribute.NextAttribute.Value;
                            }
                            if (_attribute.Value == "userEmail")
                            {
                                _userSingelAliasEntry.UserName = _attribute.NextAttribute.Value;
                            }
                            if (_userSingelAliasEntry.UserName != null && _userSingelAliasEntry.aliasEmail != null)
                            {
                                _userAliasEntry.Add(_userSingelAliasEntry);
                            }
                        }
                    }
                    foreach (var _subEntry in _entry.sub)
                    {       
                        foreach (var _attribute in _subEntry.at)
                        {
                            if (_attribute.Value == "aliasEmail" || _attribute.Value == "userEmail")
                            {
                                if (_attribute.Value == "aliasEmail")
                                {
                                    _aliasEntry.aliasEmail = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "userEmail") 
                                {
                                    _aliasEntry.UserName = _attribute.NextAttribute.Value;
                                }
                                if (_aliasEntry.UserName != null && _aliasEntry.aliasEmail != null)
                                {
                                    _userAliasEntry.Add(_aliasEntry);
                                }
                            }
                        }
                    }
                }
                return _userAliasEntry;
            }

            #endregion CreateAliasEntry

            #region CreateUserEntry

            public GDataTypes.GDataUserEntry CreateUserEntry(UserEntry userEntry)
            {
                var _gDataUserEntry = new GDataTypes.GDataUserEntry();
                _gDataUserEntry.UserName = userEntry.Login.UserName;
                _gDataUserEntry.Admin = userEntry.Login.Admin;
                _gDataUserEntry.susspended = userEntry.Login.Suspended;
                _gDataUserEntry.AgreedToTerms = userEntry.Login.AgreedToTerms;
                _gDataUserEntry.ChangePasswordAtNextLogin = userEntry.Login.ChangePasswordAtNextLogin;
                _gDataUserEntry.GivenName = userEntry.Name.GivenName;
                _gDataUserEntry.FamilyName = userEntry.Name.FamilyName;
                _gDataUserEntry.Limit = userEntry.Quota.Limit;
                _gDataUserEntry.SelfUri = userEntry.SelfUri.ToString();
                return _gDataUserEntry;
            }

            #endregion CreateUserEntry

            #region CreateUserEntrys

            public GDataTypes.GDataUserEntrys CreateUserEntrys(UserFeed _userFeed)
            {
                var _gDataUserEntrys = new GDataTypes.GDataUserEntrys();
                foreach (UserEntry _userEntry in _userFeed.Entries)
                {
                    var _gDataUserEntry = new GDataTypes.GDataUserEntry();
                    _gDataUserEntry.UserName = _userEntry.Login.UserName;
                    _gDataUserEntry.Admin = _userEntry.Login.Admin;
                    _gDataUserEntry.susspended = _userEntry.Login.Suspended;
                    _gDataUserEntry.AgreedToTerms = _userEntry.Login.AgreedToTerms;
                    _gDataUserEntry.ChangePasswordAtNextLogin = _userEntry.Login.ChangePasswordAtNextLogin;
                    _gDataUserEntry.GivenName = _userEntry.Name.GivenName;
                    _gDataUserEntry.FamilyName = _userEntry.Name.FamilyName;
                    _gDataUserEntry.Limit = _userEntry.Quota.Limit;
                    _gDataUserEntry.SelfUri = _userEntry.SelfUri.ToString();
                    _gDataUserEntrys.Add(_gDataUserEntry);
                }
                return _gDataUserEntrys;
            }

            #endregion CreateUserEntrys

            #region GetCustomerId

            public string GetCustomerId(AppsService CustIdService)
            {
                var Token = CustIdService.Groups.QueryClientLoginToken();
                var uri = new Uri("https://apps-apis.google.com/a/feeds/customer/2.0/customerId");

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml; charset=UTF-8";
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
                    if (x.Name == "customerId")
                    {
                        CustomerId = x.Value;
                    }

                }

                return CustomerId;
                

            }
            private string CustomerId;

            #endregion

            #region CreateUserAlias

            public string CreateUserAlias(string ID, AppsService userService, string userAlias)
            {
                var _domain = userService.Domain.ToString();
                char[] _delimiterChars = { '@' };
                string[] _temp = userAlias.Split(_delimiterChars);
                var _aliasDomain = _temp[1];
                var _token = userService.Groups.QueryClientLoginToken();
                var uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + _domain);

                WebRequest _webRequest = WebRequest.Create(uri);

                _webRequest.Method = "POST";
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + _token);

                string _post = "<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'><apps:property name=\"aliasEmail\" value=\"" + userAlias + "\" /><apps:property name=\"userEmail\" value=\"" + ID + "@" + _domain + "\" /></atom:entry>";

                byte[] _bytes = Encoding.UTF8.GetBytes(_post);
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;   
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);      
                
                WebResponse _webResponse = _webRequest.GetResponse();
                
                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = SR.ReadToEnd().Trim();
                return _result;
            }

            #endregion CreateUserAlias

            #region CreateOrganizationUnit

            public string CreateOrganizationUnit(AppsService OrgUnitService, string OrgUnitName, string OrgUnitDescription, string OrgUnitParentOrgUnitPath, bool OrgUnitBlockInheritance)
            {
                return null;
            }

            #endregion CreateOrganizationUnit

            #region RetriveNextPage

            private string nextPage;
            public string RetriveNextPage(string xml)
            {
                var _parseXml = new GDataTypes.ParseXML(xml.ToString());
                nextPage = "";
                foreach (var _elements in _parseXml.ListFormat)
                {
                    foreach (var _attribute in _elements.at)
                    {
                        if (_attribute.Value == "next")
                        {
                            nextPage = _attribute.NextAttribute.NextAttribute.Value;
                        }
                    }
                }
                return nextPage;
            }

            #endregion RetriveNextPage

            #region RemoveUserAlias

            public string RemoveUserAlias(AppsService userService, string userAlias)
            {
                var _domain = userService.Domain.ToString();
                char[] _delimiterChars = { '@' };
                string[] _temp = userAlias.Split(_delimiterChars);
                var _aliasDomain = _temp[1];
                var _token = userService.Groups.QueryClientLoginToken();
                var _uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + _domain + "/" + userAlias);

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.Method = "DELETE";
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + _token);
                
                string _post = "";
                byte[] _bytes = Encoding.UTF8.GetBytes(_post);
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);

                WebResponse _webResponse = _webRequest.GetResponse();

                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());
                var _result = _SR.ReadToEnd().Trim();
                return _result;
            }

            #endregion RemoveUserAlias

            #region RetriveUserAlias

            public string RetriveUserAlias(string id, AppsService userService)
            {
                var _domain = userService.Domain.ToString();
                char[] _delimiterChars = { '@' };
                var _token = userService.Groups.QueryClientLoginToken();
                var _uri = new Uri("https://apps-apis.google.com/a/feeds/alias/2.0/" + _domain + "?userEmail=" + id + "@" + _domain);

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.Method = "GET";
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + _token);

                WebResponse _webResponse = _webRequest.GetResponse();
                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();
                return _result;
            }

            #endregion RetriveUserAlias

            #region RetriveAllUserAlias

            public string RetriveAllUserAlias(AppsService userService, string nextPage)
            {
                var _domain = userService.Domain.ToString();
                char[] _delimiterChars = { '@' };

                var _token = userService.Groups.QueryClientLoginToken();

                if (nextPage == "")
                {
                    nextPage = "https://apps-apis.google.com/a/feeds/alias/2.0/" + _domain + "?start=alias@" + _domain;
                }

                var _uri = new Uri(nextPage);

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.Method = "GET";
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + _token);

                WebResponse _webResponse = _webRequest.GetResponse();
                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();
                return _result;
            }

            #endregion RetriveAllUserAlias

            #region RetrieveAllOUs

            public string RetrieveAllOUs(AppsService UserService)
            {
                var _token = UserService.Groups.QueryClientLoginToken();
                var OUService = new Dgc.GoogleAppService();
                var _custId = OUService.GetCustomerId(UserService);
                                
                var uri = new Uri("https://apps-apis.google.com/a/feeds/orgunit/2.0/" + _custId + "?get=all");

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml; charset=UTF-8";
                WebRequest.Headers.Add("Authorization: GoogleLogin auth=" + _token);

                WebResponse WebResponse = WebRequest.GetResponse();

                if (WebResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader SR = new StreamReader(WebResponse.GetResponseStream());

                var _Result = SR.ReadToEnd().Trim();
                
                return _Result;
                 
            }

            #endregion RetrieveAllOUs

            #region RetrieveOU

            public string RetrieveOU(AppsService UserService, string OuPath)
            {
                var Token = UserService.Groups.QueryClientLoginToken();
                var OUService = new Dgc.GoogleAppService();
                var CustId = OUService.GetCustomerId(UserService);

                var uri = new Uri("https://apps-apis.google.com/a/feeds/orgunit/2.0/" + CustId + "/"+OuPath);

                WebRequest WebRequest = WebRequest.Create(uri);
                WebRequest.Method = "GET";
                WebRequest.ContentType = "application/atom+xml; charset=UTF-8";
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

            #endregion RetrievOU

        }

        #endregion GoogleAppService

        #region GoogleContactsService

        public class GoogleContactsService
        {
            #region GetDomain
            public string GetDomain(ContactsService ContactService)
            {

                char[] delimiterChars = { '@' };
                string[] temp = ContactService.Credentials.Username.ToString().Split(delimiterChars);
                var Domain = temp[1];

                return Domain;
            }
            #endregion GetDomain

            #region SetContactTitle
            private string editUri;
            private string oldTitle;
            public string SetContactTitle(string token, string selfUri, string title)
            {
                var _xml = GetContact(token, selfUri);

                var _paredXml = new GDataTypes.ParseXML(_xml);
                if(_xml.Contains("<title type='text'></title>"))
                {

                }
                foreach (var _entry in _paredXml.ListFormat)
                {
                    if (_entry.name == "{http://www.w3.org/2005/Atom}link")
                    {
                        foreach (var _attribute in _entry.at)
                        {
                            if (_attribute.Value == "edit")
                            {
                                editUri = _attribute.NextAttribute.NextAttribute.Value;
                            }
                        }
                    }
                    if (_entry.name == "{http://www.w3.org/2005/Atom}title")
                    {
                        if (_entry.value != null)
                        {
                            oldTitle = "<title type='text'>" + _entry.value + "</title>";
                        }
                        else
                        {
                            oldTitle = "<title type='text'></title>";
                        }
                    }
                }

                var _uri = new Uri(editUri);

                var _newXml = _xml.Replace(oldTitle, "<title type='text'>" + title + "</title>");

                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                
                _webRequest.Method = "PUT";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + token);
                byte[] _bytes = Encoding.UTF8.GetBytes(_newXml);
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);

                _OS.Close();

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _Result = _SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion SetContact

            #region GetContact

            public string GetContact(string token, string selfUri)
            {
                var _uri = new Uri(selfUri);

                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "Get";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + token);

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _Result = _SR.ReadToEnd().Trim();

                return _Result;
            }

            #endregion GetContact

            #region CreateContactEntry

            public GDataTypes.GDataContactEntry CreateContactEntry(ContactEntry contactEntry)
            {
                var _gDataContactEntry = new GDataTypes.GDataContactEntry();
                _gDataContactEntry.Name = contactEntry.Title.Text;
                foreach (var _emailEntry in contactEntry.Emails)
                {
                    if (_emailEntry.Work == true)
                    {
                        _gDataContactEntry.Email = _emailEntry.Address;
                    }
                }
                foreach (var _phoneNumberEntry in contactEntry.Phonenumbers)
                {
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsWork)
                    {
                        _gDataContactEntry.PhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsHome)
                    {
                        _gDataContactEntry.HomePhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsMobile)
                    {
                        _gDataContactEntry.MobilePhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsOther)
                    {
                        _gDataContactEntry.OtherPhoneNumber = _phoneNumberEntry.Value;
                    }
                }
                foreach (var _postaAddressEntry in contactEntry.PostalAddresses)
                {
                    if (_postaAddressEntry.Rel == ContactsRelationships.IsWork)
                    {
                        _gDataContactEntry.PostalAddress = _postaAddressEntry.FormattedAddress;
                    }
                    if (_postaAddressEntry.Rel == ContactsRelationships.IsHome)
                    {
                        _gDataContactEntry.HomeAddress = _postaAddressEntry.FormattedAddress;
                    }
                }
                _gDataContactEntry.SelfUri = contactEntry.SelfUri.ToString();
                return _gDataContactEntry;
            }

            #endregion CreateContactEntry

            #region CreateContactModifidEntry

            public GDataTypes.GDataContactEntry CreateContactModifidEntry(ContactEntry contactEntry, string name)
            {
                var _gDataContactEntry = new GDataTypes.GDataContactEntry();
                _gDataContactEntry.Name = name;
                foreach (var _emailEntry in contactEntry.Emails)
                {
                    if (_emailEntry.Work == true)
                    {
                        _gDataContactEntry.Email = _emailEntry.Address;
                    }
                }
                foreach (var _phoneNumberEntry in contactEntry.Phonenumbers)
                {
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsWork)
                    {
                        _gDataContactEntry.PhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsHome)
                    {
                        _gDataContactEntry.HomePhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsMobile)
                    {
                        _gDataContactEntry.MobilePhoneNumber = _phoneNumberEntry.Value;
                    }
                    if (_phoneNumberEntry.Rel == ContactsRelationships.IsOther)
                    {
                        _gDataContactEntry.OtherPhoneNumber = _phoneNumberEntry.Value;
                    }
                }
                foreach (var _postaAddressEntry in contactEntry.PostalAddresses)
                {
                    if (_postaAddressEntry.Rel == ContactsRelationships.IsWork)
                    {
                        _gDataContactEntry.PostalAddress = _postaAddressEntry.FormattedAddress;
                    }
                    if (_postaAddressEntry.Rel == ContactsRelationships.IsHome)
                    {
                        _gDataContactEntry.HomeAddress = _postaAddressEntry.FormattedAddress;
                    }
                }
                _gDataContactEntry.SelfUri = contactEntry.SelfUri.ToString();
                return _gDataContactEntry;
            }

            #endregion CreateContactModifidEntry

            #region CreateContactEntrys

            public GDataTypes.GDataContactEntrys CreateContactEntrys(ContactsFeed contactEntrys)
            {
                var _gDataContactEntrys = new GDataTypes.GDataContactEntrys();

                foreach (ContactEntry _contactEntry in contactEntrys.Entries)
                {
                    var _gDataContactEntry = new GDataTypes.GDataContactEntry();
                    _gDataContactEntry.Name = _contactEntry.Title.Text;
                    foreach (var _emailEntry in _contactEntry.Emails)
                    {
                        if (_emailEntry.Work == true)
                        {
                            _gDataContactEntry.Email = _emailEntry.Address;
                        }
                    }
                    foreach (var _phoneNumberEntry in _contactEntry.Phonenumbers)
                    {
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsWork)
                        {
                            _gDataContactEntry.PhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsHome)
                        {
                            _gDataContactEntry.HomePhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsMobile)
                        {
                            _gDataContactEntry.MobilePhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsOther)
                        {
                            _gDataContactEntry.OtherPhoneNumber = _phoneNumberEntry.Value;
                        }
                    }
                    foreach (var _postaAddressEntry in _contactEntry.PostalAddresses)
                    {
                        if (_postaAddressEntry.Rel == ContactsRelationships.IsWork)
                        {
                            _gDataContactEntry.PostalAddress = _postaAddressEntry.FormattedAddress;
                        }
                        if (_postaAddressEntry.Rel == ContactsRelationships.IsHome)
                        {
                            _gDataContactEntry.HomeAddress = _postaAddressEntry.FormattedAddress;
                        }
                    }
                    _gDataContactEntry.SelfUri = _contactEntry.SelfUri.ToString();
                    _gDataContactEntrys.Add(_gDataContactEntry);
                }
                return _gDataContactEntrys;
            }
            #endregion CreateContactEntrys

            #region AppendContactEntrys

            public GDataTypes.GDataContactEntrys AppendContactEntrys(ContactEntry contactEntry, GDataTypes.GDataContactEntrys gDataContactEntrys)
            {
                    var _gDataContactEntry = new GDataTypes.GDataContactEntry();

                    _gDataContactEntry.Name = contactEntry.Title.Text;
                    foreach (var _emailEntry in contactEntry.Emails)
                    {
                        if (_emailEntry.Work == true)
                        {
                            _gDataContactEntry.Email = _emailEntry.Address;
                        }
                    }
                    foreach (var _phoneNumberEntry in contactEntry.Phonenumbers)
                    {
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsWork)
                        {
                            _gDataContactEntry.PhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsHome)
                        {
                            _gDataContactEntry.HomePhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsMobile)
                        {
                            _gDataContactEntry.MobilePhoneNumber = _phoneNumberEntry.Value;
                        }
                        if (_phoneNumberEntry.Rel == ContactsRelationships.IsOther)
                        {
                            _gDataContactEntry.OtherPhoneNumber = _phoneNumberEntry.Value;
                        }
                    }
                    foreach (var _postaAddressEntry in contactEntry.PostalAddresses)
                    {
                        if (_postaAddressEntry.Rel == ContactsRelationships.IsWork)
                        {
                            _gDataContactEntry.PostalAddress = _postaAddressEntry.FormattedAddress;
                        }
                        if (_postaAddressEntry.Rel == ContactsRelationships.IsHome)
                        {
                            _gDataContactEntry.HomeAddress = _postaAddressEntry.FormattedAddress;
                        }
                    }
                    _gDataContactEntry.SelfUri = contactEntry.SelfUri.ToString();
                    gDataContactEntrys.Add(_gDataContactEntry);

                return gDataContactEntrys;
            }
            #endregion AppendContactEntrys
        }

        #endregion GoogleContactsService

        #region GoogleCalendarsService

        public class GoogleCalendarsService
        {
            public string GetDomain(CalendarService calendarService)
            {

                char[] _delimiterChars = { '@' };
                string[] _temp = calendarService.Credentials.Username.ToString().Split(_delimiterChars);
                var _domain = _temp[1];

                return _domain;
            }

            #region CreateCalendarEntry

            public GDataTypes.GDataCalendarEntry CreateCalendarEntry(CalendarEntry calendarEntry)
            {
                var _gDataCalendarEntry = new GDataTypes.GDataCalendarEntry();
                if (calendarEntry.Title.Text != null)
                {
                    _gDataCalendarEntry.Name = calendarEntry.Title.Text;
                }
                if (calendarEntry.Color != null)
                {
                    _gDataCalendarEntry.Color = calendarEntry.Color;
                }
                if (calendarEntry.Summary != null)
                {
                    _gDataCalendarEntry.Description = calendarEntry.Summary.Text;
                }
                if (calendarEntry.Location != null)
                {
                    _gDataCalendarEntry.Location = calendarEntry.Location.ValueString;
                }
                _gDataCalendarEntry.Hidden = calendarEntry.Hidden;
                _gDataCalendarEntry.Selected = calendarEntry.Selected;
                if (calendarEntry.TimeZone != null)
                {
                    _gDataCalendarEntry.TimeZone = calendarEntry.TimeZone;
                }
                if (calendarEntry.AccessLevel != null)
                {
                    _gDataCalendarEntry.AccessLevel = calendarEntry.AccessLevel;
                }
                if (calendarEntry.SelfUri.ToString() != null)
                {
                    _gDataCalendarEntry.SelfUri = calendarEntry.SelfUri.ToString();
                }
                return _gDataCalendarEntry;
            }

            #endregion CreateCalendarEntry

            #region CreateCalendarEntrys

            public GDataTypes.GDataCalendarEntrys CreateCalendarEntrys(CalendarFeed calendarEntrys)
            {
                var _gDataCalendarEntrys = new GDataTypes.GDataCalendarEntrys();
                foreach (CalendarEntry _calendarEntry in calendarEntrys.Entries)
                {
                    var _gDataCalendarEntry = new GDataTypes.GDataCalendarEntry();
                    if (_calendarEntry.Title.Text != null)
                    {
                        _gDataCalendarEntry.Name = _calendarEntry.Title.Text;
                    }
                    if (_calendarEntry.Color != null)
                    {
                        _gDataCalendarEntry.Color = _calendarEntry.Color;
                    }
                    if (_calendarEntry.Summary != null)
                    {
                        _gDataCalendarEntry.Description = _calendarEntry.Summary.Text;
                    }
                    if (_calendarEntry.Location != null)
                    {
                        _gDataCalendarEntry.Location = _calendarEntry.Location.ValueString;
                    }
                    _gDataCalendarEntry.Hidden = _calendarEntry.Hidden;
                    _gDataCalendarEntry.Selected = _calendarEntry.Selected;
                    if (_calendarEntry.TimeZone != null)
                    {
                        _gDataCalendarEntry.TimeZone = _calendarEntry.TimeZone;
                    }
                    if (_calendarEntry.AccessLevel != null)
                    {
                        _gDataCalendarEntry.AccessLevel = _calendarEntry.AccessLevel;
                    }
                    if (_calendarEntry.SelfUri.ToString() != null)
                    {
                        _gDataCalendarEntry.SelfUri = _calendarEntry.SelfUri.ToString();
                    }
                    _gDataCalendarEntrys.Add(_gDataCalendarEntry);
                }
                return _gDataCalendarEntrys;
            }



            #endregion CreateCalendarEntrys

            #region AppendCalendarEntrys

            public GDataTypes.GDataCalendarEntrys AppendCalendarEntrys(CalendarEntry calendarEntry, GDataTypes.GDataCalendarEntrys gDataCalendarEntrys)
            {
                var _gDataCalendarEntry = new GDataTypes.GDataCalendarEntry();
                if (calendarEntry.Title.Text != null)
                {
                    _gDataCalendarEntry.Name = calendarEntry.Title.Text;
                }
                if (calendarEntry.Color != null)
                {
                    _gDataCalendarEntry.Color = calendarEntry.Color;
                }
                if (calendarEntry.Summary != null)
                {
                    _gDataCalendarEntry.Description = calendarEntry.Summary.Text;
                }
                if (calendarEntry.Location != null)
                {
                    _gDataCalendarEntry.Location = calendarEntry.Location.ValueString;
                }
                _gDataCalendarEntry.Hidden = calendarEntry.Hidden;
                _gDataCalendarEntry.Selected = calendarEntry.Selected;
                if (calendarEntry.TimeZone != null)
                {
                    _gDataCalendarEntry.TimeZone = calendarEntry.TimeZone;
                }
                if (calendarEntry.AccessLevel != null)
                {
                    _gDataCalendarEntry.AccessLevel = calendarEntry.AccessLevel;
                }
                if (calendarEntry.SelfUri.ToString() != null)
                {
                    _gDataCalendarEntry.SelfUri = calendarEntry.SelfUri.ToString();
                }
                gDataCalendarEntrys.Add(_gDataCalendarEntry);
                return gDataCalendarEntrys;
            }

            #endregion AppendCalendarEntrys

            #region CreateCalendarAclEntrys

            public GDataTypes.GDataCalendarAclEntrys CreateCalendarAclEntrys(AclFeed calendarAclEntrys)
            {
                var _gDataCalendarAclEntrys = new GDataTypes.GDataCalendarAclEntrys();
                foreach (AclEntry _calendarAclEntry in calendarAclEntrys.Entries)
                {
                    var _gDataCalendarAclEntry = new GDataTypes.GDataCalendarAclEntry();
                    _gDataCalendarAclEntry.UserId = _calendarAclEntry.Scope.Value;
                    _gDataCalendarAclEntry.AccessLevel = _calendarAclEntry.Role.Value.Replace("http://schemas.google.com/gCal/2005#","");
                    _gDataCalendarAclEntrys.Add(_gDataCalendarAclEntry);
                }
                return _gDataCalendarAclEntrys;
            }

            #endregion CreateCalendarAclEntrys

            #region CreateCalendarAclEntry

            public GDataTypes.GDataCalendarAclEntry CreateCalendarAclEntry(AclEntry calendarAclEntry)
            {
                var _gDataCalendarAclEntry = new GDataTypes.GDataCalendarAclEntry();
                _gDataCalendarAclEntry.UserId = calendarAclEntry.Scope.Value;
                _gDataCalendarAclEntry.AccessLevel = calendarAclEntry.Role.Value.Replace("http://schemas.google.com/gCal/2005#", "");

                return _gDataCalendarAclEntry;
            }

            #endregion CreateCalendarAclEntry
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

            public GDataTypes.GDataResourceService GetAuthToken(string adminUser, string adminPassword)
            {
                var _uri = new Uri("https://www.google.com/accounts/ClientLogin");

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.ContentType = "application/x-www-form-urlencoded";
                _webRequest.Method = "POST";

                byte[] _bytes = Encoding.UTF8.GetBytes("&Email=" + adminUser + "&Passwd=" + adminPassword.Replace("@", "%40") + "&accountType=HOSTED_OR_GOOGLE&service=apps&source=companyName-applicationName-versionID");
                Stream _OS = null;

                _webRequest.ContentLength = _bytes.Length;   
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);

                _OS.Close();

                WebResponse _webResponse = _webRequest.GetResponse();

                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();


                char[] _deliminator = { ';' };

                var _resE = _result.Replace("Auth=", ";");
                var _resS = _resE.Split(_deliminator);
                var _token = _resS[1].ToString();

                char[] _delimiterChars = { '@' };

                string[] _temp = adminUser.Split(_delimiterChars);
                var _domain = _temp[1];

                
                var _entry = new GDataTypes.GDataResourceService
                {
                    AdminUser = adminUser,
                    Token = _token,
                    Domain = _domain,
                };

                return _entry;
            }

            #endregion GetAuthToken

            #region CreateResourceEntrys

            public GDataTypes.GDataResourceEntrys CreateResourceEntrys(string xml, GDataTypes.GDataResourceService resourceService)
            {
                var _paresdXml = new GDataTypes.ParseXML(xml);


                var _gDataResourceEntrys = new GDataTypes.GDataResourceEntrys();
                var _gdataSingelResourceEntry = new GDataTypes.GDataResourceEntry();
                foreach (var _sEntry in _paresdXml.ListFormat)
                {
                    foreach (var _attribute in _sEntry.at)
                    {

                        if (_attribute.Value == "resourceId" || _attribute.Value == "resourceCommonName" || _attribute.Value == "resourceEmail" || _attribute.Value == "resourceDescription" || _attribute.Value == "resourceType")
                        {

                            if (_attribute.Value == "resourceId")
                            {
                                if (_attribute.NextAttribute.Value != null)
                                {
                                    _gdataSingelResourceEntry.ResourceId = _attribute.NextAttribute.Value;
                                }
                                else
                                {
                                    _gdataSingelResourceEntry.ResourceId = "_EMPTY_";
                                }
                            }
                            if (_attribute.Value == "resourceCommonName")
                            {
                                if (_attribute.NextAttribute.Value != null)
                                {
                                    _gdataSingelResourceEntry.CommonName = _attribute.NextAttribute.Value;
                                }
                                else
                                {
                                    _gdataSingelResourceEntry.CommonName = "_EMPTY_";
                                }
                            }
                            if (_attribute.Value == "resourceEmail")
                            {
                                if (_attribute.NextAttribute.Value != null)
                                {
                                    _gdataSingelResourceEntry.Email = _attribute.NextAttribute.Value;
                                }
                                else
                                {
                                    _gdataSingelResourceEntry.Email = "_EMPTY_";
                                }
                            }
                            if (_attribute.Value == "resourceDescription")
                            {
                                if (_attribute.NextAttribute.Value != null)
                                {
                                    _gdataSingelResourceEntry.Description = _attribute.NextAttribute.Value;
                                }
                                else
                                {
                                    _gdataSingelResourceEntry.Description = "_EMPTY_";
                                }
                            }
                            if (_attribute.Value == "resourceType")
                            {
                                if (_attribute.NextAttribute.Value != null)
                                {
                                    _gdataSingelResourceEntry.Type = _attribute.NextAttribute.Value;
                                }
                                else
                                {
                                    _gdataSingelResourceEntry.Type = "_EMPTY_";
                                }
                            }
                        }

                    }
                    
                    var _gdataResourceEntry = new GDataTypes.GDataResourceEntry();
                    foreach (var _subEntry in _sEntry.sub)
                    {
                        foreach (var _attribute in _subEntry.at)
                        {

                            if (_attribute.Value == "resourceId" || _attribute.Value == "resourceCommonName" || _attribute.Value == "resourceEmail" || _attribute.Value == "resourceDescription" || _attribute.Value == "resourceType")
                            {

                                if (_attribute.Value == "resourceId")
                                {

                                    _gdataResourceEntry.ResourceId = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "resourceCommonName")
                                {

                                    _gdataResourceEntry.CommonName = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "resourceEmail")
                                {

                                    _gdataResourceEntry.Email = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "resourceDescription")
                                {

                                    _gdataResourceEntry.Description = _attribute.NextAttribute.Value;
                                }
                                if (_attribute.Value == "resourceType")
                                {

                                    _gdataResourceEntry.Type = _attribute.NextAttribute.Value;
                                }

                            }

                        }
                    }
                    if (_gdataResourceEntry.ResourceId != null)
                    {
                        _gDataResourceEntrys.Add(_gdataResourceEntry);
                    }
                }
                if (_gdataSingelResourceEntry.ResourceId != null)
                {
                    _gDataResourceEntrys.Add(_gdataSingelResourceEntry);
                }
                return _gDataResourceEntrys;
            }

            #endregion CreateResourceEntrys

            #region NewResource

            public string NewResource(GDataTypes.GDataResourceService resourceService, string resourceId, string resourceType, string resurceDescription)
            {
                var _domain = resourceService.Domain;

                var _uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + _domain);
                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "POST";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + resourceService.Token);
                byte[] _bytes = Encoding.UTF8.GetBytes("<atom:entry xmlns:atom='http://www.w3.org/2005/Atom' xmlns:apps='http://schemas.google.com/apps/2006'><apps:property name='resourceId' value='" + resourceId + "'/><apps:property name='resourceCommonName' value='" + resourceId + "'/><apps:property name='resourceDescription' value='" + resurceDescription + "'/><apps:property name='resourceType' value='" + resourceType + "'/></atom:entry>");
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;   
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);         

                _OS.Close();

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();
                return _result;
            }

            #endregion NewResource

            #region RetriveResource

            public string RetriveResource(GDataTypes.GDataResourceService ResourceService, string ResourceId)
            {
                var _domain = ResourceService.Domain;
                var _uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + _domain + "/" + ResourceId);

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "GET";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + ResourceService.Token);
             

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion RetriveResource

            #region RetriveResource

            public string RetriveAllResources(GDataTypes.GDataResourceService resourceService)
            {
                var _domain = resourceService.Domain;

                var _uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + _domain + "/");
 
                WebRequest webRequest = WebRequest.Create(_uri);

                webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                webRequest.Method = "GET";
                webRequest.Headers.Add("Authorization: GoogleLogin auth=" + resourceService.Token);


                WebResponse webResponse = webRequest.GetResponse();


                if (webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion RetriveAllResources

            #region RemoveResource

            public string RemoveResources(GDataTypes.GDataResourceService resourceService, string resourceID)
            {
                var _domain = resourceService.Domain;

                var _uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + _domain + "/" + resourceID);
                
                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "DELETE";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + resourceService.Token);

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion RemoveResources

            #region SetResource

            public string SetResource(GDataTypes.GDataResourceService resourceService, string resourceID, string resourceDescription, string resourceType)
            {
                var _domain = resourceService.Domain;

                var _uri = new Uri("https://apps-apis.google.com/a/feeds/calendar/resource/2.0/" + _domain + "/" + resourceID);

                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "PUT";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + resourceService.Token);
                byte[] _bytes = Encoding.UTF8.GetBytes("<atom:entry xmlns:atom='http://www.w3.org/2005/Atom'><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceCommonName' value='" + resourceID + "'/><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceDescription' value='" + resourceDescription + "'/><apps:property xmlns:apps='http://schemas.google.com/apps/2006' name='resourceType' value='" + resourceType + "'/></atom:entry>");
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);

                _OS.Close();

                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion SetResources

        }
        #endregion GoogleResourceService
        
        #region GoogleProfileService

        public class GoogleProfileService
        {

            #region GetAuthToken

            public GDataTypes.GDataProfileService GetAuthToken(string adminUser, string adminPassword)
            {
                var _uri = new Uri("https://www.google.com/accounts/ClientLogin");

                WebRequest _webRequest = WebRequest.Create(_uri);

                _webRequest.ContentType = "application/x-www-form-urlencoded";
                _webRequest.Method = "POST";
                _webRequest.Headers.Add("GData-Version: 3.0");
                byte[] _bytes = Encoding.UTF8.GetBytes("&Email=" + adminUser.Replace("@", "%40") + "&Passwd=" + adminPassword + "&accountType=HOSTED&service=cp&source=dgctest.com-GDataCmdLet-v0508");
                Stream _OS = null;

                _webRequest.ContentLength = _bytes.Length;   
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);    

                _OS.Close();


                WebResponse _webResponse = _webRequest.GetResponse();

                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();


                char[] _deliminator = { ';' };

                var _resE = _result.Replace("Auth=", ";");
                var _resS = _resE.Split(_deliminator);
                var _token = _resS[1].ToString();

                char[] _delimiterChars = { '@' };


                string[] _temp = adminUser.Split(_delimiterChars);
                var _domain = _temp[1];


                var _entry = new GDataTypes.GDataProfileService
                {
                    AdminUser = adminUser,
                    Token = _token,
                    Domain = _domain,
                };



                return _entry;
            }

            #endregion GetAuthToken

            #region CreateProfileEntrys

            public GDataTypes.GDataProfileEntrys CreateProfileEntrys(string xml, GDataTypes.GDataProfileService profileService)
            {
                var _paresdXml = new GDataTypes.ParseXML(xml);


                var _profileEntrys = new GDataTypes.GDataProfileEntrys();

                string _profileLink = "http://www.google.com/m8/feeds/profiles/domain/" + profileService.Domain + "/full/";

                foreach (var _sEntry in _paresdXml.ListFormat)
                {
                    if (_sEntry.name == "{http://www.w3.org/2005/Atom}entry")
                    {
                        var _profileEntry = new GDataTypes.GDataProfileEntry();
                        _profileEntry.Domain = profileService.Domain;
                        foreach (var _entry in _sEntry.sub)
                        {
                            if (_entry.name == "{http://www.w3.org/2005/Atom}id")
                            {
                                _profileEntry.UserName = _entry.value.Replace(_profileLink,"");
                            }
                            if (_entry.name == "{http://schemas.google.com/g/2005}structuredPostalAddress")
                            {
                                foreach (var _attribute in _entry.at)
                                {
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                                    {
                                        _profileEntry.HomePostalAddress = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                                    {
                                        _profileEntry.PostalAddress = _entry.value;
                                    }
                                }
                            }
                            if (_entry.name == "{http://schemas.google.com/g/2005}phoneNumber")
                            {
                                foreach (var _attribute in _entry.at)
                                {
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#other")
                                    {
                                        _profileEntry.OtherPhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                                    {
                                        _profileEntry.PhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#mobile")
                                    {
                                        _profileEntry.MobilePhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                                    {
                                        _profileEntry.HomePhoneNumber = _entry.value;
                                    }
                                }
                            }
                        }
                        _profileEntrys.Add(_profileEntry);
                    }
                }


                return _profileEntrys;
            }

            #endregion CreateProfileEntrys

            #region AppendProfileEntrys

            public GDataTypes.GDataProfileEntrys AppendProfileEntrys(string xml, GDataTypes.GDataProfileService profileService, GDataTypes.GDataProfileEntrys profileEntrys)
            {
                var _paresdXml = new GDataTypes.ParseXML(xml);
                string _profileLink = "http://www.google.com/m8/feeds/profiles/domain/" + profileService.Domain + "/full/";

                foreach (var _sEntry in _paresdXml.ListFormat)
                {
                    if (_sEntry.name == "{http://www.w3.org/2005/Atom}entry")
                    {
                        var _profileEntry = new GDataTypes.GDataProfileEntry();
                        _profileEntry.Domain = profileService.Domain;
                        foreach (var _entry in _sEntry.sub)
                        {
                            if (_entry.name == "{http://www.w3.org/2005/Atom}id")
                            {
                                _profileEntry.UserName = _entry.value.Replace(_profileLink, "");
                            }
                            if (_entry.name == "{http://schemas.google.com/g/2005}structuredPostalAddress")
                            {
                                foreach (var _attribute in _entry.at)
                                {
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                                    {
                                        _profileEntry.HomePostalAddress = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                                    {
                                        _profileEntry.PostalAddress = _entry.value;
                                    }
                                }
                            }
                            if (_entry.name == "{http://schemas.google.com/g/2005}phoneNumber")
                            {
                                foreach (var _attribute in _entry.at)
                                {
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#other")
                                    {
                                        _profileEntry.OtherPhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                                    {
                                        _profileEntry.PhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#mobile")
                                    {
                                        _profileEntry.MobilePhoneNumber = _entry.value;
                                    }
                                    if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                                    {
                                        _profileEntry.HomePhoneNumber = _entry.value;
                                    }
                                }
                            }
                        }
                        profileEntrys.Add(_profileEntry);
                    }
                }


                return profileEntrys;
            }

            #endregion AppendProfileEntrys

            #region CreateProfileEntry

            public GDataTypes.GDataProfileEntry CreateProfileEntry(string xml, string id, GDataTypes.GDataProfileService profileService)
            {
                var _paresdXml = new GDataTypes.ParseXML(xml);


                var _profileEntry = new GDataTypes.GDataProfileEntry();

                _profileEntry.Domain = profileService.Domain;
                string _profileLink = "http://www.google.com/m8/feeds/profiles/domain/" + profileService.Domain + "/full/";

                foreach (var _entry in _paresdXml.ListFormat)
                {
                    if (_entry.name == "{http://www.w3.org/2005/Atom}id")
                    {
                        _profileEntry.UserName = _entry.value.Replace(_profileLink, "");
                    }
                    if (_entry.name == "{http://schemas.google.com/g/2005}structuredPostalAddress")
                    {
                        foreach (var _attribute in _entry.at)
                        {
                            if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                            {
                                _profileEntry.HomePostalAddress = _entry.value;
                            }
                            if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                            {
                                _profileEntry.PostalAddress = _entry.value;
                            }
                        }
                    }
                    if (_entry.name == "{http://schemas.google.com/g/2005}phoneNumber")
                    {
                        foreach (var _attribute in _entry.at)
                        {
                            if (_attribute.Value == "http://schemas.google.com/g/2005#other")
                            {
                                _profileEntry.OtherPhoneNumber = _entry.value;
                            }
                            if (_attribute.Value == "http://schemas.google.com/g/2005#work")
                            {
                                _profileEntry.PhoneNumber = _entry.value;
                            }
                            if (_attribute.Value == "http://schemas.google.com/g/2005#mobile")
                            {
                                _profileEntry.MobilePhoneNumber = _entry.value;
                            }
                            if (_attribute.Value == "http://schemas.google.com/g/2005#home")
                            {
                                _profileEntry.HomePhoneNumber = _entry.value;
                            }
                        }
                    }
                }
                return _profileEntry;
            }

            #endregion CreateProfileEntrys

            #region GetProfiles

            public string GetProfiles(GDataTypes.GDataProfileService profileService, string nextPage)
            {

                var _domain = profileService.Domain;

                if (nextPage == "")
                {
                    nextPage = "https://www.google.com/m8/feeds/profiles/domain/" + _domain + "/full";
                }

                var _uri = new Uri(nextPage);

                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "Get";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + profileService.Token);
                _webRequest.Headers.Add("GData-Version: 3.0");


                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion GetProfiles

            #region GetProfile

            public string GetProfile(GDataTypes.GDataProfileService profileService, string id)
            {

                var Domain = profileService.Domain;

                var _uri = new Uri("https://www.google.com/m8/feeds/profiles/domain/" + Domain + "/full/" + id);



                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "Get";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + profileService.Token);
                _webRequest.Headers.Add("GData-Version: 3.0");


                WebResponse _webResponse = _webRequest.GetResponse();


                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;
            }

            #endregion GetProfile

            #region SetProfile

            private string res;
            private XElement elem;
            private XNamespace ns;
            private string newXml;

            public string SetProfile(GDataTypes.GDataProfileService profileService, string id, string postalAddress, string phoneNumber, string mobilePhoneNumber, string otherPhoneNumber, string homePostalAddress, string homePhoneNumber)
            {

                var _domain = profileService.Domain;

                var _uri = new Uri("http://www.google.com/m8/feeds/profiles/domain/" + _domain + "/full/" + id);

                var _googleProfileService = new GoogleProfileService();
                res = _googleProfileService.GetProfile(profileService, id);



                elem = XElement.Parse(res.ToString().Trim());

                ns = "http://schemas.google.com/g/2005";

                string _formatedRes = res.Replace("\"", "'");

                if (postalAddress != null)
                {
                    if (_formatedRes.Contains("gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'"))
                    {

                        foreach (var _element in elem.Elements(ns + "structuredPostalAddress"))
                        {
                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#work")
                            {

                                _element.SetElementValue(ns + "formattedAddress", postalAddress);
                            }


                        }
                        newXml = elem.ToString();
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#work'><gd:formattedAddress>" + postalAddress + "</gd:formattedAddress></gd:structuredPostalAddress></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                if (homePostalAddress != null)
                {
                    if (_formatedRes.Contains("gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#home'"))
                    {

                        foreach (var _element in elem.Elements(ns + "structuredPostalAddress"))
                        {

                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#home")
                            {
                                _element.SetElementValue(ns + "formattedAddress", homePostalAddress);
                            }

                        }
                        newXml = elem.ToString();
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:structuredPostalAddress rel='http://schemas.google.com/g/2005#home'><gd:formattedAddress>" + homePostalAddress + "</gd:formattedAddress></gd:structuredPostalAddress></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                if (phoneNumber != null)
                {
                    if (res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#work'"))
                    {

                        foreach (var _element in elem.Elements(ns + "phoneNumber"))
                        {

                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#work")
                            {
                                _element.SetValue(phoneNumber);
                            }
                            newXml = elem.ToString();
                        }
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#work' primary='true'>" + phoneNumber + "</gd:phoneNumber></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                if (mobilePhoneNumber != null)
                {
                    if (res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#mobile'"))
                    {
                        foreach (var _element in elem.Elements(ns + "phoneNumber"))
                        {
                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#mobile")
                            {
                                _element.SetValue(mobilePhoneNumber);
                            }
                        }
                        newXml = elem.ToString();
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#mobile'>" + mobilePhoneNumber + "</gd:phoneNumber></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                if (otherPhoneNumber != null)
                {
                    if (res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#other'"))
                    {
                        foreach (var _element in elem.Elements(ns + "phoneNumber"))
                        {
                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#other")
                            {
                                _element.SetValue(otherPhoneNumber);
                            }
                            newXml = elem.ToString();
                        }
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#other'>" + mobilePhoneNumber + "</gd:phoneNumber></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                if (homePhoneNumber != null)
                {
                    if (res.Contains("phoneNumber rel='http://schemas.google.com/g/2005#home'"))
                    {
                        foreach (var _element in elem.Elements(ns + "phoneNumber"))
                        {
                            if (_element.FirstAttribute.Value == "http://schemas.google.com/g/2005#home")
                            {
                                _element.SetValue(homePhoneNumber);
                            }
                            newXml = elem.ToString();
                        }
                    }
                    else
                    {
                        newXml = elem.ToString().Replace("</entry>", "<gd:phoneNumber rel='http://schemas.google.com/g/2005#home'>" + homePhoneNumber + "</gd:phoneNumber></entry>");
                        elem = XElement.Parse(newXml.ToString().Trim());
                    }
                }

                WebRequest _webRequest = WebRequest.Create(_uri);
                _webRequest.ContentType = "application/atom+xml; charset=UTF-8";
                _webRequest.Method = "PUT";
                _webRequest.Headers.Add("Authorization: GoogleLogin auth=" + profileService.Token);
                _webRequest.Headers.Add("GData-Version: 3.0");

                var _post = newXml;

                byte[] _bytes = Encoding.UTF8.GetBytes(_post);
                Stream _OS = null;
                _webRequest.ContentLength = _bytes.Length;
                _OS = _webRequest.GetRequestStream();
                _OS.Write(_bytes, 0, _bytes.Length);

                _OS.Close();

                WebResponse _webResponse = _webRequest.GetResponse();

                if (_webResponse == null)
                {
                    throw new Exception("WebResponse is null");
                }
                StreamReader _SR = new StreamReader(_webResponse.GetResponseStream());

                var _result = _SR.ReadToEnd().Trim();

                return _result;

            }

            #endregion SetProfile

        }
        
        #endregion GoogleProfileService
    }
}

