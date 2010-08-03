using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Xml;
using System.Xml.Linq;
using Google.Contacts;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Apps;
using Google.GData.Extensions.Apps;
using Google.GData.Calendar;
using Google.GData.Apps.GoogleMailSettings;


namespace Microsoft.PowerShell.GData
{
    public class GDataTypes
    {
        #region Service

        public class GDataService
        {
            public AppsService AppsService;
            public CalendarService CalendarService;
            public CalendarService OauthCalendarService;
            public GoogleMailSettingsService GoogleMailSettingsService;
            public GDataTypes.GDataProfileService ProfileService;
            public GDataTypes.GDataResourceService ResourceService;
            public ContactsService ContactsService;
            public ContactsService OauthContactsService;
            public GDataTypes.Oauth Oauth;
        }

        #endregion Service

        #region Group


        public class GDataGroupEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataGroupEntry GDataGroupEntry)
            {
                List.Add(GDataGroupEntry);

            }
        }

        public class GDataGroupEntry
        {
            public string GroupId;
            public string GroupName;
            public string EmailPermission;
            public string Description;
            public string SelfUri;
        }

        public class GDataGroupMemberEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataGroupMemberEntry GDataGroupMemberEntry)
            {
                List.Add(GDataGroupMemberEntry);

            }
        }

        public class GDataGroupMemberEntry
        {
            public string MemberId;
            public string MemberType;
            public string DirectMember;
        }

        public class GDataGroupOwnerEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataGroupOwnerEntry GDataGroupOwnerEntry)
            {
                List.Add(GDataGroupOwnerEntry);

            }
        }

        public class GDataGroupOwnerEntry
        {
            public string MemberId;
            public string MemberType;
        }

        #endregion Group

        #region Contact

        public class GDataContactEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataContactEntry GDataContactEntry)
            {
                List.Add(GDataContactEntry);

            }
        }

        public class GDataContactEntry
        {
            public string Name;
            public string Email;
            public string PhoneNumber;
            public string PostalAddress;
            public string City;
            public string SelfUri;
        }

        public class Oauth
        {
            public string ConsumerKey;
            public string ConsumerSecret;
        }

        #endregion Contact

        #region Calendar

        public class GDataCalendarEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataCalendarEntry GDataCalendarEntry)
            {
                List.Add(GDataCalendarEntry);

            }
        }

        public class GDataCalendarEntry
        {
            public string Name;
            public string AccessLevel;
            public string Description;
            public string Location;
            public string TimeZone;
            public string Color;
            public bool Hidden;
            public bool Selected;
            public string SelfUri;
        }

        public class GDataCalendarAclEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataCalendarAclEntry GDataCalendarAclEntry)
            {
                List.Add(GDataCalendarAclEntry);

            }
        }

        public class GDataCalendarAclEntry
        {
            public string UserId;
            public string AccessLevel;
        }

        #endregion Contact

        #region Resource

        public class GDataResourceEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataResourceEntry ResourceEntry)
            {
                List.Add(ResourceEntry);

            }
        }

        public class GDataResourceEntry
        {

            public string ResourceId;
            public string CommonName;
            public string Email;
            public string Description;
            public string Type;

        }


        public class GDataResourceService
        {
            public string AdminUser { get; set; }
            public string Token { get; set; }
            public string Domain { get; set; }
        }

        #endregion Resource

        #region XmlTypes

        public class XmlReturn
        {
            public string name { get; set; }
            public string value { get; set; }
            public IEnumerable<XAttribute> at;
            public IEnumerable<XmlSubReturn> sub;
        }

        public class XmlSubReturn
        {
            public string name { get; set; }
            public string value { get; set; }
            public IEnumerable<XAttribute> at;
            public IEnumerable<XElement> sub;
        }

        #endregion XmlTypes

        #region Profile

        public class GDataProfileEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataProfileEntry ProfileEntry)
            {
                List.Add(ProfileEntry);

            }
        }

        public class GDataProfileEntry
        {

            public string UserName;
            public string Domain;
            public string PhoneNumber;
            public string MobilePhoneNumber;
            public string OtherPhoneNumber;
            public string HomePhoneNumber;
            public string HomePostalAddress;
            public string PostalAddress;
        }

        public class GDataProfileService
        {
            public string AdminUser { get; set; }
            public string Token { get; set; }
            public string Domain { get; set; }
        }

        #endregion Profile

        #region User

        public class GDataUserEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataUserEntry UserEntry)
            {
                List.Add(UserEntry);

            }
        }

        public class GDataUserEntry
        {
            public string UserName;
            public string GivenName;
            public string FamilyName;
            public bool susspended;
            public bool Admin;
            public bool AgreedToTerms;
            public bool ChangePasswordAtNextLogin;
            public int Limit;
            public string SelfUri;
        }

        public class GDataUserAliasEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataUserAliasEntry GDataUserAliasEntry)
            {
                List.Add(GDataUserAliasEntry);

            }
        }

        public class GDataUserAliasEntry : System.Collections.CollectionBase
        {
            public void Add(GDataAliasEntry GDataAliasEntry)
            {
                List.Add(GDataAliasEntry);

            }
        }

        public class GDataAliasEntry
        {

            public string UserName;
            public string aliasEmail;
        }



        #endregion User

        #region MailSettings

        public class GDataSenderAddress
        {
            public string Address;
            public string ReplyTo;
            public string Name;
            public string Default;
        }

        public class GDataImap
        {
            public string Enabled;
        }

        public class GDataPop3
        {
            public string Action;
            public string Enabled;
            public string EnabledFor;
        }

        #endregion Mailsettings

        public class ParseXML
        {
            public IEnumerable<XmlReturn> ListFormat;
            public ParseXML(string XMLString)
            {
                XElement Elem = XElement.Parse(XMLString);
                XNamespace Ns = "http://schemas.google.com/apps/2006";


                ListFormat = from c in Elem.Elements()
                             select new XmlReturn
                             {
                                 name = c.Name.ToString(),
                                 value = c.Value.ToString(),
                                 at = c.Attributes(),
                                 sub = from e in c.Elements()
                                       select new XmlSubReturn
                                       {
                                           name = e.Name.ToString(),
                                           value = e.Value.ToString(),
                                           at = e.Attributes(),
                                       }
                             };

            }
        }
    }
}
