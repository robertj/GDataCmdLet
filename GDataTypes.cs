using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Xml;
using System.Xml.Linq;


namespace Microsoft.PowerShell.GData
{
    public class GDataTypes
    {

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

        public class ProfileEntrys : System.Collections.CollectionBase
        {
            public void Add(ProfileEntry ProfileEntry)
            {
                List.Add(ProfileEntry);

            }
        }

        public class ProfileEntry
        {

            public string User;
            public string Domain;
            public string PhoneNumber;
            public string MobilePhoneNumber;
            public string OtherPhoneNumber;
            public string HomePhoneNumber;
            public string HomePostalAddress;
            public string PostalAddress;
        }

        public class ProfileService
        {
            public string AdminUser { get; set; }
            public string Token { get; set; }
            public string Domain { get; set; }
        }

        #endregion Profile

        #region user

        public class GDataUserEntrys : System.Collections.CollectionBase
        {
            public void Add(GDataUserEntry UserEntry)
            {
                List.Add(UserEntry);

            }
        }

        public class GDataUserEntry
        {
            public string SelfUri;
            public string FamilyName;
            public string GivenName;
            public string UserName;
            public bool susspended;
            public bool Admin;
            public bool AgreedToTerms;
            public bool ChangePasswordAtNextLogin;
            public int Limit;
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



        #endregion user

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
