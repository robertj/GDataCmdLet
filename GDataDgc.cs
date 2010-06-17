using System;
using System.Diagnostics;
using System.Management.Automation;
using System.ComponentModel;
using System.Collections.Generic;
using Google.Contacts;
using Google.GData.Contacts;

namespace Microsoft.PowerShell.GData
{

    public class Dgc
    {

        #region GoogleAppService

        public class GoogleAppService
        {
            public string GetDomain(string AdminUser)
            {
                char[] delimiterChars = { '@' };
                string[] temp = AdminUser.Split(delimiterChars);
                var Domain = temp[1];
                return Domain;

            }
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

    }
}

