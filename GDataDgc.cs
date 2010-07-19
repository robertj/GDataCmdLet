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
using System.Net;
using Google.GData.Calendar;


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

            public string GetDomainFromAppService(AppsService UserService)
            {
                string Domain = UserService.Domain;
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

