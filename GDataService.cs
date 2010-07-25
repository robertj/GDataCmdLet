using System;
using System.Diagnostics;
using System.Management.Automation;
using System.ComponentModel;
using Google.Contacts;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using System.Collections.Generic;
using Google.GData.Apps;
using Google.GData.Apps.GoogleMailSettings;
using Google.GData.Extensions.Apps;
using Google.GData.Calendar;
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.PowerShell.GData
{
    public class Service
    {
        #region New-GDataService

        [Cmdlet(VerbsCommon.New, "GDataService")]
        public class NewGDataService : Cmdlet
        {

            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string AdminUsername
            {
                get { return null; }
                set { _AdminUser = value; }
            }
            private string _AdminUser;

            [Parameter(
               Mandatory = true
            )]
            [ValidateNotNullOrEmpty]

            public string AdminPassword
            {
                get { return null; }
                set { _AdminPassword = value; }
            }
            private string _AdminPassword;

            #endregion Parameters

            protected override void ProcessRecord()
            {

                var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                var _Domain = _DgcGoogleAppsService.GetDomain(_AdminUser);
                var _GDataService = new GDataTypes.GDataService();

                try
                {
                    //AppsService
                    _GDataService.AppsService = new AppsService(_Domain, _AdminUser, _AdminPassword);
                    //CalendarService
                    var _CalendarService = new CalendarService("Calendar");
                    _CalendarService.setUserCredentials(_AdminUser, _AdminPassword);
                    _GDataService.CalendarService = _CalendarService;
                    //MailSettingsService
                    var _GoogleMailSettingsService = new GoogleMailSettingsService(_Domain, "GMailSettingsService");
                    _GoogleMailSettingsService.setUserCredentials(_AdminUser, _AdminPassword);
                    _GDataService.GoogleMailSettingsService = _GoogleMailSettingsService;
                    //ProfileService
                    var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                    _GDataService.ProfileService = _DgcGoogleProfileService.GetAuthToken(_AdminUser, _AdminPassword);
                    //ResourceService
                    var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                    _GDataService.ResourceService = _DgcGoogleResourceService.GetAuthToken(_AdminUser, _AdminPassword);
                    //ContactsService
                    var _ContactService = new ContactsService("GData");
                    _ContactService.setUserCredentials(_AdminUser, _AdminPassword);
                    _GDataService.ContactsService = _ContactService;


                    WriteObject(_GDataService);
                }
                catch (AppsException _Exception)
                {
                    WriteObject(_Exception, true);
                }
            }


        }

        #endregion New-GDataUserService

    }

}
