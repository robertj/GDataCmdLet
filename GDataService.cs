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
            public PSCredential Credentials
            {
                get { return null; }
                set { credentials = value; }
            }
            private PSCredential credentials;
            
            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string ConsumerKey
            {
                get { return null; }
                set { consumerKey = value; }
            }
            private string consumerKey;

            [Parameter(
               Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string ConsumerSecret
            {
                get { return null; }
                set { consumerSecret = value; }
            }
            private string consumerSecret;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            private GDataTypes.GDataService service = new GDataTypes.GDataService();
            private string adminUser;
            private string adminPassword;
            protected override void ProcessRecord()
            {
                adminUser = credentials.UserName;
                adminPassword = new Dgc.ConvertToUnsecureString(credentials.Password).PlainString;
                var _domain = dgcGoogleAppsService.GetDomain(adminUser);

                try
                {
                    //AppsService
                    service.AppsService = new AppsService(_domain, adminUser, adminPassword);
                    
                    //CalendarService
                    var _calendarService = new CalendarService("Calendar");
                    _calendarService.setUserCredentials(adminUser, adminPassword);                    
                    service.CalendarService = _calendarService;
                    
                    //OauthCalendarService
                    if (consumerKey != null)
                    {
                        if (consumerSecret == null)
                        {
                            throw new Exception("-ConsumerSecret can't be null");
                        }
                        var _oauthCalendarService = new CalendarService("Calendar");
                        var _oauth = new GDataTypes.Oauth();
                        _oauth.ConsumerKey = consumerKey;
                        _oauth.ConsumerSecret = consumerSecret;
                        service.Oauth = _oauth;
                        GOAuthRequestFactory _requestFactory = new GOAuthRequestFactory("cl", "GDataCmdLet");
                        _requestFactory.ConsumerKey = _oauth.ConsumerKey;
                        _requestFactory.ConsumerSecret = _oauth.ConsumerSecret;
                        _oauthCalendarService.RequestFactory = _requestFactory;
                        service.OauthCalendarService = _oauthCalendarService;
                    }

                    //MailSettingsService
                    var _googleMailSettingsService = new GoogleMailSettingsService(_domain, "GMailSettingsService");
                    _googleMailSettingsService.setUserCredentials(adminUser, adminPassword);
                    service.GoogleMailSettingsService = _googleMailSettingsService;
                    
                    //ProfileService
                    var _dgcGoogleProfileService = new Dgc.GoogleProfileService();
                    service.ProfileService = _dgcGoogleProfileService.GetAuthToken(adminUser, adminPassword);
                    
                    //ResourceService
                    var _dgcGoogleResourceService = new Dgc.GoogleResourceService();
                    service.ResourceService = _dgcGoogleResourceService.GetAuthToken(adminUser, adminPassword);
                    
                    //ContactsService
                    var _contactService = new ContactsService("GData");
                    _contactService.setUserCredentials(adminUser, adminPassword);
                    service.ContactsService = _contactService;

                    WriteObject(service);
                }
                catch (AppsException _exception)
                {
                    WriteObject(_exception, true);
                }
            }
        }
        #endregion New-GDataUserService

    }

}
