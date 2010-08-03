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
using Google.GData.Apps.Groups;
using Google.GData.Calendar;
using Google.AccessControl;
using Google.GData.AccessControl;
using System.Linq;

namespace Microsoft.PowerShell.GData
{

    public class Calendar
    {
        #region New-GDataCalendarService
        /*
        [Cmdlet(VerbsCommon.New, "GDataCalendarService")]
        public class NewGDataCalendarService : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GoogleApps admin user, admin@domain.com",
            HelpMessageBaseName = "GoogleApps admin user, admin@domain.com"
            )]
            [ValidateNotNullOrEmpty]
            public string AdminUsername
            {
                get { return null; }
                set { _AdminUser = value; }
            }
            private string _AdminUser;

            [Parameter(
               Mandatory = true,
               HelpMessage = "GoogleApps admin password"
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

                //var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                //var _Domain = _DgcGoogleAppsService.GetDomain(_AdminUser);

                var service = new CalendarService("Calendar");
                service.setUserCredentials(_AdminUser, _AdminPassword);
               
                WriteObject(service);
               
            }


        }
        */
        #endregion New-GDataCalendarService

        #region Remove-GDataCalendar

        [Cmdlet(VerbsCommon.Remove, "GDataCalendar")]
        public class RemoveGDataCalendar : Cmdlet
        {

            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "CalendarService, new-GdataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);
                var _calendarQuery = new CalendarQuery();
                _calendarQuery.Uri = new Uri(selfUri);
                try
                {
                    var _calendarFeed = (CalendarFeed)service.CalendarService.Query(_calendarQuery);
                    foreach (var _entry in _calendarFeed.Entries)
                    {
                        if (_entry.SelfUri.ToString() == selfUri)
                        {
                            try
                            {
                                _entry.Delete();
                                WriteObject(_entry);
                            }
                            catch (Exception _exception)
                            {
                                WriteObject(_exception);
                            }
                        }
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
           }
        }

        #endregion Remove-GDataCalendar

        #region Get-GDataCalendar

        [Cmdlet(VerbsCommon.Get, "GDataCalendar")]
        public class GetGDataCalendar : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { _calendarName = value; }
            }
            private string _calendarName;

            #endregion Parameters


            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);
                var _calendarQuery = new CalendarQuery();
                _calendarQuery.Uri = new Uri("https://www.google.com/calendar/feeds/" + id + "@" + _domain + "/allcalendars/full");
                   
                try
                {
                    var _calendarFeed = service.CalendarService.Query(_calendarQuery);

                    if (_calendarName == null)
                    {
                        var _calendarEntrys = dgcGoogleCalendarService.CreateCalendarEntrys(_calendarFeed);
                        WriteObject(_calendarEntrys, true);
                    }
                    else
                    {
                        var _calendarSelection = from _selection in _calendarFeed.Entries
                                                    where _selection.Title.Text.ToString() == _calendarName
                                                    select _selection;

                        var _calendarEntrys = new GDataTypes.GDataCalendarEntrys();
                        foreach (CalendarEntry _entry in _calendarSelection)
                        { 
                            _calendarEntrys = dgcGoogleCalendarService.AppendCalendarEntrys(_entry, _calendarEntrys);
                        }
                        WriteObject(_calendarEntrys, true);
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }

        }

        #endregion Get-GDataCalendar

        #region New-GDataCalendar

        [Cmdlet(VerbsCommon.New, "GDataCalendar")]
        public class NewGDataCalendar : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { calendarName = value; }
            }
            private string calendarName;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { description = value; }
            }
            private string description;

            [Parameter(
            Mandatory = false,
            HelpMessage = "TimeZone, Long format time zone ID (not PST, but America/Los_Angeles.) "
            )]
            [ValidateNotNullOrEmpty]
            public string TimeZone
            {
                get { return null; }
                set { timeZone = value; }
            }
            private string timeZone;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);

                var _calendar = new CalendarEntry();
                _calendar.Title.Text = calendarName;

                if (description != null)
                {
                    _calendar.Summary.Text = description;
                }
                if (timeZone != null)
                {
                    _calendar.TimeZone = timeZone;
                }
                try
                {
                    if (service.Oauth == null)
                    {
                        throw new Exception("User -ConsumerKey/-ConsumerSecret in New-GdataService");
                    }
                    if (service.OauthCalendarService == null)
                    {
                        throw new Exception("User -ConsumerKey/-ConsumerSecret in New-GdataService");
                    }
                    var _postUri = new Uri("http://www.google.com/calendar/feeds/default/owncalendars/full");
                    var _oAuth2LeggedAuthenticator = new OAuth2LeggedAuthenticator("GDataCmdLet", service.Oauth.ConsumerKey, service.Oauth.ConsumerSecret, id, _domain);
                    var _oAuthUri = _oAuth2LeggedAuthenticator.ApplyAuthenticationToUri(_postUri);
                    
                    var _createdCalendar = (CalendarEntry)service.OauthCalendarService.Insert(_oAuthUri, _calendar);
                    var _calendarEntry = dgcGoogleCalendarService.CreateCalendarEntry(_createdCalendar);
                    WriteObject(_calendarEntry);

                } catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion New-GDataCalendar

        #region Set-GDataCalendar

        [Cmdlet(VerbsCommon.Set, "GDataCalendar")]
        public class SetGDataCalendar : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { description = value; }
            }
            private string description;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { calendarName = value; }
            }
            private string calendarName;

            [Parameter(
            Mandatory = false,
            HelpMessage = "TimeZone, Long format time zone ID (not PST, but America/Los_Angeles.) "
            )]
            [ValidateNotNullOrEmpty]
            public string TimeZone
            {
                get { return null; }
                set { timeZone = value; }
            }
            private string timeZone;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);

                var _calendarQuery = new CalendarQuery();
                _calendarQuery.Uri = new Uri(selfUri);
                try
                {
                    var _calendarFeed = (CalendarFeed)service.CalendarService.Query(_calendarQuery);
                    foreach (CalendarEntry _entry in _calendarFeed.Entries)
                    {
                        if (_entry.SelfUri.ToString() == selfUri)
                        {
                            try
                            {
                                if (calendarName != null)
                                {
                                    _entry.Title.Text = calendarName;
                                }
                                if (description != null)
                                {
                                    _entry.Summary.Text = description;
                                }
                                if (timeZone != null)
                                {
                                    _entry.TimeZone = timeZone;
                                }

                                _entry.Update();
                                
                                _calendarFeed = (CalendarFeed)service.CalendarService.Query(_calendarQuery);
                                foreach (CalendarEntry _resultEntry in _calendarFeed.Entries)
                                {
                                    if (_resultEntry.SelfUri.ToString() == selfUri)
                                    {
                                        var _calendarEntry = dgcGoogleCalendarService.CreateCalendarEntry(_resultEntry);
                                        WriteObject(_calendarEntry);
                                
                                    }
                                }
                            }
                            catch (Exception _exception)
                            {
                                WriteObject(_exception);
                            }
                        }
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }

            }

        }

        #endregion Set-GDataCalendar

        #region Get-GDataCalendarAcl

        [Cmdlet(VerbsCommon.Get, "GDataCalendarAcl")]
        public class GetGDataCalendarAcl : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();  
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);
                var _query = new CalendarQuery();
                _query.Uri = new Uri(selfUri);

                try
                {
                    var _entry = service.CalendarService.Query(_query);
                    var _links = _entry.Entries[0].Links;

                    if (_links == null) 
                    {
                        throw new Exception("AclFeed new null");
                    }


                    var _linkSelection = from _selection in _links
                                         where _selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                         select _selection;

                    foreach (var _link in _linkSelection)
                    {
                        var _aclQuery = new AclQuery(_link.HRef.ToString());
                        var _feed = service.CalendarService.Query(_aclQuery);
                        var _calendarAclEntrys = dgcGoogleCalendarService.CreateCalendarAclEntrys(_feed);

                        WriteObject(_calendarAclEntrys);
                    }
                    
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Get-GDataCalendarAcl

        #region Add-GDataCalendarAcl

        [Cmdlet(VerbsCommon.Add, "GDataCalendarAcl")]
        public class AddGDataCalendarAcl : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Role"
            )]
            [ValidateNotNullOrEmpty]
            public string Role
            {
                get { return null; }
                set { _role = value; }
            }
            private string _role;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);

                var _query = new CalendarQuery();
                _query.Uri = new Uri(selfUri);

                try
                {
                    var _entry = service.CalendarService.Query(_query);
                    var _links = _entry.Entries[0].Links;

                    if (_links == null)
                    {
                        throw new Exception("AclFeed new null");
                    }

                    var _linkSelection = from _selection in _links
                                         where _selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                         select _selection;



                    foreach (var _link in _linkSelection)
                    {

                        var _aclEntry = new AclEntry();
                        _aclEntry.Scope = new AclScope();
                        _aclEntry.Scope.Type = AclScope.SCOPE_USER;
                        _aclEntry.Scope.Value = id;

                        if (_role.ToUpper() == "FREEBUSY")
                        {
                            _aclEntry.Role = AclRole.ACL_CALENDAR_FREEBUSY;
                        }
                        else if (_role.ToUpper() == "READ")
                        {
                            _aclEntry.Role = AclRole.ACL_CALENDAR_READ;
                        }
                        else if (_role.ToUpper() == "EDITOR")
                        {
                            _aclEntry.Role = AclRole.ACL_CALENDAR_EDITOR;
                        }
                        else if (_role.ToUpper() == "OWNER")
                        {
                            _aclEntry.Role = AclRole.ACL_CALENDAR_OWNER;
                        }
                        else
                        {
                            throw new Exception("-Role needs a FREEBUSY/READ/EDITOR/OWNER parameter");
                        }

                        var _aclUri = new Uri(_link.HRef.ToString());
                        var _alcEntry = service.CalendarService.Insert(_aclUri, _aclEntry) as AclEntry;
                        var _calendarAclEntry = dgcGoogleCalendarService.CreateCalendarAclEntry(_aclEntry);
                        WriteObject(_calendarAclEntry);   
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Add-GDataCalendarAcl

        #region Remove-GDataCalendarAcl

        [Cmdlet(VerbsCommon.Remove, "GDataCalendarAcl")]
        public class RemoveGDataCalendarAcl : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GDataCalendar, new-GDataCalendarService"
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { selfUri = value; }
            }
            private string selfUri;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID, test@domain.com"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            private Dgc.GoogleCalendarsService dgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleCalendarService.GetDomain(service.CalendarService);
                var _query = new CalendarQuery();
                _query.Uri = new Uri(selfUri);

                try
                {
                    var _entry = service.CalendarService.Query(_query);
                    var _links = _entry.Entries[0].Links;

                    if (_links == null)
                    {
                        throw new Exception("AclFeed new null");
                    }


                    var _linkSelection = from _selection in _links
                                where _selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                select _selection;

                    foreach (var _Link in _linkSelection)
                    {

                        var _aclQuery = new AclQuery(_Link.HRef.ToString());
                        var _Feed = service.CalendarService.Query(_aclQuery);

                        var _feedSelection = from AclEntry _selection in _Feed.Entries
                                    where _selection.Scope.Value.ToString() == id
                                    select _selection;



                        foreach (AclEntry _aclEntry in _feedSelection)
                        {
                                _aclEntry.Delete();
                                WriteObject(id);
                        }
                    
                    }   
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }

            }

        }

        #endregion Remove-GDataCalendarAcl

    }

}
