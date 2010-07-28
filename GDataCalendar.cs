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

                var _CalendarService = new CalendarService("Calendar");
                _CalendarService.setUserCredentials(_AdminUser, _AdminPassword);
               
                WriteObject(_CalendarService);
               
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;
            /*
            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;
            */

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            #endregion Parameters

            protected override void ProcessRecord()
            {
                


                    var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                    var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                    
                    var _CalendarQuery = new CalendarQuery();
                    _CalendarQuery.Uri = new Uri(_SelfUri);
                    try
                    {
                        var _CalendarFeed = (CalendarFeed)_CalendarService.CalendarService.Query(_CalendarQuery);


                        foreach (var _Entry in _CalendarFeed.Entries)
                        {
                            if (_Entry.SelfUri.ToString() == _SelfUri)
                            {

                                try
                                {
                                    _Entry.Delete();
                                    WriteObject(_Entry);

                                }
                                catch (Exception _Exception)
                                {
                                    WriteObject(_Exception);
                                }
                            }


                        }
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { _CalendarName = value; }
            }
            private string _CalendarName;

            #endregion Parameters


            protected override void ProcessRecord()
            {

           

                    var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                    var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                    
                    var _CalendarQuery = new CalendarQuery();
                    _CalendarQuery.Uri = new Uri("https://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/allcalendars/full");
                   
                    try
                    {

                        var _CalendarFeed = _CalendarService.CalendarService.Query(_CalendarQuery);

                        if (_CalendarName == null)
                        {
                            var _CalendarEntrys = _DgcGoogleCalendarService.CreateCalendarEntrys(_CalendarFeed);
                            WriteObject(_CalendarEntrys, true);
                        }
                        else
                        {

                            var _CalendarSelection = from _Selection in _CalendarFeed.Entries
                                                     where _Selection.Title.Text.ToString() == _CalendarName
                                                     select _Selection;

                            var _CalendarEntrys = new GDataTypes.GDataCalendarEntrys();
                            foreach (CalendarEntry _Entry in _CalendarSelection)
                            { 
                                _CalendarEntrys = _DgcGoogleCalendarService.AppendCalendarEntrys(_Entry, _CalendarEntrys);
                            }
                            WriteObject(_CalendarEntrys, true);
                        }
     
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { _CalendarName = value; }
            }
            private string _CalendarName;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { _Description = value; }
            }
            private string _Description;

            [Parameter(
            Mandatory = false,
            HelpMessage = "TimeZone, Long format time zone ID (not PST, but America/Los_Angeles.) "
            )]
            [ValidateNotNullOrEmpty]
            public string TimeZone
            {
                get { return null; }
                set { _TimeZone = value; }
            }
            private string _TimeZone;


            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                var _Calendar = new CalendarEntry();
                _Calendar.Title.Text = _CalendarName;

                if (_Description != null)
                {
                    _Calendar.Summary.Text = _Description;
                }
                if (_TimeZone != null)
                {
                    _Calendar.TimeZone = _TimeZone;
                }
                try
                {
                    if (_CalendarService.Oauth == null)
                    {
                        throw new Exception("User -ConsumerKey/-ConsumerSecret in New-GdataService");
                    }
                    if (_CalendarService.OauthCalendarService == null)
                    {
                        throw new Exception("User -ConsumerKey/-ConsumerSecret in New-GdataService");
                    }
                    var _PostUri = new Uri("http://www.google.com/calendar/feeds/default/owncalendars/full");
                    var _OAuth2LeggedAuthenticator = new OAuth2LeggedAuthenticator("GDataCmdLet", _CalendarService.Oauth.ConsumerKey, _CalendarService.Oauth.ConsumerSecret, _ID, _Domain);
                    var _OauthUri = _OAuth2LeggedAuthenticator.ApplyAuthenticationToUri(_PostUri);
                    
                    var _CreatedCalendar = (CalendarEntry)_CalendarService.OauthCalendarService.Insert(_OauthUri, _Calendar);
                    var _CalendarEntry = _DgcGoogleCalendarService.CreateCalendarEntry(_CreatedCalendar);
                    WriteObject(_CalendarEntry);

                } catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;
            /*
            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;
            */
            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { _Description = value; }
            }
            private string _Description;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Calendar Name"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarName
            {
                get { return null; }
                set { _CalendarName = value; }
            }
            private string _CalendarName;

            [Parameter(
            Mandatory = false,
            HelpMessage = "TimeZone, Long format time zone ID (not PST, but America/Los_Angeles.) "
            )]
            [ValidateNotNullOrEmpty]
            public string TimeZone
            {
                get { return null; }
                set { _TimeZone = value; }
            }
            private string _TimeZone;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);


                var _CalendarQuery = new CalendarQuery();
                //_CalendarQuery.Uri = new Uri("http://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/allcalendars/full");
                _CalendarQuery.Uri = new Uri(_SelfUri);
                try
                {
                    var _CalendarFeed = (CalendarFeed)_CalendarService.CalendarService.Query(_CalendarQuery);


                    foreach (CalendarEntry _Entry in _CalendarFeed.Entries)
                    {
                        if (_Entry.SelfUri.ToString() == _SelfUri)
                        {
                            try
                            {

                                if (_CalendarName != null)
                                {
                                    _Entry.Title.Text = _CalendarName;
                                }
                                if (_Description != null)
                                {
                                    _Entry.Summary.Text = _Description;
                                }
                                if (_TimeZone != null)
                                {
                                    _Entry.TimeZone = _TimeZone;
                                }

                                _Entry.Update();
                                
                                _CalendarFeed = (CalendarFeed)_CalendarService.CalendarService.Query(_CalendarQuery);
                                foreach (CalendarEntry _ResultEntry in _CalendarFeed.Entries)
                                {
                                    if (_ResultEntry.SelfUri.ToString() == _SelfUri)
                                    {
                                        var _CalendarEntry = _DgcGoogleCalendarService.CreateCalendarEntry(_ResultEntry);
                                        WriteObject(_CalendarEntry);
                                
                                    }
                                }

                            }
                            catch (Exception _Exception)
                            {
                                WriteObject(_Exception);
                            }
                        }


                    }
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                var _Query = new CalendarQuery();
                _Query.Uri = new Uri(_SelfUri);


                try
                {
                    var _Entry = _CalendarService.CalendarService.Query(_Query);

                    //WriteObject(_Entry);

                    var _Links = _Entry.Entries[0].Links;

                    if (_Links == null) 
                    {
                        throw new Exception("AclFeed new null");
                    }


                    var _LinkSelection = from _Selection in _Links
                                         where _Selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                         select _Selection;

                    foreach (var _Link in _LinkSelection)
                    {

                        var _AclQuery = new AclQuery(_Link.HRef.ToString());
                        var _Feed = _CalendarService.CalendarService.Query(_AclQuery);
                        var _CalendarAclEntrys = _DgcGoogleCalendarService.CreateCalendarAclEntrys(_Feed);

                        WriteObject(_CalendarAclEntrys);
                    }
                    
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Role"
            )]
            [ValidateNotNullOrEmpty]
            public string Role
            {
                get { return null; }
                set { _Role = value; }
            }
            private string _Role;


            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                var _Query = new CalendarQuery();
                _Query.Uri = new Uri(_SelfUri);


                try
                {
                    var _Entry = _CalendarService.CalendarService.Query(_Query);

                    //WriteObject(_Entry);

                    var _Links = _Entry.Entries[0].Links;

                    if (_Links == null)
                    {
                        throw new Exception("AclFeed new null");
                    }

                    var _LinkSelection = from _Selection in _Links
                                         where _Selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                         select _Selection;



                    foreach (var _Link in _LinkSelection)
                    {

                        var _AclEntry = new AclEntry();
                        _AclEntry.Scope = new AclScope();
                        _AclEntry.Scope.Type = AclScope.SCOPE_USER;
                        _AclEntry.Scope.Value = _ID;

                        if (_Role.ToUpper() == "FREEBUSY")
                        {
                            _AclEntry.Role = AclRole.ACL_CALENDAR_FREEBUSY;
                        }
                        else if (_Role.ToUpper() == "READ")
                        {
                            _AclEntry.Role = AclRole.ACL_CALENDAR_READ;
                        }
                        else if (_Role.ToUpper() == "EDITOR")
                        {
                            _AclEntry.Role = AclRole.ACL_CALENDAR_EDITOR;
                        }
                        else if (_Role.ToUpper() == "OWNER")
                        {
                            _AclEntry.Role = AclRole.ACL_CALENDAR_OWNER;
                        }
                        else
                        {
                            throw new Exception("-Role needs a FREEBUSY/READ/EDITOR/OWNER parameter");
                        }

                        var _AclUri = new Uri(_Link.HRef.ToString());
                        var _AlcEntry = _CalendarService.CalendarService.Insert(_AclUri, _AclEntry) as AclEntry;
                        var _CalendarAclEntry = _DgcGoogleCalendarService.CreateCalendarAclEntry(_AclEntry);
                        WriteObject(_CalendarAclEntry);
                        
                    }
                    

                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
                set { _CalendarService = value; }
            }
            private GDataTypes.GDataService _CalendarService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "SelfUri"
            )]
            [ValidateNotNullOrEmpty]
            public string SelfUri
            {
                get { return null; }
                set { _SelfUri = value; }
            }
            private string _SelfUri;

            [Parameter(
            Mandatory = true,
            HelpMessage = "User ID, test@domain.com"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService.CalendarService);

                var _Query = new CalendarQuery();
                _Query.Uri = new Uri(_SelfUri);


                try
                {
                    var _Entry = _CalendarService.CalendarService.Query(_Query);

                    //WriteObject(_Entry);

                    var _Links = _Entry.Entries[0].Links;

                    if (_Links == null)
                    {
                        throw new Exception("AclFeed new null");
                    }


                    var _LinkSelection = from _Selection in _Links
                                where _Selection.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList"
                                select _Selection;

                    foreach (var _Link in _LinkSelection)
                    {

                        var _AclQuery = new AclQuery(_Link.HRef.ToString());
                        var _Feed = _CalendarService.CalendarService.Query(_AclQuery);

                        var _FeedSelection = from AclEntry _Selection in _Feed.Entries
                                    where _Selection.Scope.Value.ToString() == _ID
                                    select _Selection;



                        foreach (AclEntry _AclEntry in _FeedSelection)
                        {
                                _AclEntry.Delete();
                                WriteObject(_ID);
                        }
                    
                    }   
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }

            }

        }

        #endregion Remove-GDataCalendarAcl

    }

}
