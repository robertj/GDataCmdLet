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


namespace Microsoft.PowerShell.GData
{

    public class Calendar
    {
        #region New-GDataCalendarService

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
            public CalendarService CalendarService
            {
                get { return null; }
                set { _CalendarService = value; }
            }
            private CalendarService _CalendarService;

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
                    var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService);

                    
                    var _CalendarQuery = new CalendarQuery();
                    _CalendarQuery.Uri = new Uri("http://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/owncalendars/full");
                    try
                    {
                        var _CalendarFeed = (CalendarFeed)_CalendarService.Query(_CalendarQuery);


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
            public CalendarService CalendarService
            {
                get { return null; }
                set { _CalendarService = value; }
            }
            private CalendarService _CalendarService;

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
            HelpMessage = "Calendar ID"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarID
            {
                get { return null; }
                set { _CalendarID = value; }
            }
            private string _CalendarID;

            #endregion Parameters


            protected override void ProcessRecord()
            {

           

                    var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                    var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService);

                    
                    var _CalendarQuery = new CalendarQuery();
                    _CalendarQuery.Uri = new Uri("https://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/owncalendars/full");
                   
                    try
                    {

                        var _CalendarFeed = _CalendarService.Query(_CalendarQuery);

                        if (_CalendarID == null)
                        {
                            WriteObject(_CalendarFeed.Entries, true);
                        }
                        else
                        {
                            foreach (var _Entry in _CalendarFeed.Entries)
                            {
                                if (_Entry.Title.Text.ToString() == _CalendarID) 
                                {
                                    WriteObject(_Entry, true);   
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
            public CalendarService CalendarService
            {
                get { return null; }
                set { _CalendarService = value; }
            }
            private CalendarService _CalendarService;

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
            HelpMessage = "CalendarID"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarID
            {
                get { return null; }
                set { _CalendarID = value; }
            }
            private string _CalendarID;

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

            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService);

                var _Calendar = new CalendarEntry();
                _Calendar.Title.Text = _CalendarID;
                _Calendar.Summary.Text = _Description;
                
                //_Calendar.TimeZone = "America/Los_Angeles";
                //_Calendar.Hidden = false;
                //_Calendar.Color = "#2952A3";
                //_Calendar.Location = new Where("", "", "Oakland");

                Uri postUri = new Uri("http://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/owncalendars/full");


                CalendarEntry _CreatedCalendar = (CalendarEntry) _CalendarService.Insert(postUri, _Calendar);


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
            public CalendarService CalendarService
            {
                get { return null; }
                set { _CalendarService = value; }
            }
            private CalendarService _CalendarService;

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
            HelpMessage = "CalendarID"
            )]
            [ValidateNotNullOrEmpty]
            public string CalendarID
            {
                get { return null; }
                set { _CalendarID = value; }
            }
            private string _CalendarID;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                var _DgcGoogleCalendarService = new Dgc.GoogleCalendarsService();
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService);


                var _CalendarQuery = new CalendarQuery();
                _CalendarQuery.Uri = new Uri("http://www.google.com/calendar/feeds/" + _ID + "@" + _Domain + "/owncalendars/full");
                try
                {
                    var _CalendarFeed = (CalendarFeed)_CalendarService.Query(_CalendarQuery);


                    foreach (var _Entry in _CalendarFeed.Entries)
                    {
                        if (_Entry.SelfUri.ToString() == _SelfUri)
                        {
                            try
                            {

                                if (_CalendarID != null)
                                {
                                    _Entry.Title.Text = _CalendarID;
                                }
                                if (_Description != null)
                                {
                                    _Entry.Summary.Text = _Description;
                                }

                                var _Result = _Entry.Update();
                                WriteObject(_Result);

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
            public CalendarService CalendarService
            {
                get { return null; }
                set { _CalendarService = value; }
            }
            private CalendarService _CalendarService;

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
                var _Domain = _DgcGoogleCalendarService.GetDomain(_CalendarService);

                var _Query = new CalendarQuery();
                _Query.Uri = new Uri(_SelfUri);


                try
                {
                    var _Entry = _CalendarService.Query(_Query);

                    //WriteObject(_Entry);

                    var _Links = _Entry.Entries[0].Links;

                    if (_Links == null) 
                    {
                        throw new Exception("AclFeed new null");
                    }
                        
                            
                        foreach ( var _Link in _Links ) {

                            if (_Link.Rel.ToString() == "http://schemas.google.com/acl/2007#accessControlList")
                            {

                                var _AclQuery = new AclQuery(_Link.HRef.ToString());
                                var _Feed = _CalendarService.Query(_AclQuery);
                                WriteObject(_Feed.Entries);

                            }   
                        }
                    
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }


            }

        }

        #endregion Get-GDataCalendarAcl

    }

}
