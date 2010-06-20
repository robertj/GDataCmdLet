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


namespace Microsoft.PowerShell.GData
{

    public class Group
    {        
        #region New-GDataGroupService

        [Cmdlet(VerbsCommon.New, "GDataGroupService")]
        public class NewGDataGroupService : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GoogleApps admin user, admin@domain.com"
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

                var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                var _Domain = _DgcGoogleAppsService.GetDomain(_AdminUser);


                var _GroupService = new AppsService(_Domain, _AdminUser, _AdminPassword);
                WriteObject(_GroupService);
            }


        }

        #endregion New-GDataGroupService

        #region Remove-GDataGroup

        [Cmdlet(VerbsCommon.Remove, "GDataGroup")]
        public class RemoveGDataGroup : Cmdlet
        {

            #region Parameters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Goup ID"
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
                try
                {
                    _GroupService.Groups.DeleteGroup(_ID);
                    WriteObject(_ID);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }


        }

        #endregion Remove-GDataGroup

        #region Get-GDataGroup

        [Cmdlet(VerbsCommon.Get, "GDataGroup")]
        public class GetGDataGroup : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Group ID"
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

                if (_ID == null)
                {
                    try
                    {
                        var feed = _GroupService.Groups.RetrieveAllGroups();
                        WriteObject(feed.Entries);
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
                    }
                }
                else
                {
                    try
                    {
                        var entry = _GroupService.Groups.RetrieveGroup(_ID);
                        WriteObject(entry);
                    }
                    catch (Exception _Exception) 
                    { 
                        WriteObject(_Exception); 
                    }
                }

            }

        }

        #endregion Get-GDataGroup

        #region Get-GDataGroupMember

        [Cmdlet(VerbsCommon.Get, "GDataGroupMember")]
        public class GetGDataGroupMember : Cmdlet
        {
            #region Parameters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
            )]
            [ValidateNotNullOrEmpty]
            public string Id
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                try
                {
                    var _Feed = _GroupService.Groups.RetrieveAllMembers(_ID);
                    WriteObject(_Feed.Entries);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion Get-GDataGroupMember

        #region Add-GDataGroupMember

        [Cmdlet(VerbsCommon.Add, "GDataGroupMember")]
        public class AddGDataGroupMember : Cmdlet
        {
            #region Paradmeters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { _UserID = value; }
            }
            private string _UserID;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                try 
                {
                    var _Entry = _GroupService.Groups.AddMemberToGroup(_UserID, _ID);
                    WriteObject(_Entry);
                }catch (Exception _Exception) 
                {
                    WriteObject(_Exception);
                }

            }

        }

        #endregion Add-GDataGroupMember

        #region Remove-GDataGroupMember

        [Cmdlet(VerbsCommon.Remove, "GDataGroupMember")]
        public class RemoveGDataGroupMember : Cmdlet
        {
            #region Paradmeters


            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, New-GDataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { _UserID = value; }
            }
            private string _UserID;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                try 
                {
                    _GroupService.Groups.RemoveMemberFromGroup(_UserID, _ID);
                    WriteObject(_UserID);
                }
                catch (Exception _Exception) 
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion Remove-GDataGroupMember

        #region Set-GDataGroup

        [Cmdlet(VerbsCommon.Set, "GDataGroup")]
        public class SetGDataGroup : Cmdlet
        {

            #region Parameters

            [Parameter(
                  Mandatory = true,
                  HelpMessage = "GroupService, New-GDataGroupService"
            )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Group name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name = value; }
            }
            private string _Name;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Group description"
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
            HelpMessage = "Group EmailPermission, Owner, Member, Domain, Anyone"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailPermission
            {
                get { return null; }
                set { _EmailPermission = value; }
            }
            private string _EmailPermission;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                AppsExtendedEntry _Group = _GroupService.Groups.RetrieveGroup(_ID);

                foreach (Google.GData.Extensions.Apps.PropertyElement property in _Group.Properties)
                {

                    if (_Name == null && property.Name == "groupName")
                    {
                        _Name = property.Value;
                    }
                    else if (_Description == null && property.Name == "description")
                    {
                        _Description = property.Value;
                    }
                    else if (_EmailPermission == null && property.Name == "emailPermission")
                    {
                        _EmailPermission = property.Value;
                    }

                }
                try
                {
                    var _Entry = _GroupService.Groups.UpdateGroup(_ID, _Name, _Description, _EmailPermission);
                    WriteObject(_Entry);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion Set-GDataGroup

        #region New-GDataGroup

        [Cmdlet(VerbsCommon.New, "GDataGroup")]
        public class NewGDataGroup : Cmdlet
        {


            #region Parameters


            [Parameter(
         Mandatory = true,
         HelpMessage = "GroupService, New-GDataGroupService"
         )]
            [ValidateNotNullOrEmpty]
            public AppsService GroupService
            {
                get { return null; }
                set { _GroupService = value; }
            }
            private AppsService _GroupService;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Group name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name = value; }
            }
            private string _Name;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { _Description = value; }
            }
            private string _Description;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group EmailPermission, Owner, Member, Domain, Anyone"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailPermission
            {
                get { return null; }
                set { _EmailPermission = value; }
            }
            private string _EmailPermission;

            #endregion Parameters


            protected override void ProcessRecord()
            {

                try
                {

                    var _Entry = _GroupService.Groups.CreateGroup(_ID, _Name, _Description, _EmailPermission);
                    WriteObject(_Entry);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }



        #endregion New-GDataGroup

    }

}
