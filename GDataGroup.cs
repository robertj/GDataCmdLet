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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Goup ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            protected override void ProcessRecord()
            {
                try
                {
                    service.AppsService.Groups.DeleteGroup(id);
                    WriteObject(id);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Group ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                if (id == null)
                {
                    try
                    {
                        var _entrys = dgcGoogleAppsService.CreateGroupEntrys(service.AppsService.Groups.RetrieveAllGroups());
                        WriteObject(_entrys, true);
                    }
                    catch (Exception _exception)
                    {
                        WriteObject(_exception);
                    }
                }
                else
                {
                    try
                    {
                        var _entry = dgcGoogleAppsService.CreateGroupEntry(service.AppsService.Groups.RetrieveGroup(id));
                        WriteObject(_entry);
                    }
                    catch (Exception _exception) 
                    { 
                        WriteObject(_exception); 
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
            )]
            [ValidateNotNullOrEmpty]
            public string Id
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try
                {
                    var _entry = dgcGoogleAppsService.CreateGroupMemberEntrys(service.AppsService.Groups.RetrieveAllMembers(id));
                    WriteObject(_entry, true);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { userID = value; }
            }
            private string userID;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try 
                {
                    var _entry = dgcGoogleAppsService.CreateGroupMemberEntry(service.AppsService.Groups.AddMemberToGroup(userID, id));
                    WriteObject(_entry);
                }catch (Exception _exception) 
                {
                    WriteObject(_exception);
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { userID = value; }
            }
            private string userID;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try 
                {
                    service.AppsService.Groups.RemoveMemberFromGroup(userID, id);
                    WriteObject(userID);
                }
                catch (Exception _exception) 
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Remove-GDataGroupMember

        #region Get-GDataGroupOwner

        [Cmdlet(VerbsCommon.Get, "GDataGroupOwner")]
        public class GetGDataGroupOwner : Cmdlet
        {
            #region Paradmeters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
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
            HelpMessage = "Group ID"
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                var _domain = dgcGoogleAppsService.GetDomainFromAppService(service.AppsService);

                try
                {
                    
                    var _Feed = service.AppsService.Groups.RetrieveGroupOwners(id);
                    var _GroupOwnerEntrys = dgcGoogleAppsService.CreateGroupOwnerEntrys(_Feed);
                    WriteObject(_GroupOwnerEntrys, true);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }

            }

        }

        #endregion Get-GDataGroupOwner

        #region Add-GDataGroupOwner

        [Cmdlet(VerbsCommon.Add, "GDataGroupOwner")]
        public class AddGDataGroupOwner : Cmdlet
        {
            #region Paradmeters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
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
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { userID = value; }
            }
            private string userID;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {

                var _domain = service.AppsService.Domain.ToString();

                try
                {
                    var _userCheck = service.AppsService.Groups.RetrieveMember(userID, id);

                    if (_userCheck == null)
                    {
                        throw new Exception(userID + " is note meber of " + id);
                    }

                    var _entry = dgcGoogleAppsService.CreateGroupOwnerEntry(service.AppsService.Groups.AddOwnerToGroup(userID, id));
                    WriteObject(_entry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Add-GDataGroupOwner

        #region Remove-GDataGroupOwner

        [Cmdlet(VerbsCommon.Remove, "GDataGroupOwner")]
        public class RemoveGDataGroupOwner : Cmdlet
        {
            #region Paradmeters

            [Parameter(
            Mandatory = true,
            HelpMessage = "GroupService, new-GdataGroupService"
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
            HelpMessage = "Group ID"
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
            HelpMessage = "Username"
            )]
            [ValidateNotNullOrEmpty]
            public string UserID
            {
                get { return null; }
                set { userID = value; }
            }
            private string userID;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                var _domain = service.AppsService.Domain.ToString();
                try
                {
                    service.AppsService.Groups.RemoveOwnerFromGroup(userID, id);
                    WriteObject(id);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Remove-GDataGroupOwner

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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Group name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { name = value; }
            }
            private string name;

            [Parameter(
            Mandatory = false,
            HelpMessage = "Group description"
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
            HelpMessage = "Group EmailPermission, Owner, Member, Domain, Anyone"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailPermission
            {
                get { return null; }
                set { emailPermission = value; }
            }
            private string emailPermission;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                AppsExtendedEntry _group = service.AppsService.Groups.RetrieveGroup(id);

                foreach (Google.GData.Extensions.Apps.PropertyElement property in _group.Properties)
                {

                    if (name == null && property.Name == "groupName")
                    {
                        name = property.Value;
                    }
                    else if (description == null && property.Name == "description")
                    {
                        description = property.Value;
                    }
                    else if (emailPermission == null && property.Name == "emailPermission")
                    {
                        emailPermission = property.Value;
                    }

                }
                try
                {
                    var _entry = dgcGoogleAppsService.CreateGroupEntry(service.AppsService.Groups.UpdateGroup(id, name, description, emailPermission));
                    WriteObject(_entry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group ID"
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
            HelpMessage = "Group name"
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { name = value; }
            }
            private string name;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group description"
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { description = value; }
            }
            private string description;

            [Parameter(
            Mandatory = true,
            HelpMessage = "Group EmailPermission, Owner, Member, Domain, Anyone"
            )]
            [ValidateNotNullOrEmpty]
            public string EmailPermission
            {
                get { return null; }
                set { emailPermission = value; }
            }
            private string emailPermission;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try
                {
                    var _entry = dgcGoogleAppsService.CreateGroupEntry(service.AppsService.Groups.CreateGroup(id, name, description, emailPermission));
                    WriteObject(_entry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion New-GDataGroup
    }
}
