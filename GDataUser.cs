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
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.PowerShell.GData
{
    public class User
    {
        #region Remove-GDatauser

        [Cmdlet(VerbsCommon.Remove, "GDataUser")]
        public class RemoveGDataUser : Cmdlet
        {
            #region Parameters
            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            protected override void ProcessRecord()
            {
                try
                {
                    service.AppsService.DeleteUser(id);
                    WriteObject(id);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion Remove-GDataUser

        #region Get-GDataUser

        [Cmdlet(VerbsCommon.Get, "GDataUser")]
        public class GetGDataUser : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
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
                        var _feed = service.AppsService.RetrieveAllUsers();
                        var _userEntrys = dgcGoogleAppsService.CreateUserEntrys(_feed);
                        WriteObject(_userEntrys, true);
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
                        var _entry = service.AppsService.RetrieveUser(id);
                        var _userEntry = dgcGoogleAppsService.CreateUserEntry(_entry);
                        WriteObject(_userEntry);
                    }
                    catch (AppsException _exception)
                    {
                        WriteObject(_exception,true);
                    }
                }

            }

        }

        #endregion Get-GDataUser

        #region Set-GDataUser

        [Cmdlet(VerbsCommon.Set, "GDataUser")]
        public class SetGDataUser : Cmdlet
        {

            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string GivenName
            {
                set { givenName = value; }
            }
            private string givenName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string FamilyName
            {
                set { familyName = value; }
            }
            private string familyName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Passsword
            {
                set { password = value; }
            }
            private string password;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string NewID
            {
                set { newID = value; }
            }
            private string newID;

            [Parameter(Mandatory = false)]
            [ValidateNotNullOrEmpty]
            public string ChangePassNextLogon
            {
                set { changePassNextLogon = value; }
            }
            private string changePassNextLogon;
            private bool chPass;

            [Parameter(Mandatory = false)]
            public string IsAdmin
            {
                set { isAdmin = value; }
            }
            private string isAdmin;
            private bool isAdminBool;

            [Parameter(Mandatory = false)]
            public string Suspended
            {
                set { suspended = value; }
            }
            private string suspended;
            private bool suspendedBool;
          
            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                var _entry = service.AppsService.RetrieveUser(id);

                if (familyName != null)
                {
                    _entry.Name.FamilyName = familyName;
                }
                if (givenName != null)
                {
                    _entry.Name.GivenName = givenName;
                }
                if (password != null)
                {
                    _entry.Login.Password = password;
                }
                if (newID != null)
                {
                    _entry.Login.UserName = newID;
                }
                
                try
                {

                    if (changePassNextLogon != null)
                    {
                        changePassNextLogon = changePassNextLogon.ToLower();
                        if (changePassNextLogon == "true")
                        {
                            chPass = true;
                        }
                        else if (changePassNextLogon == "false")
                        {
                            chPass = false;
                        }
                        else
                        {
                            throw new Exception("-ChangePassNextLogon needs a true or false statement");
                        }
                        _entry.Login.ChangePasswordAtNextLogin = chPass;
                    }

                    if (isAdmin != null)
                    {
                        isAdmin = isAdmin.ToLower();
                        if (isAdmin == "true")
                        {
                            isAdminBool = true;
                        }
                        else if (isAdmin == "false")
                        {
                            isAdminBool = false;
                        }
                        else
                        {
                            throw new Exception("-IsAdmin needs a true or false statement");
                        }
                        _entry.Login.Admin = isAdminBool;
                    }

                    if (suspended != null)
                    {
                        suspended = suspended.ToLower();
                        if (suspended == "true")
                        {
                            suspendedBool = true;
                        }
                        else if (suspended == "false")
                        {
                            suspendedBool = false;
                        }
                        else
                        {
                            throw new Exception("-Suspended needs a true or false statement");
                        }
                        _entry.Login.Suspended = suspendedBool;
                    }

                    
                    var _userEntry = dgcGoogleAppsService.CreateUserEntry(service.AppsService.UpdateUser(_entry));
                    WriteObject(_userEntry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }

        }

        #endregion Set-GDataUser

        #region New-GDataUser

        [Cmdlet(VerbsCommon.New, "GDataUser")]
        public class NewGDataUser : Cmdlet
        {

            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string GivenName
            {
                set { givenName = value; }
            }
            private string givenName;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string FamilyName
            {
                set { familyName = value; }
            }
            private string familyName;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Passsword
            {
                set { password = value; }
            }
            private string password;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try
                {
                    var _userEntry = dgcGoogleAppsService.CreateUserEntry(service.AppsService.CreateUser(id, givenName, familyName, password));
                    WriteObject(_userEntry);
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }
        }

        #endregion New-GDataUser

        #region Get-GDataUserNickName
        [Cmdlet(VerbsCommon.Get, "GDataUserNickName")]
        public class GetGDataUserNickName : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                set { legacy = value; }
            }
            private bool legacy;

            #endregion Parameters

            private string nextPage;
            private GDataTypes.ParseXML parseXML;
            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                if (id == null)
                {
                    try 
                    {
                        if (!legacy == true)
                        {
                            nextPage = "";
                            var _xml = dgcGoogleAppsService.RetriveAllUserAlias(service.AppsService, nextPage);
                            var _userAliasEntry = dgcGoogleAppsService.CreateUserAliasEntry(_xml);
                            parseXML = new GDataTypes.ParseXML(_xml.ToString());
                            
                            while (nextPage != null)
                            {
                                nextPage = null;
                                foreach (var _elements in parseXML.ListFormat)
                                {
                                    foreach (var _attribute in _elements.at)
                                    {
                                        if (_attribute.Value == "next")
                                        {
                                            nextPage = _attribute.NextAttribute.NextAttribute.Value;
                                        }
                                    }
                                }
                                if (nextPage != null)
                                {
                                    _xml = dgcGoogleAppsService.RetriveAllUserAlias(service.AppsService, nextPage);
                                    parseXML = new GDataTypes.ParseXML(_xml.ToString());
                                    _userAliasEntry = dgcGoogleAppsService.AppendUserAliasEntry(_xml, _userAliasEntry);
                                }
                            }

                            WriteObject(_userAliasEntry,true);
                        }
                        else
                        {
                            var _feed = service.AppsService.RetrieveAllNicknames();
                            WriteObject(_feed.Entries);
                        }
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
                        if (!legacy == true)
                        {

                            /*
                            if (!id.Contains("@"))
                            {
                                throw new Exception("-ID must contain Domain, user@domain.com");
                            }
                            */
                            var _userAliasEntry = dgcGoogleAppsService.CreateUserAliasEntry(dgcGoogleAppsService.RetriveUserAlias(id, service.AppsService));
                            WriteObject(_userAliasEntry,true);
                        }
                        else
                        {
                            var _feed = service.AppsService.RetrieveNicknames(id);
                            WriteObject(_feed);
                        }
             
                    }
                    catch (Exception _exception)
                    {
                        WriteObject(_exception);
                    }
                }
            }
        }

        #endregion Get-GDataUserNickname

        #region Add-GDataUserNickName

        [Cmdlet(VerbsCommon.Add, "GDataUserNickName")]
        public class AddGDataUserNickName : Cmdlet
        {
            #region Parameters

            [Parameter(
                      Mandatory = true
                      )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                set { id = value; }
            }
            private string id;


            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string NickName
            {
                set { nickName = value; }
            }
            private string nickName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                set { legacy = value; }
            }
            private bool legacy;


            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try
                {
                    if (!legacy == true)
                    {
                        if (!nickName.Contains("@"))
                        {
                            throw new Exception("-NickName must contain Domain, user@domain.com");
                        }
                        /*
                        if (!id.Contains("@"))
                        {
                            throw new Exception("-ID must contain Domain, user@domain.com");
                        }
                        */
                        var _userAliasEntry = dgcGoogleAppsService.CreateUserAliasEntry(dgcGoogleAppsService.CreateUserAlias(id, service.AppsService, nickName));
                        WriteObject(_userAliasEntry);
                    }
                    else
                    {
                        var _entry = service.AppsService.CreateNickname(id, nickName);
                        WriteObject(_entry);
                    }

                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }

        }

        #endregion Add-GDataUserNickname

        #region Remove-GDataUserNickName

        [Cmdlet(VerbsCommon.Remove, "GDataUserNickName")]
        public class RemoveGDataUserNickName : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string NickName
            {
                set { nickName = value; }
            }
            private string nickName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                set { legacy = value; }
            }
            private bool legacy;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                try
                {
                    if (!legacy == true)
                    {
                        if(!nickName.Contains("@"))
                        {
                            throw new Exception("-NickName must contain EmailDomain, user@domain.com");
                        }

                        dgcGoogleAppsService.RemoveUserAlias(service.AppsService, nickName);
                        WriteObject(nickName);
                    }
                    else
                    {
                        service.AppsService.DeleteNickname(nickName);
                        WriteObject(nickName);
                    }
                }
                catch (Exception _exception)
                {
                    WriteObject(_exception);
                }
            }

        }
        #endregion Remove-GDataUserNickname
    }
}
