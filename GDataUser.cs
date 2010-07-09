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

namespace Microsoft.PowerShell.GData
{
    public class User
    {
        #region New-GDataUserService

        [Cmdlet(VerbsCommon.New, "GDataUserService")]
        public class NewGDataUserService : Cmdlet
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

                try
                {
                    var _UserService = new AppsService(_Domain, _AdminUser, _AdminPassword);
                    WriteObject(_UserService);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }


        }

        #endregion New-GDataUserService

        #region Remove-GDatauser

        [Cmdlet(VerbsCommon.Remove, "GDataUser")]
        public class RemoveGDataUser : Cmdlet
        {

            #region Parameters
            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = true
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
                    _UserService.DeleteUser(_ID);
                    WriteObject(_ID);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = false
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
                        var _Feed = _UserService.RetrieveAllUsers();
                        WriteObject(_Feed.Entries);
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
                        var _Entry = _UserService.RetrieveUser(_ID);
                        WriteObject(_Entry);
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string GivenName
            {
                get { return null; }
                set { _GivenName = value; }
            }
            private string _GivenName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string FamilyName
            {
                get { return null; }
                set { _FamilyName = value; }
            }
            private string _FamilyName;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Passsword
            {
                get { return null; }
                set { _Password = value; }
            }
            private string _Password;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string NewID
            {
                get { return null; }
                set { _NewId = value; }
            }
            private string _NewId;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                var _Entry = _UserService.RetrieveUser(_ID);

                if (_FamilyName != null)
                {
                    _Entry.Name.FamilyName = _FamilyName;
                }
                if (_GivenName != null)
                {
                    _Entry.Name.GivenName = _GivenName;
                }
                if (_Password != null)
                {
                    _Entry.Login.Password = _Password;
                }
                if (_NewId != null)
                {
                    _Entry.Login.UserName = _NewId;
                }

                try
                {
                    var _update = _UserService.UpdateUser(_Entry);
                    WriteObject(_update);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string GivenName
            {
                get { return null; }
                set { _GivenName = value; }
            }
            private string _GivenName;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string FamilyName
            {
                get { return null; }
                set { _FamilyName = value; }
            }
            private string _FamilyName;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Passsword
            {
                get { return null; }
                set { _Password = value; }
            }
            private string _Password;

            #endregion Parameters

            protected override void ProcessRecord()
            {
                try
                {
                    var _Entry = _UserService.CreateUser(_ID, _GivenName, _FamilyName, _Password);
                    WriteObject(_Entry);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = false
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
                        var _Feed = _UserService.RetrieveAllNicknames();
                        WriteObject(_Feed.Entries);
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
                        var _Feed = _UserService.RetrieveNicknames(_ID);

                        if (_Feed.Entries.Count == 0)
                        {
                            Except throw newion("No NickNames found!");
                        }
                        WriteObject(_Feed);

                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { _ID = value; }
            }
            private string _ID;


            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string NickName
            {
                get { return null; }
                set { _NickName = value; }
            }
            private string _NickName;


            #endregion Parameters


            protected override void ProcessRecord()
            {
                try
                {
                    var _Entry = _UserService.CreateNickname(_ID, _NickName);
                    WriteObject(_Entry);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
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
            public AppsService UserService
            {
                get { return null; }
                set { _UserService = value; }
            }
            private AppsService _UserService;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string NickName
            {
                get { return null; }
                set { _NickName = value; }
            }
            private string _NickName;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                try
                {
                    _UserService.DeleteNickname(_NickName);
                    WriteObject(_NickName);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion Remove-GDataUserNickname

    }

}
