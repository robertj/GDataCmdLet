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
                catch (AppsException _Exception)
                {
                    WriteObject(_Exception,true);
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
                var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                if (_ID == null)
                {
                    try
                    {
                        var _Feed = _UserService.RetrieveAllUsers();
                        var _GDataUserEntrys = _DgcGoogleAppsService.CreateUserEntrys(_Feed);
                        WriteObject(_GDataUserEntrys, true);
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
                        var _UserEntry = _UserService.RetrieveUser(_ID);

                        var _GdataUserEntry = _DgcGoogleAppsService.CreateUserEntry(_UserEntry);
                        WriteObject(_GdataUserEntry);
                    }
                    catch (AppsException _Exception)
                    {
                        WriteObject(_Exception,true);
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

            [Parameter(Mandatory = false)]
            [ValidateNotNullOrEmpty]
            public string ChangePassNextLogon
            {
                get { return null; }
                set { _ChangePassNextLogon = value; }
            }
            private string _ChangePassNextLogon;
            private bool _ChPass;

            [Parameter(Mandatory = false)]
            public string IsAdmin
            {
                get { return null; }
                set { _IsAdmin = value; }
            }
            private string _IsAdmin;
            private bool _IsAdminBool;
            [Parameter(Mandatory = false)]
            public string Suspended
            {
                get { return null; }
                set { _Suspended = value; }
            }
            private string _Suspended;
            private bool _SuspendedBool;
          
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

                    if (_ChangePassNextLogon != null)
                    {
                        _ChangePassNextLogon = _ChangePassNextLogon.ToLower();
                        if (_ChangePassNextLogon == "true")
                        {
                            _ChPass = true;
                        }
                        else if (_ChangePassNextLogon == "false")
                        {
                            _ChPass = false;
                        }
                        else
                        {
                            throw new Exception("-ChangePassNextLogon needs a true or false statement");
                        }
                        _Entry.Login.ChangePasswordAtNextLogin = _ChPass;
                    }

                    if (_IsAdmin != null)
                    {
                        _IsAdmin = _IsAdmin.ToLower();
                        if (_IsAdmin == "true")
                        {
                            _IsAdminBool = true;
                        }
                        else if (_IsAdmin == "false")
                        {
                            _IsAdminBool = false;
                        }
                        else
                        {
                            throw new Exception("-IsAdmin needs a true or false statement");
                        }
                        _Entry.Login.Admin = _IsAdminBool;
                    }

                    if (_Suspended != null)
                    {
                        _Suspended = _Suspended.ToLower();
                        if (_Suspended == "true")
                        {
                            _SuspendedBool = true;
                        }
                        else if (_Suspended == "false")
                        {
                            _SuspendedBool = false;
                        }
                        else
                        {
                            throw new Exception("-Suspended needs a true or false statement");
                        }
                        _Entry.Login.Suspended = _SuspendedBool;
                    }
                
                    var _UserEntry = _UserService.UpdateUser(_Entry);
                    var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                    var _GdataUserEntry = _DgcGoogleAppsService.CreateUserEntry(_UserEntry);
                    WriteObject(_GdataUserEntry);
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
                    var _UserEntry = _UserService.CreateUser(_ID, _GivenName, _FamilyName, _Password);
                    var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                    var _GdataUserEntry = _DgcGoogleAppsService.CreateUserEntry(_UserEntry);
                    WriteObject(_GdataUserEntry);
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

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                //get { return null; }
                set { _Legacy = value; }
            }
            private bool _Legacy;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                if (_ID == null)
                {
                    try 
                    {
                        if (!_Legacy == true)
                        {
                            var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                            var _Xml = _DgcGoogleAppsService.RetriveAllUserAlias(_UserService);
                            
                            var _UserAliasEntry = _DgcGoogleAppsService.CreateUserAliasEntry(_Xml);
                            
                            
                            WriteObject(_UserAliasEntry,true);
                        }
                        else
                        {
                            var _Feed = _UserService.RetrieveAllNicknames();
                            WriteObject(_Feed.Entries);
                        }
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
                        if (!_Legacy == true)
                        {
                            if (!_ID.Contains("@"))
                            {
                                throw new Exception("-ID must contain EmailDomain, user@domain.com");
                            }
                            var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                            var _Xml = _DgcGoogleAppsService.RetriveUserAlias(_ID, _UserService);

                            var _UserAliasEntry = _DgcGoogleAppsService.CreateUserAliasEntry(_Xml);

                            //var _Reader = XmlReader.Create(_Entry);
                            /*
                            XmlDocument _XmlDoc = new XmlDocument();
                            _XmlDoc.InnerXml = _Xml;
                            XmlElement _Entry = _XmlDoc.DocumentElement;
                            */

                            WriteObject(_UserAliasEntry,true);

                        }
                        else
                        {
                            var _Feed = _UserService.RetrieveNicknames(_ID);
                            WriteObject(_Feed);
                        }
             
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

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                //get { return null; }
                set { _Legacy = value; }
            }
            private bool _Legacy;


            #endregion Parameters


            protected override void ProcessRecord()
            {
                try
                {
                    if (!_Legacy == true)
                    {
                        if (!_NickName.Contains("@"))
                        {
                            throw new Exception("-NickName must contain EmailDomain, user@domain.com");
                        }
                            var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                            var _Xml = _DgcGoogleAppsService.CreateUserAlias(_ID, _UserService, _NickName);
                            var _UserAliasEntry = _DgcGoogleAppsService.CreateUserAliasEntry(_Xml);


                            WriteObject(_UserAliasEntry);
                    }
                    else
                    {
                        var _Entry = _UserService.CreateNickname(_ID, _NickName);
                        WriteObject(_Entry);
                    }

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

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Legacy
            {
                //get { return null; }
                set { _Legacy = value; }
            }
            private bool _Legacy;


            #endregion Parameters


            protected override void ProcessRecord()
            {
                try
                {
                    if (!_Legacy == true)
                    {
                        if(!_NickName.Contains("@"))
                        {
                            throw new Exception("-NickName must contain EmailDomain, user@domain.com");
                        }

                        var _DgcGoogleAppsService = new Dgc.GoogleAppService();
                        var _Xml = _DgcGoogleAppsService.RemoveUserAlias(_UserService, _NickName);


                        WriteObject(_NickName);

                    }
                    else
                    {
                        _UserService.DeleteNickname(_NickName);
                        WriteObject(_NickName);
                    }
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
