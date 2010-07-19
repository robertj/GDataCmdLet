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
using System.Linq;

namespace Microsoft.PowerShell.GData
{

    public class MailSetting
    {

        #region Pop3ActionDelete


        private class Pop3ActionDelete
        {
            public string GetAction(bool Pop3Action)
            {
                if (Pop3Action == true)
                {
                    string _Pop3Action = "DELETE";
                    return (_Pop3Action);
                }
                else
                {
                    string _Pop3Action = "KEEP";
                    return (_Pop3Action);
                }

            }
        }

        #endregion Pop3ActionDelete
        
        #region New-GDataGDataMailSettingService

        [Cmdlet(VerbsCommon.New, "GDataMailSettingService")]
        public class NewGDatMmailSettingService : Cmdlet
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
                    var _GoogleMailSettingsService = new GoogleMailSettingsService(_Domain, "GMailSettingsService");
                    _GoogleMailSettingsService.setUserCredentials(_AdminUser, _AdminPassword);
                    WriteObject(_GoogleMailSettingsService);
                }
                catch (Exception _Exception)
                {
                    WriteObject(_Exception);
                }
            }


        }

        #endregion New-GDataGDataMailSettingService

        #region Set-GDataGDataMailSetting

        [Cmdlet(VerbsCommon.Set, "GDataMailSetting")]
        public class SetGDataMailSetting : Cmdlet
        {

            #region Parameters


            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GoogleMailSettingsService MailSettingsService
            {
                get { return null; }
                set { _MailSettingsService = value; }
            }
            private GoogleMailSettingsService _MailSettingsService;

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
            public string SenderAdress
            {
                get { return null; }
                set { _SenderAdress = value; }
            }
            private string _SenderAdress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter IsDefault
            {
                set { _IsDefault = value; }
            }
            private bool _IsDefault;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                get { return null; }
                set { _Name  = value; }
            }
            private string _Name ;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter EnablePop3
            {
                //get { return null; }
                set { _EnablePop3 = value; }
            }
            private bool _EnablePop3;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter DisablePop3
            {
                //get { return null; }
                set { _DisablePop3 = value; }
            }
            private bool _DisablePop3;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter EnableImap
            {
                //get { return null; }
                set { _EnableImap = value; }
            }
            private bool _EnableImap;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter DisableImap
            {
                //get { return null; }
                set { _DisableImap = value; }
            }
            private bool _DisableImap;
            
            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Pop3ActionDelete
            {
                //get { return null; }
                set { _Pop3ActionDelete = value; }
            }
            private bool _Pop3ActionDelete;

            #endregion Parameters

            //private string Default;

            protected override void ProcessRecord()
            {

                if (_SenderAdress != null)
                {
                    
                    if (_Name  == null)
                    {
                        throw new ArgumentException("Parameter Name is null"); 
                    }

                    var _StringIsDefault = _IsDefault.ToString();


                    try
                    {
                        var _Entry = _MailSettingsService.CreateSendAs(_ID, _Name, _SenderAdress, _SenderAdress, _StringIsDefault);
                        WriteObject(_Entry);
                    }
                    catch (Exception _Exception )
                    {
                        WriteObject( _Exception );
                    }
                }

                var _GetPop3Action = new Pop3ActionDelete();
                var _Pop3Action = _GetPop3Action.GetAction(_Pop3ActionDelete);
                
                if (_EnablePop3 == true)
                {
                    try
                    {
                        var _Entry = _MailSettingsService.UpdatePop(_ID, "True", "ALL_MAIL", _Pop3Action);
                        WriteObject(_Entry);
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
                    }
                }
                if (_DisablePop3 == true)
                {
                    try
                    {
                        var _Entry = _MailSettingsService.UpdatePop(_ID, "False", "ALL_MAIL", _Pop3Action);
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
                    }
                }
                if (_EnableImap == true)
                {
                    try
                    {
                        var _Entry = _MailSettingsService.UpdateImap(_ID, "True");
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
                    }
                }
                if (_DisableImap == true)
                {
                    try
                    {
                        var _Entry = _MailSettingsService.UpdateImap(_ID, "False");
                    }
                    catch (Exception _Exception)
                    {
                        WriteObject(_Exception);
                    }
                }
            }
        }

        #endregion Set-GDataMailSetting
  
      

        
    }


}
