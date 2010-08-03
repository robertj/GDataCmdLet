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
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.PowerShell.GData
{

    public class MailSetting
    {
        #region Pop3ActionDelete

        private class Pop3ActionDelete
        {
            public string Action;
            public Pop3ActionDelete(bool pop3Action)
            {
                if (pop3Action == true)
                {
                    Action = "DELETE";
                }
                else
                {
                    Action = "KEEP";
                }

            }
        }

        #endregion Pop3ActionDelete
        
        #region Set-GDataGDataMailSetting

        [Cmdlet(VerbsCommon.Set, "GDataMailSetting")]
        public class SetGDataMailSetting : Cmdlet
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
            public string SenderAdress
            {
                set { senderAdress = value; }
            }
            private string senderAdress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter IsDefault
            {
                set { isDefault = value; }
            }
            private bool isDefault;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string Name
            {
                set { name  = value; }
            }
            private string name ;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter EnablePop3
            {
                set { enablePop3 = value; }
            }
            private bool enablePop3;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter DisablePop3
            {
                set { disablePop3 = value; }
            }
            private bool disablePop3;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter EnableImap
            {
                set { enableImap = value; }
            }
            private bool enableImap;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter DisableImap
            {
                set { disableImap = value; }
            }
            private bool disableImap;
            
            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public SwitchParameter Pop3ActionDelete
            {
                set { pop3ActionDelete = value; }
            }
            private bool pop3ActionDelete;

            #endregion Parameters

            private Dgc.GoogleAppService dgcGoogleAppsService = new Dgc.GoogleAppService();
            protected override void ProcessRecord()
            {
                if (senderAdress != null)
                {
                    
                    if (name  == null)
                    {
                        throw new ArgumentException("Parameter Name is null"); 
                    }

                    var _stringIsDefault = isDefault.ToString();


                    try
                    {
                        var _entry = dgcGoogleAppsService.CreateSenderAddressEntry(service.GoogleMailSettingsService.CreateSendAs(id, name, senderAdress, senderAdress, _stringIsDefault));
                        WriteObject(_entry);
                    }
                    catch (Exception _exception   )
                    {
                        WriteObject( _exception   );
                    }
                }

                var _pop3Action = new Pop3ActionDelete(pop3ActionDelete).Action;

                if (enablePop3 == true)
                {
                    try
                    {
                        var _entry = dgcGoogleAppsService.CreatePop3Entry(service.GoogleMailSettingsService.UpdatePop(id, "True", "ALL_MAIL", _pop3Action));
                        WriteObject(_entry);
                    }
                    catch (Exception _exception  )
                    {
                        WriteObject(_exception  );
                    }
                }
                if (disablePop3 == true)
                {
                    try
                    {
                        var _entry = dgcGoogleAppsService.CreatePop3Entry(service.GoogleMailSettingsService.UpdatePop(id, "False", "ALL_MAIL", _pop3Action));
                        WriteObject(_entry);
                    }
                    catch (Exception _exception  )
                    {
                        WriteObject(_exception  );
                    }
                }
                if (enableImap == true)
                {
                    try
                    {
                        var _entry = dgcGoogleAppsService.CreateIMapEntry(service.GoogleMailSettingsService.UpdateImap(id, "True"));
                        WriteObject(_entry);
                    }
                    catch (Exception _exception  )
                    {
                        WriteObject(_exception  );
                    }
                }
                if (disableImap == true)
                {
                    try
                    {
                        var _entry = service.GoogleMailSettingsService.UpdateImap(id, "False");
                        var _imapEntry = dgcGoogleAppsService.CreateIMapEntry(_entry);
                        WriteObject(_imapEntry);
                    }
                    catch (Exception _exception  )
                    {
                        WriteObject(_exception  );
                    }
                }
            }
        }

        #endregion Set-GDataMailSetting
    }
}
