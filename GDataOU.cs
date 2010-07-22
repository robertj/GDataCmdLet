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
    public class OrganizationalUnit
    {
        
        #region New-GDataOUService

        [Cmdlet(VerbsCommon.New, "GDataOUService")]
        public class NewGDataOUService : Cmdlet
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

        #region Get-GDataOU

        [Cmdlet(VerbsCommon.Get, "GDataOU")]
        public class GetGDataUser : Cmdlet
        {
            #region Parameters

            [Parameter(
                      Mandatory = true
                      )]
            [ValidateNotNullOrEmpty]
            public AppsService OUService
            {
                get { return null; }
                set { _OUService = value; }
            }
            private AppsService _OUService;

            #endregion Parameters


            protected override void ProcessRecord()
            {
                var OUService = new Dgc.GoogleAppService();
                var _Xml = OUService.RetrievAllOUs(_OUService);
                
                XmlDocument _XmlDoc = new XmlDocument();
                _XmlDoc.InnerXml = _Xml;
                XmlElement _Entry = _XmlDoc.DocumentElement;

                WriteObject(_Entry);
            }

        }

        #endregion Get-GDataUser
        
    }

}
