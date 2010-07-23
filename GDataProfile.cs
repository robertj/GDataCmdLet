using System;
using System.Diagnostics;
using System.Management.Automation;
using System.ComponentModel;
using System.Net;
using System.Text;
using System.Linq;
using System.IO;
using System.Web;
using System.Xml;
using System.Xml.Linq;


namespace Microsoft.PowerShell.GData
{

    public class Profile
    {

        #region New-GDataProfileService

        [Cmdlet(VerbsCommon.New, "GDataProfileService")]
        public class NewGDataProfileService : Cmdlet
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

                try
                {
                    var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                    var _Entry = _DgcGoogleProfileService.GetAuthToken(_AdminUser, _AdminPassword);

                    WriteObject(_Entry);

                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }


        }

        #endregion New-GDataProfileService

        #region Get-GDataProfile

        [Cmdlet(VerbsCommon.Get, "GDataProfile")]
        public class GetGDataProfile : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public Dgc.GoogleProfileService.ProfileService ProfileService
            {
                get { return null; }
                set { _ProfileService = value; }
            }
            private Dgc.GoogleProfileService.ProfileService _ProfileService;

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
                    var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                    var _Xml = _DgcGoogleProfileService.GetProfile(_ProfileService, _ID);

                    XmlDocument _XmlDoc = new XmlDocument();
                    _XmlDoc.InnerXml = _Xml;
                    XmlElement _Entry = _XmlDoc.DocumentElement;

                    WriteObject(_Entry);
                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

        #endregion Get-GDataProfile

        #region Set-GDataProfile

        [Cmdlet(VerbsCommon.Set, "GDataProfile")]
        public class SetGDataProfile : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public Dgc.GoogleProfileService.ProfileService ProfileService
            {
                get { return null; }
                set { _ProfileService = value; }
            }
            private Dgc.GoogleProfileService.ProfileService _ProfileService;

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
            public string PostalAddress
            {
                get { return null; }
                set { _PostalAddress = value; }
            }
            private string _PostalAddress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string HomePostalAddress
            {
                get { return null; }
                set { _HomePostalAddress = value; }
            }
            private string _HomePostalAddress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { _PhoneNumber = value; }
            }
            private string _PhoneNumber;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string MobilePhoneNumber
            {
                get { return null; }
                set { _MobilePhoneNumber = value; }
            }
            private string _MobilePhoneNumber;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string OtherPhoneNumber
            {
                get { return null; }
                set { _OtherPhoneNumber = value; }
            }
            private string _OtherPhoneNumber;

            #endregion Parameters

            protected override void ProcessRecord()
            {

                try
                {

                    
                    var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                    var _Xml = _DgcGoogleProfileService.SetProfile(_ProfileService, _ID, _PostalAddress, _PhoneNumber, _MobilePhoneNumber, _OtherPhoneNumber, _HomePostalAddress);
                    //var _Xml = _DgcGoogleProfileService.SetProfile(_ProfileService, ID, _PostalAddress, _PhoneNumber, _MobilePhoneNumber, _OtherPhoneNumber, _HomePostalAddress);

                    XmlDocument _XmlDoc = new XmlDocument();
                    _XmlDoc.InnerXml = _Xml;
                    XmlElement _Entry = _XmlDoc.DocumentElement;

                    WriteObject(_Entry);
                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }

    }

        #endregion Set-GdataProfile
}
