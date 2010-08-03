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

        #region Get-GDataProfile

        [Cmdlet(VerbsCommon.Get, "GDataProfile")]
        public class GetGDataProfile : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            #endregion Parameters

            private string nextPage;
            private string parseXML;
            private Dgc.GoogleProfileService dgcGoogleProfileService = new Dgc.GoogleProfileService();
            protected override void ProcessRecord()
            {

                try
                {
                    if (id != null)
                    {
                        var _xml = dgcGoogleProfileService.GetProfile(service.ProfileService, id);

                        var _profileEntry = dgcGoogleProfileService.CreateProfileEntry(_xml, id, service.ProfileService);
                        WriteObject(_profileEntry);

                    }
                    else
                    {
                        nextPage = "";
                        var _xml = dgcGoogleProfileService.GetProfiles(service.ProfileService, nextPage);

                        var parseXML = new GDataTypes.ParseXML(_xml.ToString());
                        var _profileEntrys = dgcGoogleProfileService.CreateProfileEntrys(_xml, service.ProfileService);

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

                                _xml = dgcGoogleProfileService.GetProfiles(service.ProfileService, nextPage);
                                parseXML = new GDataTypes.ParseXML(_xml.ToString());
                                _profileEntrys = dgcGoogleProfileService.AppendProfileEntrys(_xml, service.ProfileService, _profileEntrys);
                            }
                        }

                        WriteObject(_profileEntrys, true);
                    }

                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { service = value; }
            }
            private GDataTypes.GDataService service;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string ID
            {
                get { return null; }
                set { id = value; }
            }
            private string id;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PostalAddress
            {
                get { return null; }
                set { postalAddress = value; }
            }
            private string postalAddress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string HomePostalAddress
            {
                get { return null; }
                set { homePostalAddress = value; }
            }
            private string homePostalAddress;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string PhoneNumber
            {
                get { return null; }
                set { phoneNumber = value; }
            }
            private string phoneNumber;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string MobilePhoneNumber
            {
                get { return null; }
                set { mobilePhoneNumber = value; }
            }
            private string mobilePhoneNumber;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string OtherPhoneNumber
            {
                get { return null; }
                set { otherPhoneNumber = value; }
            }
            private string otherPhoneNumber;

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string HomePhoneNumber
            {
                get { return null; }
                set { homePhoneNumber = value; }
            }
            private string homePhoneNumber;

            #endregion Parameters

            private Dgc.GoogleProfileService dgcGoogleProfileService = new Dgc.GoogleProfileService();
            protected override void ProcessRecord()
            {
                try
                {
                    var _xml = dgcGoogleProfileService.SetProfile(service.ProfileService, id, postalAddress, phoneNumber, mobilePhoneNumber, otherPhoneNumber, homePostalAddress, homePhoneNumber);
                    var _profileEntry = dgcGoogleProfileService.CreateProfileEntry(_xml, id, service.ProfileService);

                    WriteObject(_profileEntry);
                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
                }
            }

        }
        #endregion Set-GdataProfile

    }
}
