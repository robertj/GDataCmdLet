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
        /*
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
        */
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ProfileService = value; }
            }
            private GDataTypes.GDataService _ProfileService;

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

            private string NextPage;
            private string ParseXML;
            protected override void ProcessRecord()
            {

                try
                {
                    if (_ID != null)
                    {
                        var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                        var _Xml = _DgcGoogleProfileService.GetProfile(_ProfileService.ProfileService, _ID);

                        var _ProfileEntry = _DgcGoogleProfileService.CreateProfileEntry(_Xml, _ID, _ProfileService.ProfileService);
                        WriteObject(_ProfileEntry);

                    }
                    else
                    {
                        
                        var _DgcGoogleProfileService = new Dgc.GoogleProfileService();


                        NextPage = "";
                        var _Xml = _DgcGoogleProfileService.GetProfiles(_ProfileService.ProfileService, NextPage);

                        var ParseXML = new GDataTypes.ParseXML(_Xml.ToString());
                        var _ProfileEntrys = _DgcGoogleProfileService.CreateProfileEntrys(_Xml, _ProfileService.ProfileService);

                        while (NextPage != null)
                        {
                            NextPage = null;
                            foreach (var _Elements in ParseXML.ListFormat)
                            {
                                foreach (var _Attribute in _Elements.at)
                                {
                                    if (_Attribute.Value == "next")
                                    {
                                        //Console.WriteLine("Next/n");
                                        NextPage = _Attribute.NextAttribute.NextAttribute.Value;
                                    }
                                }
                            }
                            if (NextPage != null)
                            {

                                _Xml = _DgcGoogleProfileService.GetProfiles(_ProfileService.ProfileService, NextPage);
                                ParseXML = new GDataTypes.ParseXML(_Xml.ToString());
                                _ProfileEntrys = _DgcGoogleProfileService.AppendProfileEntrys(_Xml, _ProfileService.ProfileService, _ProfileEntrys);
                            }
                        }

                        WriteObject(_ProfileEntrys,true);
                    }

                    #region comment
                    /*

                    var _ParesdXml = new GDataTypes.ParseXML(_Xml);


                    var _ProfileEntry = new GDataTypes.ProfileEntry();

                    _ProfileEntry.User = _ID;
                    _ProfileEntry.Domain = _ProfileService.Domain;

                    foreach (var _Entry in _ParesdXml.ListFormat)
                    {
                        if (_Entry.name == "{http://schemas.google.com/g/2005}structuredPostalAddress")
                        {
                            foreach (var _Attribute in _Entry.at)
                            {
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#home")
                                {
                                    _ProfileEntry.HomePostalAddress = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#work")
                                {
                                    _ProfileEntry.PostalAddress = _Entry.value;
                                }
                            }
                        }
                        if (_Entry.name == "{http://schemas.google.com/g/2005}phoneNumber")
                        {
                            foreach (var _Attribute in _Entry.at)
                            {
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#other")
                                {
                                    _ProfileEntry.OtherPhoneNumber = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#work")
                                {
                                    _ProfileEntry.PhoneNumber = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#mobile")
                                {
                                    _ProfileEntry.MobilePhoneNumber = _Entry.value;
                                }
                            }
                        }
                    }
                    */
                    #endregion comment
                    
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
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ProfileService = value; }
            }
            private GDataTypes.GDataService _ProfileService;

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

            [Parameter(
            Mandatory = false
            )]
            [ValidateNotNullOrEmpty]
            public string HomePhoneNumber
            {
                get { return null; }
                set { _HomePhoneNumber = value; }
            }
            private string _HomePhoneNumber;

            #endregion Parameters

            protected override void ProcessRecord()
            {

                try
                {

                    
                    var _DgcGoogleProfileService = new Dgc.GoogleProfileService();
                    var _Xml = _DgcGoogleProfileService.SetProfile(_ProfileService.ProfileService, _ID, _PostalAddress, _PhoneNumber, _MobilePhoneNumber, _OtherPhoneNumber, _HomePostalAddress, _HomePhoneNumber);

                    var _ProfileEntry = _DgcGoogleProfileService.CreateProfileEntry(_Xml, _ID, _ProfileService.ProfileService);

                    #region comment
                    /*
                    var _ParesdXml = new GDataTypes.ParseXML(_Xml);

                   
                    var _ProfileEntry = new GDataTypes.ProfileEntry();

                    _ProfileEntry.User = _ID;
                    _ProfileEntry.Domain = _ProfileService.Domain;

                    foreach (var _Entry in _ParesdXml.ListFormat)
                    {
                        if (_Entry.name == "{http://schemas.google.com/g/2005}structuredPostalAddress")
                        {
                            foreach (var _Attribute in _Entry.at)
                            {
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#home")
                                {
                                    _ProfileEntry.HomePostalAddress = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#work")
                                {
                                    _ProfileEntry.PostalAddress = _Entry.value;
                                }
                            }
                        }
                        if (_Entry.name == "{http://schemas.google.com/g/2005}phoneNumber")
                        {
                            foreach (var _Attribute in _Entry.at)
                            {
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#other")
                                {
                                    _ProfileEntry.OtherPhoneNumber = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#work")
                                {
                                    _ProfileEntry.PhoneNumber = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#mobile")
                                {
                                    _ProfileEntry.MobilePhoneNumber = _Entry.value;
                                }
                                if (_Attribute.Value == "http://schemas.google.com/g/2005#mobile")
                                {
                                    _ProfileEntry.MobilePhoneNumber = _Entry.value;
                                }
                            }
                        }
                    }
                    */
                    #endregion comment
                    WriteObject(_ProfileEntry);
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
