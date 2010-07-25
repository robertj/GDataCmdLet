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

    public class Resource
    {

        #region New-GDataResourceService
        /*
        [Cmdlet(VerbsCommon.New, "GDataResourceService")]
        public class NewGDataResourceService : Cmdlet
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
                    var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                    var _Entry = _DgcGoogleResourceService.GetAuthToken(_AdminUser, _AdminPassword);

                    WriteObject(_Entry);

                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);    
                }
            }


        }
        */
        #endregion New-GDataResourceService
        
        #region New-GDataResource

        [Cmdlet(VerbsCommon.New, "GDataResource")]
        public class NewGDataResource : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ResourceService = value; }
            }
            private GDataTypes.GDataService _ResourceService;

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
            public string Type
            {
                get { return null; }
                set { _Type = value; }
            }
            private string _Type;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { _Description = value; }
            }
            private string _Description;

            #endregion Parameters

            protected override void ProcessRecord()
            {

                try
                {
                    var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                    var _Xml = _DgcGoogleResourceService.NewResource(_ResourceService.ResourceService, _ID, _Type, _Description);
                    var _ResourceEntrys = _DgcGoogleResourceService.CreateResourceEntrys(_Xml, _ResourceService.ResourceService);
                    /*
                    XmlDocument _XmlDoc = new XmlDocument();
                    _XmlDoc.InnerXml = _Xml;
                    XmlElement _Entry = _XmlDoc.DocumentElement;
                    */
                    WriteObject(_ResourceEntrys,true);
                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }
        #endregion New-GDataResource

        #region Get-GDataResource

        [Cmdlet(VerbsCommon.Get, "GDataResource")]
        public class GetGDataResource : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ResourceService = value; }
            }
            private GDataTypes.GDataService _ResourceService;

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

                try
                {
                    if (_ID != null)
                    {
                        var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                        var _Xml = _DgcGoogleResourceService.RetriveResource(_ResourceService.ResourceService, _ID);
                        var _ResourceEntrys = _DgcGoogleResourceService.CreateResourceEntrys(_Xml, _ResourceService.ResourceService);
                        /*
                        XmlDocument _XmlDoc = new XmlDocument();
                        _XmlDoc.InnerXml = _Xml;
                        XmlElement _Entry = _XmlDoc.DocumentElement;
                        */
                        WriteObject(_ResourceEntrys,true);
                    }
                    else
                    {
                        var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                        var _Xml = _DgcGoogleResourceService.RetriveAllResources(_ResourceService.ResourceService);
                        var _ResourceEntrys = _DgcGoogleResourceService.CreateResourceEntrys(_Xml, _ResourceService.ResourceService);
                        /*
                        XmlDocument _XmlDoc = new XmlDocument();
                        _XmlDoc.InnerXml = _Xml;
                        XmlElement _Entry = _XmlDoc.DocumentElement;
                        */
                        WriteObject(_ResourceEntrys,true);
                    }
                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }
        }
        
        #endregion Get-GDataResource
        
        #region Remove-GDataResource

        [Cmdlet(VerbsCommon.Remove, "GDataResource")]
        public class RemoveGDataResource : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ResourceService = value; }
            }
            private GDataTypes.GDataService _ResourceService;

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

                    var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                    var _Xml = _DgcGoogleResourceService.RemoveResources(_ResourceService.ResourceService, _ID);

                    WriteObject(_ID);

                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }
        
        #endregion Remove-GDataResource

        #region Set-GDataResource

        [Cmdlet(VerbsCommon.Set, "GDataResource")]
        public class SetGDataResource : Cmdlet
        {
            #region Parameters

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public GDataTypes.GDataService Service
            {
                get { return null; }
                set { _ResourceService = value; }
            }
            private GDataTypes.GDataService _ResourceService;

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
            public string Type
            {
                get { return null; }
                set { _Type = value; }
            }
            private string _Type;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { _Description = value; }
            }
            private string _Description;

            #endregion Parameters

            protected override void ProcessRecord()
            {

                try
                {
                    var _DgcGoogleResourceService = new Dgc.GoogleResourceService();
                    _DgcGoogleResourceService.SetResource(_ResourceService.ResourceService, _ID, _Type, _Description);

                    var _Xml = _DgcGoogleResourceService.RetriveResource(_ResourceService.ResourceService, _ID);
                    var _ResourceEntrys = _DgcGoogleResourceService.CreateResourceEntrys(_Xml, _ResourceService.ResourceService);
                    /*
                    XmlDocument _XmlDoc = new XmlDocument();
                    _XmlDoc.InnerXml = _Xml;
                    XmlElement _Entry = _XmlDoc.DocumentElement;
                    */
                    WriteObject(_ResourceEntrys);
                }
                catch (WebException _Exception)
                {
                    WriteObject(_Exception);
                }
            }

        }
        #endregion New-GDataResource
    
    
    }

}
