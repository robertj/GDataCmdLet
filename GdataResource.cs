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
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Type
            {
                get { return null; }
                set { type = value; }
            }
            private string type;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { description = value; }
            }
            private string description;

            #endregion Parameters
            private Dgc.GoogleResourceService dgcGoogleResourceService = new Dgc.GoogleResourceService();
            protected override void ProcessRecord()
            {

                try
                {
                    var _xml = dgcGoogleResourceService.NewResource(service.ResourceService, id, type, description);
                    var _resourceEntrys = dgcGoogleResourceService.CreateResourceEntrys(_xml, service.ResourceService);

                    WriteObject(_resourceEntrys, true);
                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
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

            private Dgc.GoogleResourceService dgcGoogleResourceService = new Dgc.GoogleResourceService();
            protected override void ProcessRecord()
            {

                try
                {
                    if (id != null)
                    {
                        var _xml = dgcGoogleResourceService.RetriveResource(service.ResourceService, id);
                        var _resourceEntrys = dgcGoogleResourceService.CreateResourceEntrys(_xml, service.ResourceService);

                        WriteObject(_resourceEntrys, true);
                    }
                    else
                    {
                        var _xml = dgcGoogleResourceService.RetriveAllResources(service.ResourceService);
                        var _resourceEntrys = dgcGoogleResourceService.CreateResourceEntrys(_xml, service.ResourceService);

                        WriteObject(_resourceEntrys, true);
                    }
                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
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

            #endregion Parameters

            private Dgc.GoogleResourceService dgcGoogleResourceService = new Dgc.GoogleResourceService();
            protected override void ProcessRecord()
            {

                try
                {
                    var _xml = dgcGoogleResourceService.RemoveResources(service.ResourceService, id);

                    WriteObject(id);

                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
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
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Type
            {
                get { return null; }
                set { type = value; }
            }
            private string type;

            [Parameter(
            Mandatory = true
            )]
            [ValidateNotNullOrEmpty]
            public string Description
            {
                get { return null; }
                set { description = value; }
            }
            private string description;

            #endregion Parameters

            private Dgc.GoogleResourceService dgcGoogleResourceService = new Dgc.GoogleResourceService();
            protected override void ProcessRecord()
            {

                try
                {
                    dgcGoogleResourceService.SetResource(service.ResourceService, id, type, description);

                    var _xml = dgcGoogleResourceService.RetriveResource(service.ResourceService, id);
                    var _resourceEntrys = dgcGoogleResourceService.CreateResourceEntrys(_xml, service.ResourceService);

                    WriteObject(_resourceEntrys);
                }
                catch (WebException _exception)
                {
                    WriteObject(_exception);
                }
            }

        }
        #endregion New-GDataResource

    }
}
