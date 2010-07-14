using System;
using System.Diagnostics;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Google.Contacts;
using Google.GData.Client;
using Google.GData.Contacts;
using Google.GData.Extensions;
using System.Collections.Generic;


namespace Microsoft.PowerShell.GData
{

 
    #region PowerShell snap-in

    [RunInstaller(true)]
    public class Ps : PSSnapIn

    {
        public Ps()
            : base()
        {
        }

        public override string Name
        {
            get
            {
                return "GData";
            }
        }

        public override string Vendor
        {
            get
            {
                return "plan-tre.net";
            }
        }

        public override string VendorResource
        {
            get
            {
                return "PSGData,plan-tre,net";
            }
        }

        public override string Description
        {
            get
            {
                return "This is a PowerShell snap-in that includes the GData cmdlet's.";
            }
        }
 
    }

    #endregion PowerShell snap-in


}
